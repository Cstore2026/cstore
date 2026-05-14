const RAW_DATA = [];
const STORE_KEY = 'cstore_dashboard_pro_settings_v1';
const DEFAULT_PASSWORD = '17512';
const state = {
  orders: RAW_DATA, search:"", activeTab:"overview", showOrders:false,
  filters:{years:[], months:[], days:[], branches:[], statuses:[], preparers:[], drivers:[], areas:[], platforms:[], payments:[]},
  settings:{password:DEFAULT_PASSWORD, viewerPassword:'12345', salesTargets:{}, bands:{branches:{}, preparers:{}, drivers:{}, customers:{}, platforms:{}, areas:{}}, alertThresholds:{cancelRate:10}},
  sessionRole:'', lang:(localStorage.getItem('cstore_lang')||''), kpi:{month:'', rows:[], sourceName:'', sourceType:''}, pendingTargets:{}, dailySelectedDate:'', periodFrom:'', periodTo:'', branchPeakSelected:''
};


const KPI_PREPARERS_MONTH = '';
const KPI_PREPARERS_ROWS = [];

function normalizeKpiName(name){
  return String(name||'')
    .trim()
    .replace(/[أإآ]/g,'ا')
    .replace(/ى/g,'ي')
    .replace(/ة/g,'ه')
    .replace(/\s+/g,' ')
    .toLowerCase();
}

function getBestKpiPreparer(){
  const rows = getActiveKpiRows();
  return rows.slice().sort((a,b)=>(b.pct||0)-(a.pct||0) || (b.final||0)-(a.final||0))[0] || null;
}

function findKpiForPreparer(name){
  const target = normalizeKpiName(name);
  const rows = getActiveKpiRows();
  if(!target) return null;
  const exact = rows.find(r => normalizeKpiName(r.name) === target);
  if(exact) return exact;
  const partial = rows.find(r => {
    const n = normalizeKpiName(r.name);
    return n.includes(target) || target.includes(n);
  });
  return partial || null;
}



const KPI_STORE_KEY = 'cstore_dashboard_kpi_live_v1';

function loadKpiStore(){
  try{
    const saved = JSON.parse(localStorage.getItem(KPI_STORE_KEY) || '{}');
    if(saved && Array.isArray(saved.rows)){
      state.kpi = {
        month: String(saved.month || ''),
        rows: saved.rows.map(r=>({
          name: String(r.name || '').trim(),
          pct: Number(r.pct || 0),
          final: Number(r.final || 0)
        })).filter(r=>r.name),
        sourceName: String(saved.sourceName || ''),
        sourceType: String(saved.sourceType || '')
      };
    }
  }catch(e){}
}
function saveKpiStore(){
  localStorage.setItem(KPI_STORE_KEY, JSON.stringify(state.kpi || {}));
}
function clearKpiStore(){
  state.kpi = {month:'', rows:[], sourceName:'', sourceType:''};
  saveKpiStore();
}

function getActiveKpiRows(){
  return (state.kpi && Array.isArray(state.kpi.rows) && state.kpi.rows.length) ? state.kpi.rows : KPI_PREPARERS_ROWS;
}
function getActiveKpiMonth(){
  return (state.kpi && state.kpi.month) ? state.kpi.month : KPI_PREPARERS_MONTH;
}
function getActiveKpiSourceName(){
  return (state.kpi && state.kpi.sourceName) ? state.kpi.sourceName : 'لا يوجد';
}

function extractPctValue(v){
  if(v === null || v === undefined || v === '') return 0;
  const s = String(v).replace('%','').replace(/,/g,'').trim();
  const n = Number(s);
  if(!Number.isFinite(n)) return 0;
  return n <= 1 ? n * 100 : n;
}

function parseKpiWorkbook(workbook){
  const sheet = workbook.Sheets['Final'];
  if(!sheet) throw new Error('لم يتم العثور على تبويب Final داخل ملف KPI');
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
  const monthRows = [];
  rows.forEach((row, idx)=>{
    const first = String((row && row[0]) || '').trim();
    if(first.startsWith('شهر ')){
      monthRows.push({row: idx, month: first.replace(/^شهر\s+/, '').trim()});
    }
  });
  if(!monthRows.length) throw new Error('لم يتم العثور على أي شهر داخل تبويب Final');

  const latest = monthRows[monthRows.length - 1];
  let headerRowIndex = -1;
  for(let i = latest.row; i < Math.min(rows.length, latest.row + 8); i++){
    const line = (rows[i] || []).map(x=>String(x || '').trim());
    if(line.some(x => x.includes('الاسم')) && line.some(x => x.includes('النسبة'))){
      headerRowIndex = i;
      break;
    }
  }
  if(headerRowIndex === -1) throw new Error('تعذر تحديد عناوين جدول KPI داخل شهر ' + latest.month);

  const header = (rows[headerRowIndex] || []).map(x=>String(x || '').trim());
  const nameCol = header.findIndex(x=>x.includes('الاسم'));
  const pctCol = header.findIndex(x=>x.includes('النسبة'));
  const finalCol = header.findIndex(x=>x.includes('الدرجة النهائية'));
  if(nameCol === -1 || pctCol === -1) throw new Error('أعمدة الاسم أو النسبة غير موجودة داخل تبويب Final');

  const out = [];
  for(let i = headerRowIndex + 1; i < rows.length; i++){
    const line = rows[i] || [];
    const first = String(line[0] || '').trim();
    if(first.startsWith('شهر ')) break;
    const name = String(line[nameCol] || '').trim();
    const pct = extractPctValue(line[pctCol]);
    const final = finalCol >= 0 ? Number(line[finalCol] || 0) : 0;
    if(!name) continue;
    if(name === 'الدرجة' || name === 'الاسم' || name === 'الاسم / الشروط') continue;
    if(!pct) continue;
    out.push({name, pct, final});
  }

  const cleaned = out.filter(r=>r.name).sort((a,b)=>(b.pct||0)-(a.pct||0) || (b.final||0)-(a.final||0));
  if(!cleaned.length) throw new Error('تم العثور على الشهر لكن لا توجد صفوف KPI صالحة');
  return {month: latest.month, rows: cleaned};
}

function normalizeGoogleSheetUrl(url){
  const s = String(url || '').trim();
  const m = s.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if(m) return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;
  return s;
}

async function importKpiFromFile(){
  const file = document.getElementById('kpiFileInput')?.files?.[0];
  const kpiMsg = document.getElementById('kpiMsg');
  try{
    if(kpiMsg) kpiMsg.textContent = '';
    if(!file){
      const text = (state.lang==='en' ? 'Choose KPI file first' : 'اختر ملف KPI أولاً');
      document.getElementById('msg').textContent = text;
      if(kpiMsg) kpiMsg.textContent = text;
      return;
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const parsed = parseKpiWorkbook(wb);
    state.kpi = {month: parsed.month, rows: parsed.rows, sourceName: file.name, sourceType:'file'};
    saveKpiStore();
    const text = `تم استيراد KPI بنجاح من الملف: ${file.name} — الشهر: ${parsed.month} — عدد الصفوف: ${parsed.rows.length}`;
    document.getElementById('msg').textContent = text;
    if(kpiMsg) kpiMsg.textContent = text;
    render();
applyLanguage();
  }catch(err){
    const text = (state.lang==='en' ? 'Failed to read KPI file: ' : 'فشل قراءة ملف KPI: ') + (err?.message || err);
    document.getElementById('msg').textContent = text;
    if(kpiMsg) kpiMsg.textContent = text;
  }
}

async function importKpiFromLink(){
  document.getElementById('msg').textContent = 'قراءة KPI من الرابط متوقفة حاليًا. استخدم Excel فقط.';
}

function getBestKpiPreparer(){
  const rows = getActiveKpiRows();
  return rows.slice().sort((a,b)=>(b.pct||0)-(a.pct||0) || (b.final||0)-(a.final||0))[0] || null;
}


function loadSettings(){try{const saved=JSON.parse(localStorage.getItem(STORE_KEY)||'{}'); if(saved.password) state.settings.password=saved.password; if(saved.viewerPassword) state.settings.viewerPassword=saved.viewerPassword; if(saved.salesTargets) state.settings.salesTargets=saved.salesTargets; if(saved.monthlySalesTargets) state.settings.monthlySalesTargets=saved.monthlySalesTargets; if(saved.alertThresholds) state.settings.alertThresholds=saved.alertThresholds;}catch(e){}}
function saveSettings(){localStorage.setItem(STORE_KEY, JSON.stringify(state.settings))}
loadSettings();
loadKpiStore();



const I18N_EN = {
  "العربية":"Arabic",
  "English":"English",
  "اختيار اللغة":"Choose language",
  "Language / اللغة":"Language",
  "يناير":"January",
  "فبراير":"February",
  "مارس":"March",
  "أبريل":"April",
  "مايو":"May",
  "يونيو":"June",
  "يوليو":"July",
  "أغسطس":"August",
  "سبتمبر":"September",
  "أكتوبر":"October",
  "نوفمبر":"November",
  "ديسمبر":"December",
  "تم":"Done",
  "ملغي / مرتجع":"Cancelled / Returned",
  "دخول":"Login",
  "خروج":"Logout",
  "كلمة المرور":"Password",
  "لوحة تحكم توصيل سي ستور":"C Store Delivery Dashboard",
  "متابعة العمليات":"Operations Monitoring",
  "أداء الفروع":"Branch Performance",
  "تحليلات التوصيل":"Delivery Insights",
  "ابحث برقم الأوردر / العميل / الهاتف / المنطقة / الطيار...":"Search by order / customer / phone / area / driver...",
  "نظرة عامة":"Overview",
  "تقرير يومي منفصل":"Daily Report",
  "تحليل الفترات":"Period Analysis",
  "الرسوم البيانية":"Charts",
  "ذروة الفروع اليومية":"Daily Branch Peaks",
  "الفروع":"Branches",
  "الطيارين":"Drivers",
  "المنصات":"Platforms",
  "المناطق":"Areas",
  "العملاء":"Customers",
  "الإعدادات":"Settings",
  "إجمالي الطلبات":"Total Orders",
  "إجمالي المبيعات":"Total Sales",
  "متوسط الأوردر":"Avg Order",
  "الملغي / المرتجع":"Cancelled / Returned",
  "نسبة الإلغاء":"Cancellation Rate",
  "أعلى ساعة ضغط":"Peak Hour",
  "معدل الأداء العام":"Overall Performance",
  "Last Day Orders":"Last Day Orders",
  "أفضل محضّر":"Best Preparer",
  "أفضل طيار":"Best Driver",
  "أعلى عميل":"Top Customer",
  "أفضل فرع":"Top Branch",
  "عدد السجلات":"Records Count",
  "جدول الأوردرات التفصيلي":"Detailed Orders Table",
  "اختياري":"Optional",
  "إظهار جدول الأوردرات":"Show Orders Table",
  "إخفاء جدول الأوردرات":"Hide Orders Table",
  "السنة":"Year",
  "الشهر":"Month",
  "اليوم":"Day",
  "الفرع":"Branch",
  "المنطقة":"Area",
  "الحالة":"Status",
  "المحضّر":"Preparer",
  "الطيار":"Driver",
  "المنصة":"Platform",
  "الدفع":"Payment",
  "لوحة متابعة الإدارة":"Management Dashboard",
  "عرض فقط":"View Only",
  "ملخص جاهز للإدارة":"Executive Summary",
  "التارجت":"Target",
  "المبيعات":"Sales",
  "الإلغاء":"Cancellation",
  "نسبة التحقيق":"Achievement Rate",
  "الإجمالي الحالي":"Current Total",
  "مؤشر سريع":"Quick Indicator",
  "تحليل المناطق":"Areas Analysis",
  "تحليل العملاء":"Customers Analysis",
  "No Data مناطق لعرضها.":"No area data to display.",
  "غير محدد":"Not specified",
  "ربط KPI":"KPI Link",
  "ملخص KPI":"KPI Summary",
  "أفضل 5 محضّرين من KPI":"Top 5 KPI Preparers",
  "أفضل محضّر في الداشبورد":"Best Dashboard Preparer",
  "جدول KPI الكامل":"Full KPI Table",
  "المصدر الحالي":"Current Source",
  "الشهر الحالي":"Current Month",
  "لا يوجد":"None",
  "Viewer":"Viewer",
  "Admin":"Admin"
};
const I18N_AR = Object.fromEntries(Object.entries(I18N_EN).map(([ar,en])=>[en,ar]));

function translateTextValue(value){
  const key = String(value || '').trim();
  if(!key) return value;
  if(state.lang === 'en') return I18N_EN[key] || value;
  return I18N_AR[key] || value;
}


function tr(value){
  return translateTextValue(value);
}

function monthNameLocal(m){
  const ar = ["","يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"][Number(m)||0] || '';
  return state.lang === 'en' ? (I18N_EN[ar] || ar) : ar;
}

function statusLocal(v){
  return state.lang === 'en' ? (I18N_EN[v] || v) : v;
}

function translateNodeText(root){
  if(!root) return;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
  const nodes = [];
  while(walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach(node=>{
    const current = String(node.nodeValue || '');
    const trimmed = current.trim();
    if(!trimmed) return;
    const translated = translateTextValue(trimmed);
    if(translated !== trimmed){
      node.nodeValue = current.replace(trimmed, translated);
    }
  });

  root.querySelectorAll('[placeholder]').forEach(el=>{
    const p = el.getAttribute('placeholder');
    if(p){
      if(p.includes('ابحث') || p.includes('Search by')){
        el.setAttribute('placeholder', state.lang === 'en'
          ? 'Search by order / customer / phone / area / driver...'
          : 'ابحث برقم الأوردر / العميل / الهاتف / المنطقة / الطيار...');
      } else if(p.includes('كلمة المرور') || p.includes('Password')){
        el.setAttribute('placeholder', state.lang === 'en' ? 'Password' : 'كلمة المرور');
      } else {
        el.setAttribute('placeholder', translateTextValue(p));
      }
    }
  });
}

function applyLanguage(){
  const activeLang = state.lang === 'en' ? 'en' : 'ar';
  document.documentElement.lang = activeLang;
  document.documentElement.dir = activeLang === 'en' ? 'ltr' : 'rtl';
  document.body.classList.toggle('lang-en', activeLang === 'en');

  const topSel = document.getElementById('langSelect');
  const loginSel = document.getElementById('langSelectLogin');
  if(topSel) topSel.value = state.lang || '';
  if(loginSel) loginSel.value = state.lang || '';

  const setTxt = (id, txt) => { const el = document.getElementById(id); if(el) el.textContent = txt; };

  setTxt('langBadgeTop', activeLang === 'en' ? 'Language' : 'اللغة');
  setTxt('roleBadgeTop', isAdmin() ? 'Admin' : 'Viewer');

  setTxt('adminControlsTitle', activeLang === 'en' ? 'Data Upload & Control' : 'تحميل البيانات والتحكم');
  setTxt('themeBtn', activeLang === 'en'
    ? (document.body.classList.contains('light') ? 'Dark Mode' : 'Light Mode')
    : (document.body.classList.contains('light') ? 'الوضع الداكن' : 'الوضع الفاتح'));
  setTxt('exportPdfBtn', activeLang === 'en' ? 'Export PDF' : 'تصدير PDF');
  setTxt('printA4Btn', activeLang === 'en' ? 'Print / PDF A4' : 'طباعة / PDF A4');
  setTxt('exportExcelBtn', activeLang === 'en' ? 'Export Excel' : 'تصدير Excel');
  setTxt('boardPdfBtn', activeLang === 'en' ? 'Board Report / Print PDF' : 'تقرير مجلس الإدارة / طباعة PDF');
  setTxt('adminLogoutBtn', activeLang === 'en' ? 'Logout' : 'خروج');
  setTxt('addExcelBtn', activeLang === 'en' ? 'Add & Merge Excel' : 'إضافة ودمج Excel');
  setTxt('replaceExcelBtn', activeLang === 'en' ? 'Replace Current Data' : 'استبدال البيانات الحالية');
  setTxt('clearDataBtn', activeLang === 'en' ? 'Clear Data' : 'حذف الداتا');
  setTxt('resetFiltersBtn', activeLang === 'en' ? 'Clear Filters' : 'مسح الفلاتر');

  const uploadNote = document.getElementById('uploadNote');
  if(uploadNote){
    uploadNote.textContent = activeLang === 'en'
      ? 'Upload one or more Excel files. Empty cells remain blank. Platforms are read directly from column D (Tags). Duplicate orders are removed automatically using the order number.'
      : 'ترفع Excel، أو تعتمد على الداتا الحالية. أي خانة ناقصة ستظهر فاضية. المنصات تُقرأ مباشرة من العمود D: علامات التصنيف. يمكنك الآن رفع أكثر من ملف ودمجهم مع منع التكرار برقم الطلب.';
  }

  const loginBtn = document.getElementById('loginBtn');
  if(loginBtn) loginBtn.textContent = activeLang === 'en' ? 'Login' : 'دخول';
  const loginTitle = document.querySelector('.login-title-clean');
  if(loginTitle) loginTitle.textContent = activeLang === 'en' ? 'Login' : 'تسجيل الدخول';
  const loginSub = document.querySelector('.login-sub-clean');
  if(loginSub) loginSub.textContent = activeLang === 'en' ? 'Choose access type then enter the password' : 'اختر نوع الدخول ثم أدخل كلمة المرور';
  const loginPassLabel = document.querySelector('.login-pass-label');
  if(loginPassLabel) loginPassLabel.textContent = activeLang === 'en' ? 'Password' : 'كلمة المرور';

  const passInput = document.getElementById('loginPass');
  if(passInput) passInput.placeholder = activeLang === 'en' ? 'Password' : 'كلمة المرور';

  const bannerTitle = document.getElementById('bannerTitle');
  if(bannerTitle) bannerTitle.textContent = activeLang === 'en' ? 'C Store Delivery Dashboard' : 'لوحة تحكم توصيل سي ستور';
  const subtitle = document.getElementById('bannerSubtitle');
  if(subtitle){
    const spans = subtitle.querySelectorAll('span');
    if(spans.length >= 5){
      spans[0].textContent = activeLang === 'en' ? 'Operations Monitoring' : 'متابعة العمليات';
      spans[2].textContent = activeLang === 'en' ? 'Branch Performance' : 'أداء الفروع';
      spans[4].textContent = activeLang === 'en' ? 'Delivery Insights' : 'تحليلات التوصيل';
    }
  }

  const searchInput = document.getElementById('search');
  if(searchInput){
    searchInput.placeholder = activeLang === 'en'
      ? 'Search by order / customer / phone / area / driver...'
      : 'ابحث برقم الأوردر / العميل / الهاتف / المنطقة / الطيار...';
  }

  const toggleOrdersBtn = document.getElementById('toggleOrdersBtn');
  if(toggleOrdersBtn){
    toggleOrdersBtn.textContent = state.showOrders
      ? (activeLang === 'en' ? 'Hide Orders Table' : 'إخفاء جدول الأوردرات')
      : (activeLang === 'en' ? 'Show Orders Table' : 'إظهار جدول الأوردرات');
  }

  const countText = document.getElementById('countText');
  if(countText){
    const c = calc();
    countText.textContent = activeLang === 'en' ? ('Records Count: ' + c.data.length) : ('عدد السجلات: ' + c.data.length);
  }
}
function setLanguage(lang){
  if(lang !== 'ar' && lang !== 'en'){
    state.lang = '';
    localStorage.removeItem('cstore_lang');
    applyLanguage();
    return;
  }
  state.lang = lang;
  localStorage.setItem('cstore_lang', state.lang);
  render();
  applyLanguage();
}

function isAdmin(){ return state.sessionRole === 'admin'; }
function isViewer(){ return state.sessionRole === 'viewer'; }
function applyRoleUI(){
  const adminCard = document.getElementById('adminControlsCard');
  const topBadge = document.getElementById('roleBadgeTop');
  const roleBadge = document.getElementById('roleBadge');
  const label = isAdmin() ? 'Admin' : 'Viewer';
  if(topBadge){ topBadge.textContent = state.lang === 'en' ? (isAdmin() ? 'Admin' : 'Viewer') : label; topBadge.className = isAdmin() ? 'badge ok' : 'badge warn'; }
  if(roleBadge){ roleBadge.textContent = label; roleBadge.className = isAdmin() ? 'badge ok' : 'badge warn'; }
  if(adminCard){ adminCard.style.display = isAdmin() ? '' : 'none'; }
  const ordersBtn = document.getElementById('toggleOrdersBtn');
  if(ordersBtn){ ordersBtn.style.display = isAdmin() ? '' : 'none'; }
  document.body.classList.toggle('viewer-mode-clean', isViewer());
  document.querySelectorAll('[data-admin-only="true"]').forEach(el=>{
    el.style.display = isAdmin() ? '' : 'none';
  });
}

function setLoginRole(role){
  const roleSelect = document.getElementById('loginRole');
  if(roleSelect) roleSelect.value = role;
  document.querySelectorAll('.login-role-card').forEach(el=>{
    el.classList.toggle('active', el.getAttribute('data-role') === role);
  });
}

function logoutUser(){
  state.sessionRole = '';
  document.getElementById('appView').classList.add('hidden');
  document.getElementById('loginView').classList.remove('hidden');
  document.getElementById('loginPass').value = '';
  document.getElementById('loginMsg').textContent = '';
}



function bandOf(){ return ''; }
function setBand(){ return; }

function fmtPct(v){ return `${Number(v||0).toFixed(1)}%`; }


function money(v){
  return new Intl.NumberFormat("ar-EG",{maximumFractionDigits:2}).format(Number(v||0));
}
function moneyShort(v){
  const n = Number(v||0);
  const abs = Math.abs(n);
  if(abs >= 1000000000) return (n/1000000000).toFixed(1).replace(/\.0$/,'') + 'B';
  if(abs >= 1000000) return (n/1000000).toFixed(1).replace(/\.0$/,'') + 'M';
  if(abs >= 1000) return (n/1000).toFixed(1).replace(/\.0$/,'') + 'K';
  return new Intl.NumberFormat("ar-EG",{maximumFractionDigits:2}).format(n);
}


function ymd(d){const x=new Date(d); return {y:x.getFullYear(),m:x.getMonth()+1,day:x.getDate()}}
function hourLabel(t){return String(t||"").slice(0,2)+":00"}
function performanceScore(successRate,salesWeight,volumeWeight,cancelRate){let score=(successRate*0.5)+(salesWeight*0.25)+(volumeWeight*0.2)-((cancelRate||0)*0.15); score=Math.max(0,Math.min(100,score)); return score}
function scoreBadge(score){if(score>=85) return '<span class="badge ok">ممتاز</span>'; if(score>=70) return '<span class="badge warn">جيد</span>'; return '<span class="badge bad">يحتاج تحسين</span>'}
function normalizeStatus(v){const s=String(v||"").trim(); if(s==="مفوتر بالكامل") return "تم"; if(s==="لا توجد مبالغ لفوترتها") return "ملغي / مرتجع"; return s||""}
function firstNonEmpty(){for(const v of arguments){const s=String(v??"").trim(); if(s&&s.toLowerCase()!=='nan') return s;} return ""}
function toAmount(v){if(v===null||v===undefined||v==="") return 0; const s=String(v).replace(/,/g,'').trim(); const n=Number(s); return isNaN(n)?0:n}
function transformExcelRows(rows){
  const branchMap={"فرع تريومف":"فرع مصر الجديدة","التجمع":"فرع التجمع الاول","فرع ميدان الجزائر":"فرع المعادي","ABBAS BRANCH":"فرع مدينة نصر"};
  function excelDate(v){const p=n=>String(n).padStart(2,'0'); if(typeof v==="number" && window.XLSX && XLSX.SSF && XLSX.SSF.parse_date_code){const d=XLSX.SSF.parse_date_code(v); if(d) return {date:`${d.y}-${p(d.m)}-${p(d.d)}`,time:`${p(d.H||0)}:${p(d.M||0)}`};} const dt=new Date(v); if(isNaN(dt)) return {date:firstNonEmpty(v).slice(0,10),time:firstNonEmpty(v).slice(11,16)}; return {date:`${dt.getFullYear()}-${p(dt.getMonth()+1)}-${p(dt.getDate())}`,time:`${p(dt.getHours())}:${p(dt.getMinutes())}`}}
  return rows.map(r=>{const dt=excelDate(r["تاريخ الطلب"]); return {orderNumber:firstNonEmpty(r["مرجع الطلب"]), orderDate:dt.date, orderTime:dt.time, branch:branchMap[firstNonEmpty(r["المستودع"])]||firstNonEmpty(r["المستودع"],""), platform:firstNonEmpty(r["علامات التصنيف"], ""), customerName:firstNonEmpty(r["عنوان الفاتورة"],r["العميل/علامات التصنيف /اسم العرض"],"عميل غير مسمى"), customerPhone:firstNonEmpty(r["العميل/الهاتف المحمول"],r["العميل/رقم الهاتف"],""), area:firstNonEmpty(r["العميل/الشارع"],r["العميل/الشارع 2"],""), amount:toAmount(r["الإجمالي"]), paymentMethod:firstNonEmpty(r["Payment Method(POS)(1)"],r["Payment Method(POS)(2)"],""), preparer:firstNonEmpty(r["Prepared By"],""), driver:firstNonEmpty(r["Delivery Man/اسم العرض"],""), status:normalizeStatus(r["حالة الفاتورة"])};}).filter(x=>x.orderNumber)
}
function filtered(){return state.orders.filter(o=>{const p=ymd(o.orderDate),f=state.filters,text=[o.orderNumber,o.customerName,o.customerPhone,o.area,o.branch,o.preparer,o.driver,o.platform,o.paymentMethod].join(" ").toLowerCase(); return (!f.years.length||f.years.includes(p.y))&&(!f.months.length||f.months.includes(p.m))&&(!f.days.length||f.days.includes(p.day))&&(!f.branches.length||f.branches.includes(o.branch))&&(!f.statuses.length||f.statuses.includes(o.status))&&(!f.preparers.length||f.preparers.includes(o.preparer))&&(!f.drivers.length||f.drivers.includes(o.driver))&&(!f.areas.length||f.areas.includes(o.area))&&(!f.platforms.length||f.platforms.includes(o.platform))&&(!f.payments.length||f.payments.includes(o.paymentMethod))&&(!state.search||text.includes(state.search.toLowerCase()));})}

function countBy(data,key){return Object.entries(data.reduce((a,o)=>{const k=o[key]||""; if(!k) return a; a[k]=(a[k]||0)+1; return a;},{})).map(([name,count])=>({name,count})).sort((a,b)=>b.count-a.count)}
function getMonthSummary(data){const map={}; data.forEach(o=>{const mk=(o.orderDate||'').slice(0,7); if(!mk) return; if(!map[mk]) map[mk]={month:mk,orders:0,sales:0,canceled:0,completed:0}; map[mk].orders++; if(o.status==='تم'){map[mk].completed++; map[mk].sales+=Number(o.amount||0)} else if(o.status==='ملغي / مرتجع') map[mk].canceled++;}); return Object.values(map).map(m=>({...m,cancelRate:m.orders?(m.canceled/m.orders)*100:0,avgOrder:m.completed?(m.sales/m.completed):0})).sort((a,b)=>a.month.localeCompare(b.month))}
function getDailySummary(data){const dates=[...new Set(data.map(x=>x.orderDate).filter(Boolean))].sort(); const today=dates.slice(-1)[0]||'', prev=dates.slice(-2,-1)[0]||''; const dayStats=d=>{const rows=data.filter(x=>x.orderDate===d), done=rows.filter(x=>x.status==='تم'), bad=rows.filter(x=>x.status==='ملغي / مرتجع'); return {date:d,orders:rows.length,sales:done.reduce((s,x)=>s+Number(x.amount||0),0),cancelRate:rows.length?(bad.length/rows.length)*100:0,peakHour:countBy(rows.map(x=>({...x,hour:hourLabel(x.orderTime)})),'hour')[0]?.name||'—'}}; return {today:dayStats(today), yesterday:dayStats(prev)}}
function calc(){
  const data=filtered(), done=data.filter(x=>x.status==="تم"), bad=data.filter(x=>x.status==="ملغي / مرتجع");
  const sales=done.reduce((s,x)=>s+Number(x.amount||0),0), avg=done.length?sales/done.length:0, cancelRate=data.length?(bad.length/data.length)*100:0, hourly=countBy(data.map(x=>({...x,hour:hourLabel(x.orderTime)})),'hour');
  const branchMap={}, prepMap={}, driverMap={}, platMap={}, areaMap={}, customerMap={};
  const maxSalesSeed=Math.max(1,sales), maxOrdersSeed=Math.max(1,data.length);
  data.forEach(o=>{
    if(o.branch){if(!branchMap[o.branch]) branchMap[o.branch]={name:o.branch,orders:0,completed:0,canceled:0,sales:0,areas:{},customers:{},hours:{},preparers:{},drivers:{}}; const b=branchMap[o.branch]; b.orders++; b.hours[hourLabel(o.orderTime)]=(b.hours[hourLabel(o.orderTime)]||0)+1; if(o.area) b.areas[o.area]=(b.areas[o.area]||0)+1; if(o.customerName) b.customers[o.customerName]=(b.customers[o.customerName]||0)+1; if(o.preparer) b.preparers[o.preparer]=(b.preparers[o.preparer]||0)+1; if(o.driver) b.drivers[o.driver]=(b.drivers[o.driver]||0)+1; if(o.status==="تم"){b.completed++; b.sales+=Number(o.amount||0)} else b.canceled++;}
    if(o.preparer){if(!prepMap[o.preparer]) prepMap[o.preparer]={name:o.preparer,orders:0,completed:0,canceled:0,sales:0}; prepMap[o.preparer].orders++; if(o.status==="تم"){prepMap[o.preparer].completed++; prepMap[o.preparer].sales+=Number(o.amount||0)} else prepMap[o.preparer].canceled++;}
    if(o.driver){if(!driverMap[o.driver]) driverMap[o.driver]={name:o.driver,orders:0,completed:0,canceled:0,sales:0}; driverMap[o.driver].orders++; if(o.status==="تم"){driverMap[o.driver].completed++; driverMap[o.driver].sales+=Number(o.amount||0)} else driverMap[o.driver].canceled++;}
    if(o.platform){if(!platMap[o.platform]) platMap[o.platform]={name:o.platform,orders:0,completed:0,canceled:0,sales:0}; platMap[o.platform].orders++; if(o.status==="تم"){platMap[o.platform].completed++; platMap[o.platform].sales+=Number(o.amount||0)} else platMap[o.platform].canceled++;}
    if(o.area){if(!areaMap[o.area]) areaMap[o.area]={name:o.area,orders:0,completed:0,sales:0,canceled:0}; areaMap[o.area].orders++; if(o.status==="تم"){areaMap[o.area].completed++; areaMap[o.area].sales+=Number(o.amount||0)} else areaMap[o.area].canceled++;}
    if(o.customerPhone || o.customerName){const key=o.customerPhone || ('name:'+o.customerName); if(!customerMap[key]) customerMap[key]={name:o.customerName,phone:o.customerPhone,orders:0,spent:0,branches:{}}; customerMap[key].orders++; if(o.status==="تم") customerMap[key].spent+=Number(o.amount||0); if(o.branch) customerMap[key].branches[o.branch]=(customerMap[key].branches[o.branch]||0)+1;}
  });
  let branches=Object.values(branchMap).map(b=>{const successRate=b.orders?(b.completed/b.orders)*100:0, avgOrder=b.completed?(b.sales/b.completed):0, cancelRateB=b.orders?(b.canceled/b.orders)*100:0, salesWeight=(b.sales/maxSalesSeed)*100, volumeWeight=(b.orders/maxOrdersSeed)*100, score=performanceScore(successRate,salesWeight,volumeWeight,cancelRateB), branchMonths=[...new Set(data.filter(o=>o.branch===b.name).map(o=>getMonthKeyFromRowDate(o.orderDate)).filter(Boolean))], target=branchMonths.reduce((s,mk)=>s+getMonthlyTarget(b.name,mk),0), targetPct=target?(b.sales/target)*100:0; return {...b,successRate,avgOrder,cancelRate:cancelRateB,score,target,targetPct,topArea:Object.entries(b.areas).sort((a,b)=>b[1]-a[1])[0]?.[0]||'',topCustomer:Object.entries(b.customers).sort((a,b)=>b[1]-a[1])[0]?.[0]||'',peak:Object.entries(b.hours).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—',bestPreparer:Object.entries(b.preparers).sort((a,b)=>b[1]-a[1])[0]?.[0]||'',bestDriver:Object.entries(b.drivers).sort((a,b)=>b[1]-a[1])[0]?.[0]||''};}).sort((a,b)=>b.score-a.score||b.sales-a.sales);
  const prepMaxSales=Math.max(1,...Object.values(prepMap).map(x=>x.sales),1), prepMaxOrders=Math.max(1,...Object.values(prepMap).map(x=>x.orders),1);
  const prepMaxCompleted=Math.max(1,...Object.values(prepMap).map(x=>x.completed),1);
  let preparers=Object.values(prepMap).map(p=>{
    const successRate=p.orders?(p.completed/p.orders)*100:0, cancelRateP=p.orders?(p.canceled/p.orders)*100:0, avgOrder=p.completed?(p.sales/p.completed):0;
    const kpi = findKpiForPreparer(p.name);
    const completedScore = (Number(p.completed||0) / prepMaxCompleted) * 100;
    return {...p,successRate,cancelRate:cancelRateP,avgOrder,score:completedScore,kpiPct:kpi?.pct||null,kpiFinal:kpi?.final||null};
  }).sort((a,b)=>(Number(b.completed||0)-Number(a.completed||0)) || (Number(b.kpiPct||0)-Number(a.kpiPct||0)) || (Number(b.orders||0)-Number(a.orders||0)));
  const driverMaxSales=Math.max(1,...Object.values(driverMap).map(x=>x.sales),1), driverMaxOrders=Math.max(1,...Object.values(driverMap).map(x=>x.orders),1);
  let drivers=Object.values(driverMap).map(d=>{const successRate=d.orders?(d.completed/d.orders)*100:0, cancelRateD=d.orders?(d.canceled/d.orders)*100:0, avgOrder=d.completed?(d.sales/d.completed):0; return {...d,successRate,cancelRate:cancelRateD,avgOrder,score:performanceScore(successRate,(d.sales/driverMaxSales)*100,(d.orders/driverMaxOrders)*100,cancelRateD)}}).sort((a,b)=>b.score-a.score||b.completed-a.completed);
  let platforms=Object.values(platMap).map(p=>{const successRate=p.orders?(p.completed/p.orders)*100:0, cancelRateP=p.orders?(p.canceled/p.orders)*100:0, avgOrder=p.completed?(p.sales/p.completed):0; return {...p,successRate,cancelRate:cancelRateP,avgOrder}}).sort((a,b)=>b.sales-a.sales||b.orders-a.orders);
  let areas=Object.values(areaMap).map(a=>{const cancelRateA=a.orders?(a.canceled/a.orders)*100:0, avgOrder=a.completed?(a.sales/a.completed):0; return {...a,cancelRate:cancelRateA,avgOrder}}).sort((a,b)=>b.completed-a.completed||b.sales-a.sales);
  let customers=Object.values(customerMap).map(c=>{const topBranch=Object.entries(c.branches).sort((a,b)=>b[1]-a[1])[0]?.[0]||''; const avgOrder=c.orders?(c.spent/c.orders):0; return {...c,topBranch,avgOrder}}).sort((a,b)=>b.orders-a.orders||b.spent-a.spent);
  const monthlyOrders=getMonthSummary(data), daily=getDailySummary(data), payments=countBy(data,'paymentMethod'), alerts=[];
  const visibleMonthKeys = [...new Set(data.map(o=>getMonthKeyFromRowDate(o.orderDate)).filter(Boolean))];
  const targetSummary = {
    target: branches.reduce((s,b)=>s + visibleMonthKeys.reduce((mSum,mk)=>mSum + getMonthlyTarget(b.name, mk), 0),0),
    sales: branches.reduce((s,b)=>s+Number(b.sales||0),0)
  };
  targetSummary.pct = targetSummary.target ? (targetSummary.sales/targetSummary.target)*100 : 0;
  targetSummary.remaining = Math.max(0, targetSummary.target - targetSummary.sales);
  if(cancelRate>Number(state.settings.alertThresholds.cancelRate||10)) alerts.push('Current cancellation rate مرتفعة ('+cancelRate.toFixed(1)+'%)');
  if(branches.some(b=>b.target && b.targetPct<70)) alerts.push('هناك فروع أقل من 70% من هدف المبيعات');
  if(monthlyOrders.length>=2){const cur=monthlyOrders[monthlyOrders.length-1], prev=monthlyOrders[monthlyOrders.length-2]; if(cur.sales<prev.sales) alerts.push('مبيعات '+cur.month+' أقل من '+prev.month); if(cur.cancelRate>prev.cancelRate) alerts.push('إلغاء '+cur.month+' أعلى من '+prev.month);}
  const bestKpiPreparer = getBestKpiPreparer();
  return {data,total:done.length,totalAll:data.length,sales,avg,badCount:bad.length,cancelRate,branches,preparers,drivers,platforms,areas,customers,payments,hourly,monthlyOrders,daily,alerts,targetSummary,topCustomer:customers[0]||null,peakHour:hourly[0]?.name||"—",bestKpiPreparer,kpiPreparerMonth:getActiveKpiMonth()}
}
function toggleFilter(key,val){
  const arr=state.filters[key];
  state.filters[key]=arr.includes(val)?arr.filter(x=>x!==val):[...arr,val];
  render();
}
function toggleFilterFromEvent(el){
  const key = el.getAttribute('data-key');
  let val = decodeURIComponent(el.getAttribute('data-value') || '');
  if(key==='years' || key==='months' || key==='days'){
    const num = Number(val);
    val = Number.isNaN(num) ? val : num;
  }
  toggleFilter(key, val);
}
function renderMultiSelect(title,items,key){
  const selected=state.filters[key], label=selected.length?`${title}: ${selected.length}`:title;
  return `<div class="ms"><button type="button" class="ms-btn" onclick="toggleMenu(this)">${label} ▼</button><div class="ms-menu">${items.map(it=>{
    const v=it.value??it, t=it.label??it, checked=selected.includes(v)?'checked':'';
    const encoded = encodeURIComponent(String(v));
    return `<label class="ms-option"><input type="checkbox" data-key="${key}" data-value="${encoded}" ${checked} onchange="toggleFilterFromEvent(this)"><span>${t}</span></label>`;
  }).join("")}</div></div>`
}
function toggleMenu(btn){document.querySelectorAll('.ms').forEach(x=>{if(x!==btn.parentElement)x.classList.remove('open')}); btn.parentElement.classList.toggle('open')}
document.addEventListener('click',(e)=>{if(!e.target.closest('.ms'))document.querySelectorAll('.ms').forEach(x=>x.classList.remove('open'))})
function quickFilter(key,val){if(!val) return; state.filters[key]=[val]; if(key==='branches') state.activeTab='branches'; if(key==='preparers') state.activeTab='preparers'; if(key==='drivers') state.activeTab='drivers'; if(key==='platforms') state.activeTab='platforms'; if(key==='areas') state.activeTab='areas'; render()}
function setSearchValue(val){state.search=val||''; document.getElementById('search').value=state.search; state.activeTab='customers'; render()}
function clearAllData(){state.orders=[]; state.search=''; state.filters={years:[], months:[], days:[], branches:[], statuses:[], preparers:[], drivers:[], areas:[], platforms:[], payments:[]}; document.getElementById('search').value=''; document.getElementById('msg').textContent='تم حذف الداتا الحالية.'; document.getElementById('uploadStatus').innerHTML=''; render()}
function progressList(items,valKey,suffix,filterKey=''){if(!items.length) return '<div class="muted">No Data</div>'; const max=Math.max(...items.map(x=>x[valKey]||0),1); return items.map(i=>`<div class="row" style="display:block;cursor:${filterKey?'pointer':'default'}" ${filterKey?`onclick="quickFilter('${filterKey}', ${JSON.stringify(i.name)})"`:''}><div style="display:flex;justify-content:space-between"><strong>${i.name}</strong><span>${money(i[valKey]||0)}${suffix}</span></div><div class="progress"><div style="width:${((i[valKey]||0)/max)*100}%"></div></div></div>`).join('')}
function updatePrintReport(c){
  if(!document.getElementById('reportDate')) return;
  const weakBranch = [...(c.branches||[])].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const weakDriver = [...(c.drivers||[])].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const weakPlatform = [...(c.platforms||[])].sort((a,b)=>(a.completed||0)-(b.completed||0))[0];
  const paymentStatsForReport = (c.payments||[]).map(p=>{const rows=(c.data||[]).filter(x=>x.paymentMethod===p.name); const done=rows.filter(x=>x.status==="تم"); const sales=done.reduce((s,x)=>s+Number(x.amount||0),0); return {name:p.name, orders:rows.length, completed:done.length, sales};}).sort((a,b)=>b.completed-a.completed||b.sales-a.sales);
  const dailyRowsForReport = {};
  (c.data||[]).forEach(o=>{const d=o.orderDate||''; if(!d) return; if(!dailyRowsForReport[d]) dailyRowsForReport[d]={date:d,orders:0,completed:0,canceled:0,sales:0,branches:{}}; dailyRowsForReport[d].orders++; if(o.branch) dailyRowsForReport[d].branches[o.branch]=(dailyRowsForReport[d].branches[o.branch]||0)+1; if(o.status==="تم"){dailyRowsForReport[d].completed++; dailyRowsForReport[d].sales+=Number(o.amount||0)} else if(o.status==="ملغي / مرتجع"){dailyRowsForReport[d].canceled++;}});
  const latestDaily = Object.values(dailyRowsForReport).sort((a,b)=>b.date.localeCompare(a.date))[0] || null;

  document.getElementById('reportDate').textContent='تاريخ الإصدار: '+new Date().toLocaleString('ar-EG');
  document.getElementById('reportExecutive').innerHTML =
    `يعرض هذا التقرير قراءة تنفيذية مركزة لأداء نشاط التوصيل والمبيعات خلال الفترة الحالية. بلغ إجمالي الطلبات المكتملة <strong>${c.total}</strong> طلبًا، بإجمالي مبيعات <strong>${money(c.sales)}</strong> ومتوسط أوردر <strong>${money(c.avg)}</strong>. سجلت العمليات نسبة إلغاء قدرها <strong>${c.cancelRate.toFixed(1)}%</strong>، بينما بلغت نسبة تحقيق التارجت الكلية <strong>${(c.targetSummary?.pct||0).toFixed(1)}%</strong>. ويظهر من المؤشرات أن الفرع الأعلى أداءً هو <strong>${c.branches[0]?.name||'—'}</strong>، في حين يحتاج <strong>${weakBranch?.name||'—'}</strong> إلى متابعة أقرب لتحسين النتائج. كما تعد منصة <strong>${c.platforms[0]?.name||'—'}</strong> هي الأكثر مساهمة، بينما يمثل كل من <strong>${c.drivers[0]?.name||'—'}</strong> و<strong>${c.preparers[0]?.name||'—'}</strong> الأفضل تشغيليًا ضمن الفترة الحالية.`;

  document.getElementById('reportKpis').innerHTML=
    `<div class="a4-card"><div class="label">الطلبات المكتملة</div><div class="value">${c.total}</div></div>
     <div class="a4-card"><div class="label">إجمالي المبيعات</div><div class="value">${money(c.sales)}</div></div>
     <div class="a4-card"><div class="label">متوسط الأوردر</div><div class="value">${money(c.avg)}</div></div>
     <div class="a4-card"><div class="label">نسبة الإلغاء</div><div class="value">${c.cancelRate.toFixed(1)}%</div></div>
     <div class="a4-card"><div class="label">تحقيق التارجت</div><div class="value">${(c.targetSummary?.pct||0).toFixed(1)}%</div></div>
     <div class="a4-card"><div class="label">أعلى ساعة ضغط</div><div class="value">${c.peakHour}</div></div>`;

  document.getElementById('reportSummary').innerHTML=
    `<tr class="a4-highlight"><td>أفضل محضّر</td><td>${c.preparers[0]?.name||'—'}</td></tr>
     <tr><td>إجمالي الأوردرات المحققة</td><td>${c.preparers[0]?.completed||0} أوردر</td></tr>
     <tr class="a4-highlight"><td>أفضل طيار</td><td>${c.drivers[0]?.name||'—'}</td></tr>
     <tr><td>إجمالي الأوردرات المسلمة</td><td>${c.drivers[0]?.completed||0} أوردر</td></tr>
     <tr class="a4-highlight"><td>أفضل عميل</td><td>${c.topCustomer?.name||'—'}</td></tr>
     <tr><td>عدد الطلبات خلال الفترة</td><td>${c.topCustomer?.orders||0} طلب</td></tr>`;

  document.getElementById('reportManager').innerHTML=
    `<tr class="a4-highlight"><td>أفضل فرع</td><td>${c.branches[0]?.name||'—'}</td></tr>
     <tr><td>الفرع الأضعف</td><td>${weakBranch?.name||'—'}</td></tr>
     <tr class="a4-highlight"><td>أفضل طيار</td><td>${c.drivers[0]?.name||'—'}</td></tr>
     <tr><td>الطيار الأضعف</td><td>${weakDriver?.name||'—'}</td></tr>
     <tr class="a4-highlight"><td>نسبة تحقيق التارجت</td><td>${(c.targetSummary?.pct||0).toFixed(1)}%</td></tr>
     <tr><td>Remaining للوصول للتارجت</td><td>${money(c.targetSummary?.remaining||0)}</td></tr>`;

  document.getElementById('reportTop').innerHTML=
    `<tr class="a4-highlight"><td>أفضل منصة</td><td>${c.platforms[0]?.name||'—'}</td></tr>
     <tr><td>عدد أوردرات أفضل منصة</td><td>${c.platforms[0]?.completed||0} طلب</td></tr>
     <tr class="a4-highlight"><td>أعلى فرع مبيعات</td><td>${[...c.branches].sort((a,b)=>b.sales-a.sales)[0]?.name||'—'}</td></tr>
     <tr><td>أعلى منطقة</td><td>${c.areas[0]?.name||'—'}</td></tr>`;

  document.getElementById('reportLeaders').innerHTML=
    `<tr class="a4-highlight"><td>أفضل فرع أداء</td><td>${c.branches[0]?.name||'—'} (${(c.branches[0]?.score||0).toFixed(1)}%)</td></tr>
     <tr><td>أفضل محضّر أداء</td><td>${c.preparers[0]?.name||'—'} (${(c.preparers[0]?.score||0).toFixed(1)}%)</td></tr>
     <tr class="a4-highlight"><td>أفضل طيار أداء</td><td>${c.drivers[0]?.name||'—'} (${(c.drivers[0]?.score||0).toFixed(1)}%)</td></tr>
     <tr><td>أضعف منصة تشغيلية</td><td>${weakPlatform?.name||'—'} (${weakPlatform?.completed||0} طلب مكتمل)</td></tr>`;

  document.getElementById('reportDaily').innerHTML=
    `<tr class="a4-highlight"><td>آخر يوم في البيانات</td><td>${latestDaily?.date||'—'}</td></tr>
     <tr><td>إجمالي الطلبات</td><td>${latestDaily?.orders||0}</td></tr>
     <tr><td>الطلبات المكتملة</td><td>${latestDaily?.completed||0}</td></tr>
     <tr><td>الملغي / المرتجع</td><td>${latestDaily?.canceled||0}</td></tr>
     <tr><td>Last Day Sales</td><td>${money(latestDaily?.sales||0)}</td></tr>
     <tr><td>أفضل فرع في آخر يوم</td><td>${latestDaily ? (Object.entries(latestDaily.branches).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—') : '—'}</td></tr>`;

  document.getElementById('reportPayments').innerHTML=
    `<tr class="a4-highlight"><td>أفضل طريقة دفع</td><td>${paymentStatsForReport[0]?.name||'—'}</td></tr>
     <tr><td>عدد الطلبات المكتملة لأفضل طريقة</td><td>${paymentStatsForReport[0]?.completed||0} طلب</td></tr>
     <tr><td>إجمالي مبيعات أفضل طريقة</td><td>${money(paymentStatsForReport[0]?.sales||0)}</td></tr>`;

  document.getElementById('reportCharts').innerHTML=
    `<div class="a4-two-col">
      <div class="a4-section">
        <div class="a4-small"><strong>مبيعات الفروع</strong></div>
        ${reportBars(c.branches.slice(0,6),'name','sales')}
      </div>
      <div class="a4-section">
        <div class="a4-small"><strong>الطلبات المكتملة حسب المنصات</strong></div>
        ${reportBars(c.platforms.slice(0,6),'name','completed',' طلب')}
      </div>
     </div>
     <div class="a4-two-col">
      <div class="a4-section">
        <div class="a4-small"><strong>أفضل المحضّرين</strong></div>
        ${reportBars(c.preparers.slice(0,6),'name','completed',' أوردر')}
      </div>
      <div class="a4-section">
        <div class="a4-small"><strong>أفضل الطيارين</strong></div>
        ${reportBars(c.drivers.slice(0,6),'name','completed',' أوردر')}
      </div>
     </div>
     <div class="a4-section">
       <div class="a4-small"><strong>المبيعات الشهرية</strong></div>
       ${reportBars(c.monthlyOrders.slice(-6),'month','sales')}
     </div>`;
}
function tabsHtml(){
  const isEn = state.lang === 'en';
  const adminTabs=[
    ['overview', isEn ? 'Overview' : 'نظرة عامة','🏠'],
    ['daily', isEn ? 'Daily Report' : 'تقرير يومي منفصل','📅'],
    ['period', isEn ? 'Period Analysis' : 'تحليل الفترات','📈'],
    ['charts', isEn ? 'Charts' : 'الرسوم البيانية','📊'],
    ['branchpeaks', isEn ? 'Daily Branch Peaks' : 'ذروة الفروع اليومية','⏰'],
    ['branches', isEn ? 'Branches' : 'الفروع','🏬'],
    ['drivers', isEn ? 'Drivers' : 'الطيارين','🛵'],
    ['platforms', isEn ? 'Platforms' : 'المنصات','📱'],
    ['areas', isEn ? 'Areas' : 'المناطق','📍'],
    ['customers', isEn ? 'Customers' : 'العملاء','👥'],
    ['kpi','KPI','🎯'],
    ['settings', isEn ? 'Settings' : 'الإعدادات','⚙️']
  ];
  const viewerTabs=[
    ['overview', isEn ? 'Overview' : 'نظرة عامة','🏠'],
    ['period', isEn ? 'Period Analysis' : 'تحليل الفترات','📈'],
    ['charts', isEn ? 'Charts' : 'الرسوم البيانية','📊'],
    ['branchpeaks', isEn ? 'Daily Branch Peaks' : 'ذروة الفروع اليومية','⏰'],
    ['branches', isEn ? 'Branches' : 'الفروع','🏬'],
    ['drivers', isEn ? 'Drivers' : 'الطيارين','🛵'],
    ['platforms', isEn ? 'Platforms' : 'المنصات','📱'],
    ['areas', isEn ? 'Areas' : 'المناطق','📍'],
    ['customers', isEn ? 'Customers' : 'العملاء','👥'],
    ['kpi','KPI','🎯']
  ];
  const tabs = isAdmin() ? adminTabs : viewerTabs;
  return tabs.map(t=>`<button class="tab ${state.activeTab===t[0]?'active':''}" onclick="setTab('${t[0]}')"><span class="tab-icon">${t[2]}</span><span class="tab-label">${t[1]}</span></button>`).join('');
}
function setTab(t){ if(t==='settings' && !isAdmin()) return; if(t==='manager' || t==='insights' || t==='rankings' || t==='payments' || t==='preparers') t='overview'; state.activeTab=t; render()}


function renderLineChart(items, labelKey, valueKey, suffix=''){
  if(!items || !items.length) return '<div class="muted">No Data</div>';
  const width = Math.max(560, items.length * 90);
  const height = 280;
  const margin = {top: 22, right: 20, bottom: 62, left: 46};
  const innerW = width - margin.left - margin.right;
  const innerH = height - margin.top - margin.bottom;
  const values = items.map(x=>Number(x[valueKey]||0));
  const maxV = Math.max(...values, 1);
  const stepX = items.length > 1 ? innerW / (items.length - 1) : innerW / 2;
  const points = items.map((it, i) => {
    const v = Number(it[valueKey]||0);
    const x = margin.left + (items.length > 1 ? i * stepX : innerW / 2);
    const y = margin.top + innerH - ((v / maxV) * innerH);
    return {x, y, v, label: it[labelKey]||''};
  });
  const polyline = points.map(p => `${p.x},${p.y}`).join(' ');
  const yTicks = 4;
  const grid = Array.from({length:yTicks+1}, (_,i)=>{
    const val = Math.round((maxV / yTicks) * (yTicks-i));
    const y = margin.top + (innerH / yTicks) * i;
    return `<line class="grid-line" x1="${margin.left}" y1="${y}" x2="${width-margin.right}" y2="${y}"></line>
            <text class="axis-label" x="${margin.left-8}" y="${y+4}" text-anchor="end">${money(val)}</text>`;
  }).join('');
  const xLabels = points.map(p=>`<text class="axis-label" x="${p.x}" y="${height-18}" text-anchor="middle">${p.label}</text>`).join('');
  const dots = points.map(p=>`<circle class="point-fill" cx="${p.x}" cy="${p.y}" r="4.5"></circle>
    <text class="point-value" x="${p.x}" y="${p.y-10}" text-anchor="middle">${money(p.v)}${suffix}</text>`).join('');
  return `<div class="line-chart-wrap"><svg class="line-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMidYMid meet">
    ${grid}
    <line class="axis-line" x1="${margin.left}" y1="${margin.top+innerH}" x2="${width-margin.right}" y2="${margin.top+innerH}"></line>
    <polyline class="line-stroke" points="${polyline}"></polyline>
    ${dots}
    ${xLabels}
  </svg></div>`;
}
function renderColumnChart(items, labelKey, valueKey, suffix=''){
  if(!items || !items.length) return '<div class="muted">No Data</div>';
  const width = Math.max(560, items.length * 78);
  const height = 300;
  const margin = {top: 24, right: 20, bottom: 72, left: 46};
  const innerW = width - margin.left - margin.right;
  const innerH = height - margin.top - margin.bottom;
  const values = items.map(x=>Number(x[valueKey]||0));
  const maxV = Math.max(...values, 1);
  const barGap = 14;
  const barW = Math.max(22, (innerW - barGap*(items.length-1)) / items.length);
  const yTicks = 4;
  const grid = Array.from({length:yTicks+1}, (_,i)=>{
    const val = Math.round((maxV / yTicks) * (yTicks-i));
    const y = margin.top + (innerH / yTicks) * i;
    return `<line class="grid-line" x1="${margin.left}" y1="${y}" x2="${width-margin.right}" y2="${y}"></line>
            <text class="axis-label" x="${margin.left-8}" y="${y+4}" text-anchor="end">${money(val)}</text>`;
  }).join('');
  const bars = items.map((it, i)=>{
    const v = Number(it[valueKey]||0);
    const h = (v/maxV) * innerH;
    const x = margin.left + i * (barW + barGap);
    const y = margin.top + innerH - h;
    return `<rect class="column-bar" x="${x}" y="${y}" width="${barW}" height="${h}" rx="8"></rect>
            <text class="point-value" x="${x + barW/2}" y="${y-8}" text-anchor="middle">${money(v)}${suffix}</text>
            <text class="axis-label" x="${x + barW/2}" y="${height-18}" text-anchor="middle">${it[labelKey]||''}</text>`;
  }).join('');
  return `<div class="column-chart-wrap"><svg class="column-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMidYMid meet">
    <defs><linearGradient id="barGradPeak" x1="0" x2="0" y1="0" y2="1"><stop offset="0%" stop-color="#ff5da4"/><stop offset="100%" stop-color="#b72067"/></linearGradient></defs>
    ${grid}
    <line class="axis-line" x1="${margin.left}" y1="${margin.top+innerH}" x2="${width-margin.right}" y2="${margin.top+innerH}"></line>
    ${bars}
  </svg></div>`;
}

function renderBars(items, labelKey, valueKey, suffix=''){
  if(!items || !items.length) return '<div class="muted">No Data</div>';
  const max=Math.max(...items.map(x=>Number(x[valueKey]||0)),1);
  return `<div class="chart-list">${items.map(it=>{
    const v=Number(it[valueKey]||0), w=(v/max)*100;
    return `<div class="chart-row"><div>${it[labelKey]||''}</div><div class="chart-bar-wrap"><div class="chart-bar" style="width:${w}%"></div></div><div>${money(v)}${suffix}</div></div>`;
  }).join('')}</div>`;
}

function monthlyCompareCard(c){if(c.monthlyOrders.length<2) return '<div class="note">No Data كافية للمقارنة الشهرية.</div>'; const cur=c.monthlyOrders[c.monthlyOrders.length-1], prev=c.monthlyOrders[c.monthlyOrders.length-2], salesDiff=cur.sales-prev.sales, ordersDiff=cur.orders-prev.orders, cancelDiff=cur.cancelRate-prev.cancelRate; return `<div class="tableWrap"><table><thead><tr><th>المقارنة</th><th>${prev.month}</th><th>${cur.month}</th><th>الفرق</th></tr></thead><tbody><tr><td>الطلبات</td><td>${prev.orders}</td><td>${cur.orders}</td><td>${ordersDiff}</td></tr><tr><td>المبيعات</td><td>${money(prev.sales)}</td><td>${money(cur.sales)}</td><td>${money(salesDiff)}</td></tr><tr><td>نسبة الإلغاء</td><td>${prev.cancelRate.toFixed(1)}%</td><td>${cur.cancelRate.toFixed(1)}%</td><td>${cancelDiff.toFixed(1)}%</td></tr></tbody></table></div>`}
function safeId(value){
  return String(value||'').replace(/[^a-zA-Z0-9؀-ۿ_-]/g,'_');
}
function parseTargetNumber(val){
  const arabic = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  let s = String(val||'').trim().replace(/[٠-٩]/g, d => arabic[d]).replace(/,/g,'').replace(/٬/g,'');
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function setPendingTarget(branch,val){}
function applyAllMonthlyTargets(){
  if(!state.settings.monthlySalesTargets) state.settings.monthlySalesTargets = {};
  const inputs = Array.from(document.querySelectorAll('input[data-target-branch][data-target-month]'));
  let count = 0;
  inputs.forEach(input=>{
    const branch = input.getAttribute('data-target-branch') || '';
    const monthKey = input.getAttribute('data-target-month') || '';
    if(!branch || !monthKey) return;
    if(!state.settings.monthlySalesTargets[branch]) state.settings.monthlySalesTargets[branch] = {};
    state.settings.monthlySalesTargets[branch][monthKey] = parseTargetNumber(input.value);
    count++;
  });
  saveSettings();
  document.getElementById('msg').textContent = 'تم تطبيق ' + count + ' هدف شهري بنجاح';
  render();
}
function changePassword(){const v=document.getElementById('newPass').value.trim(); if(!v) return; state.settings.password=v; saveSettings(); document.getElementById('msg').textContent='تم تغيير كلمة مرور Admin.'; document.getElementById('newPass').value=''}
function changeViewerPassword(){const v=document.getElementById('newViewerPass').value.trim(); if(!v) return; state.settings.viewerPassword=v; saveSettings(); document.getElementById('msg').textContent='تم تغيير كلمة مرور Viewer.'; document.getElementById('newViewerPass').value=''}
function saveAlertSettings(){state.settings.alertThresholds.cancelRate=Number(document.getElementById('cancelThreshold').value||10); saveSettings(); document.getElementById('msg').textContent='تم حفظ إعدادات التنبيه.'; render()}

function getMonthKeyFromRowDate(dateStr){
  return String(dateStr||'').slice(0,7);
}
function allMonthKeysFromData(){
  return [...new Set((state.orders||[]).map(o=>getMonthKeyFromRowDate(o.orderDate)).filter(Boolean))].sort().reverse();
}
function getMonthlyTarget(branch, monthKey){
  return Number((((state.settings||{}).monthlySalesTargets||{})[branch]||{})[monthKey]||0);
}
function setMonthlyTarget(branch, monthKey, value){
  if(!state.settings.monthlySalesTargets[branch]) state.settings.monthlySalesTargets[branch] = {};
  state.settings.monthlySalesTargets[branch][monthKey] = parseTargetNumber(value);
  saveSettings();
}
function monthlyTargetHistoryRows(c){
  const monthKeys = allMonthKeysFromData().sort().reverse();
  const rows = [];
  (c.branches||[]).forEach(b=>{
    monthKeys.forEach(monthKey=>{
      const target = getMonthlyTarget(b.name, monthKey);
      const monthRows = (state.orders||[]).filter(o=>o.branch===b.name && getMonthKeyFromRowDate(o.orderDate)===monthKey && o.status==="تم");
      const achieved = monthRows.reduce((s,o)=>s+Number(o.amount||0),0);
      const pct = target ? (achieved/target)*100 : 0;
      rows.push({
        branch: b.name,
        month: monthKey,
        target,
        achieved,
        pct,
        remaining: Math.max(0, target-achieved)
      });
    });
  });
  return rows;
}

function targetStatusBadge(pct,target){
  if(!target) return '<span class="target-pill target-mid">بدون هدف</span>';
  if(pct >= 100) return '<span class="target-pill target-good">تم تحقيقه</span>';
  if(pct >= 75) return '<span class="target-pill target-mid">قريب</span>';
  return '<span class="target-pill target-bad">أقل من المطلوب</span>';
}

function renderSettingsTab(c){
  const monthKeys = allMonthKeysFromData();
  const monthHeader = monthKeys.map(m=>`<th>${m}</th>`).join('');
  const targetsRows=(c.branches||[]).map(b=>{
    const monthCells = monthKeys.map(monthKey=>{
      const targetVal=getMonthlyTarget(b.name, monthKey);
      const achieved=(state.orders||[]).filter(o=>o.branch===b.name && getMonthKeyFromRowDate(o.orderDate)===monthKey && o.status==="تم").reduce((s,o)=>s+Number(o.amount||0),0);
      const pct=targetVal?(achieved/targetVal)*100:0;
      return `<td>
        <div style="display:grid;gap:6px">
          <input data-target-branch="${b.name}" data-target-month="${monthKey}" class="input" style="width:120px" value="${targetVal||''}" placeholder="تارجت">
          <div class="small muted">حقق: ${money(achieved)}</div>
          <div class="small muted">نسبة: ${pct.toFixed(1)}%</div>
        </div>
      </td>`;
    }).join('');
    return `<tr><td>${b.name}</td>${monthCells}</tr>`;
  }).join('');

  const historyRows = monthlyTargetHistoryRows(c).filter(x=>x.target || x.achieved).map(x=>`<tr><td>${x.branch}</td><td>${x.month}</td><td>${money(x.target)}</td><td>${money(x.achieved)}</td><td>${x.pct.toFixed(1)}%</td><td>${money(x.remaining)}</td><td>${targetStatusBadge(x.pct,x.target)}</td></tr>`).join('');

  return `<div class="grid2eq">
    <div class="card"><div class="section-title"><h3 style="margin:0">كلمة مرور Admin</h3></div><div class="toolbar"><input id="newPass" class="input" type="password" placeholder="كلمة مرور الأدمن الجديدة"><button class="btn btn-primary" onclick="changePassword()">حفظ</button></div><div class="small muted" style="margin-top:10px">تستخدم للدخول الكامل والتعديل.</div></div>
    <div class="card"><div class="section-title"><h3 style="margin:0">كلمة مرور Viewer</h3></div><div class="toolbar"><input id="newViewerPass" class="input" type="password" placeholder="كلمة مرور الإدارة / العرض"><button class="btn btn-primary" onclick="changeViewerPassword()">حفظ</button></div><div class="small muted" style="margin-top:10px">تستخدم للعرض فقط بدون رفع أو تعديل.</div></div>
  </div>
  <div class="card" style="margin-top:16px"><div class="section-title"><h3 style="margin:0">إعدادات التنبيه</h3></div><div class="toolbar"><input id="cancelThreshold" class="input" type="number" value="${state.settings.alertThresholds.cancelRate||10}" placeholder="حد نسبة الإلغاء"><button class="btn btn-primary" onclick="saveAlertSettings()">حفظ التنبيه</button></div></div>
  <div class="card">
    <div class="section-title"><h3 style="margin:0">أهداف مبيعات الفروع حسب الشهر</h3><button class="btn btn-primary" onclick="applyAllMonthlyTargets()">تطبيق كل الأهداف الشهرية</button></div>
    <div class="small muted" style="margin-bottom:10px">اكتب تارجت كل فرع لكل شهر ثم اضغط "تطبيق كل الأهداف الشهرية". سيبقى محفوظًا، وستشوف حقق منه كام وAchievement Rate لكل شهر.</div>
    <div class="tableWrap"><table><thead><tr><th>الفرع</th>${monthHeader}</tr></thead><tbody>${targetsRows}</tbody></table></div>
  </div>
  <div class="card" style="margin-top:16px">
    <div class="section-title"><h3 style="margin:0">سجل التارجت الشهري والتحقيق</h3></div>
    <div class="tableWrap"><table><thead><tr><th>الفرع</th><th>الشهر</th><th>التارجت</th><th>المحقق</th><th>Achievement Rate</th><th>Remaining</th><th>الحالة</th></tr></thead><tbody>${historyRows || `<tr><td colspan="7">No Data تارجت شهرية بعد.</td></tr>`}</tbody></table></div>
  </div>`;
}
function bestWorstSection(title, items, metricKey, suffix=''){
  if(!items || !items.length) return `<div class="card"><div class="section-title"><h3 style="margin:0">${title}</h3></div><div class="note">No Data</div></div>`;
  const best = [...items].sort((a,b)=>(Number(b[metricKey]||0)-Number(a[metricKey]||0))).slice(0,5);
  const worst = [...items].sort((a,b)=>(Number(a[metricKey]||0)-Number(b[metricKey]||0))).slice(0,5);
  return `<div class="grid2eq">
    <div class="card">
      <div class="section-title"><h3 style="margin:0">${title} - أفضل 5</h3></div>
      ${progressList(best.map(x=>({name:x.name,[metricKey]:x[metricKey]})),metricKey,suffix)}
    </div>
    <div class="card">
      <div class="section-title"><h3 style="margin:0">${title} - أقل 5</h3></div>
      ${progressList(worst.map(x=>({name:x.name,[metricKey]:x[metricKey]})),metricKey,suffix)}
    </div>
  </div>`;
}


function managerSummaryHtml(c){
  const bestBranch = c.branches[0];
  const weakBranch = [...c.branches].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const bestDriver = c.drivers[0];
  const weakDriver = [...c.drivers].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const bestPrep = c.preparers[0];
  const weakPlatform = [...c.platforms].sort((a,b)=>(a.completed||0)-(b.completed||0))[0];
  const topArea = c.areas[0];
  const topCustomer = c.topCustomer;
  return `
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص تنفيذي سريع</h3></div>
        <div class="note">
          • إجمالي الطلبات المكتملة: <strong>${c.total}</strong><br>
          • إجمالي المبيعات: <strong>${money(c.sales)}</strong><br>
          • متوسط الأوردر: <strong>${money(c.avg)}</strong><br>
          • نسبة الإلغاء: <strong>${c.cancelRate.toFixed(1)}%</strong><br>
          • أعلى ساعة ضغط: <strong>${c.peakHour}</strong>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص التارجت</h3></div>
        <div class="note">
          • Total Target: <strong>${money(c.targetSummary.target)}</strong><br>
          • Total Achieved: <strong>${money(c.targetSummary.sales)}</strong><br>
          • Achievement Rate: <strong>${c.targetSummary.pct.toFixed(1)}%</strong><br>
          • Remaining: <strong>${money(c.targetSummary.remaining)}</strong>
        </div>
      </div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل العناصر</h3></div>
        <div class="note">
          • أفضل فرع: <strong>${bestBranch?.name||'—'}</strong> (${(bestBranch?.score||0).toFixed(1)}%)<br>
          • أفضل محضّر: <strong>${bestPrep?.name||'—'}</strong> (${(bestPrep?.score||0).toFixed(1)}%)<br>
          • أفضل طيار: <strong>${bestDriver?.name||'—'}</strong> (${(bestDriver?.score||0).toFixed(1)}%)<br>
          • أعلى منطقة: <strong>${topArea?.name||'—'}</strong><br>
          • أعلى عميل: <strong>${topCustomer?.name||'—'}</strong> (${topCustomer?.orders||0} طلب)
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أهم نقاط المتابعة</h3></div>
        <div class="note">
          • الفرع الأضعف حاليًا: <strong>${weakBranch?.name||'—'}</strong> (${(weakBranch?.score||0).toFixed(1)}%)<br>
          • الطيار الأضعف: <strong>${weakDriver?.name||'—'}</strong> (${(weakDriver?.score||0).toFixed(1)}%)<br>
          • المنصة الأضعف: <strong>${weakPlatform?.name||'—'}</strong> (${weakPlatform?.completed||0} طلب مكتمل)<br>
          • التنبيهات الحالية: <strong>${buildColorAlerts(c).length}</strong><br>
          • آخر يوم في البيانات: <strong>${c.daily.today.date||'—'}</strong>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="section-title"><h3 style="margin:0">قرار إداري مقترح</h3></div>
      <div class="note">
        ${bestBranch ? `• دعم الفرع الأقوى حاليًا وهو <strong>${bestBranch.name}</strong> للحفاظ على الأداء.<br>` : ''}
        ${weakBranch ? `• مراجعة أسباب انخفاض أداء <strong>${weakBranch.name}</strong> وتحليل الإلغاء ووقت الذروة به.<br>` : ''}
        ${weakPlatform ? `• إعادة تقييم أداء منصة <strong>${weakPlatform.name}</strong> لأنها الأقل في الطلبات المكتملة.<br>` : ''}
        ${weakDriver ? `• متابعة أداء <strong>${weakDriver.name}</strong> تشغيليًا وتحليل أسباب التراجع.<br>` : ''}
        • استخدم هذا الملخص كصفحة سريعة قبل مراجعة التفاصيل.
      </div>
    </div>
  `;
}


function paymentAnalysisHtml(c){
  const payments = [...(c.payments||[])];
  const paymentStats = payments.map(p=>{
    const rows = (c.data||[]).filter(x=>x.paymentMethod===p.name);
    const done = rows.filter(x=>x.status==="تم");
    const sales = done.reduce((s,x)=>s+Number(x.amount||0),0);
    const canceled = rows.filter(x=>x.status==="ملغي / مرتجع").length;
    const cancelRate = rows.length ? (canceled/rows.length)*100 : 0;
    const avg = done.length ? sales/done.length : 0;
    return {name:p.name, orders:rows.length, completed:done.length, canceled, sales, cancelRate, avg};
  }).sort((a,b)=>b.completed-a.completed||b.sales-a.sales);

  const best = paymentStats[0];
  const worst = [...paymentStats].sort((a,b)=>a.completed-b.completed||a.sales-b.sales)[0];

  return `
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص طرق الدفع</h3></div>
        <div class="note">
          • عدد طرق الدفع المستخدمة: <strong>${paymentStats.length}</strong><br>
          • أفضل طريقة دفع: <strong>${best?.name||'—'}</strong><br>
          • الأقل استخدامًا: <strong>${worst?.name||'—'}</strong><br>
          • إجمالي الطلبات المكتملة عبر الدفع: <strong>${paymentStats.reduce((s,x)=>s+x.completed,0)}</strong>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">رسم طرق الدفع</h3></div>
        ${renderBars(paymentStats.slice(0,10),'name','completed',' طلب')}
      </div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل 5 طرق دفع</h3></div>
        ${progressList(paymentStats.slice(0,5),'completed',' طلب')}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أقل 5 طرق دفع</h3></div>
        ${progressList([...paymentStats].sort((a,b)=>a.completed-b.completed).slice(0,5),'completed',' طلب')}
      </div>
    </div>
    <div class="card">
      <div class="section-title"><h3 style="margin:0">تفصيل طرق الدفع</h3></div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>طريقة الدفع</th>
              <th>إجمالي الطلبات</th>
              <th>المكتمل</th>
              <th>الملغي / المرتجع</th>
              <th>إجمالي المبيعات</th>
              <th>متوسط الأوردر</th>
              <th>نسبة الإلغاء</th>
            </tr>
          </thead>
          <tbody>
            ${paymentStats.map(p=>`<tr><td>${p.name||''}</td><td>${p.orders}</td><td>${p.completed}</td><td>${p.canceled}</td><td>${money(p.sales)}</td><td>${money(p.avg)}</td><td>${p.cancelRate.toFixed(1)}%</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}



function buildSelectedDayAlerts(selected, days){
  if(!selected) return [];
  const alerts = [];
  const compareDays = days.filter(d=>d.date!==selected.date);
  const avgSales = compareDays.length ? compareDays.reduce((s,d)=>s+Number(d.sales||0),0)/compareDays.length : Number(selected.sales||0);
  const avgCancel = compareDays.length ? compareDays.reduce((s,d)=>s+Number(d.cancelRate||0),0)/compareDays.length : Number(selected.cancelRate||0);
  const avgCompleted = compareDays.length ? compareDays.reduce((s,d)=>s+Number(d.completed||0),0)/compareDays.length : Number(selected.completed||0);

  if((selected.cancelRate||0) > Math.max(10, avgCancel + 3)){
    alerts.push({level:'red', title:'إلغاء مرتفع في اليوم المختار', text:`نسبة الإلغاء ${selected.cancelRate.toFixed(1)}% وهي أعلى من المعتاد.`});
  } else if((selected.cancelRate||0) > Math.max(5, avgCancel)){
    alerts.push({level:'yellow', title:'إلغاء يحتاج متابعة', text:`نسبة الإلغاء ${selected.cancelRate.toFixed(1)}% أعلى قليلًا من المتوسط.`});
  } else {
    alerts.push({level:'green', title:'الإلغاء جيد', text:`نسبة الإلغاء ${selected.cancelRate.toFixed(1)}% في حدود آمنة.`});
  }

  if((selected.sales||0) < avgSales * 0.75){
    alerts.push({level:'red', title:'مبيعات منخفضة', text:`مبيعات اليوم المختار أقل بوضوح من المتوسط العام.`});
  } else if((selected.sales||0) < avgSales){
    alerts.push({level:'yellow', title:'مبيعات أقل من المتوسط', text:`مبيعات اليوم المختار أقل قليلًا من متوسط الأيام الأخرى.`});
  } else {
    alerts.push({level:'green', title:'مبيعات جيدة', text:`مبيعات اليوم المختار مساوية أو أعلى من المتوسط.`});
  }

  if((selected.completed||0) < avgCompleted * 0.7){
    alerts.push({level:'red', title:'حجم تنفيذ منخفض', text:`عدد الطلبات المكتملة في هذا اليوم أقل من المعتاد.`});
  } else if((selected.completed||0) < avgCompleted){
    alerts.push({level:'yellow', title:'تنفيذ أقل من المتوسط', text:`الطلبات المكتملة أقل قليلًا من متوسط الأيام الأخرى.`});
  } else {
    alerts.push({level:'green', title:'تنفيذ جيد', text:`عدد الطلبات المكتملة في مستوى جيد.`});
  }

  if(selected.topBranch && selected.topBranch !== '—'){
    alerts.push({level:'green', title:'أفضل فرع في اليوم المختار', text:`الفرع الأقوى كان ${selected.topBranch}.`});
  }
  if(selected.topPlatform && selected.topPlatform !== '—'){
    alerts.push({level:'yellow', title:'المنصة الأبرز في اليوم المختار', text:`أكثر منصة مساهمة كانت ${selected.topPlatform}.`});
  }

  return alerts.slice(0,5);
}


function selectedDayBranchComparisonHtml(selectedDate, dataRows){
  const rows = (dataRows||[]).filter(o=>o.orderDate===selectedDate);
  const branchMap = {};
  rows.forEach(o=>{
    if(!o.branch) return;
    if(!branchMap[o.branch]) branchMap[o.branch] = {name:o.branch, orders:0, completed:0, canceled:0, sales:0};
    const b = branchMap[o.branch];
    b.orders++;
    if(o.status==="تم"){ b.completed++; b.sales += Number(o.amount||0); }
    else if(o.status==="ملغي / مرتجع"){ b.canceled++; }
  });
  const items = Object.values(branchMap).map(b=>{
    const cancelRate = b.orders ? (b.canceled/b.orders)*100 : 0;
    const avg = b.completed ? b.sales/b.completed : 0;
    return {...b, cancelRate, avg};
  }).sort((a,b)=>b.completed-a.completed || b.sales-a.sales);

  if(!items.length){
    return '<div class="card"><div class="section-title"><h3 style="margin:0">مقارنة الفروع في اليوم المختار</h3></div><div class="note">No Data فروع في هذا اليوم.</div></div>';
  }

  const best = items[0];
  const worst = [...items].sort((a,b)=>a.completed-b.completed || a.sales-b.sales)[0];

  return `
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل وأسوأ فرع في اليوم المختار</h3></div>
        <div class="note">
          • أفضل فرع: <strong>${best?.name||'—'}</strong><br>
          • طلبات مكتملة: <strong>${best?.completed||0}</strong><br>
          • مبيعات: <strong>${money(best?.sales||0)}</strong><br><br>
          • الفرع الأضعف: <strong>${worst?.name||'—'}</strong><br>
          • طلبات مكتملة: <strong>${worst?.completed||0}</strong><br>
          • مبيعات: <strong>${money(worst?.sales||0)}</strong>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">الطلبات المكتملة حسب الفرع</h3></div>
        ${renderBars(items,'name','completed',' طلب')}
      </div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">مبيعات الفروع في اليوم المختار</h3></div>
        ${renderBars(items,'name','sales')}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">نسبة الإلغاء حسب الفرع</h3></div>
        ${renderBars(items,'name','cancelRate','%')}
      </div>
    </div>
    <div class="card">
      <div class="section-title"><h3 style="margin:0">الجدول التفصيلي للفروع في اليوم المختار</h3></div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>الفرع</th>
              <th>إجمالي الطلبات</th>
              <th>المكتمل</th>
              <th>الملغي / المرتجع</th>
              <th>المبيعات</th>
              <th>متوسط الأوردر</th>
              <th>نسبة الإلغاء</th>
            </tr>
          </thead>
          <tbody>
            ${items.map(b=>`<tr><td>${b.name}</td><td>${b.orders}</td><td>${b.completed}</td><td>${b.canceled}</td><td>${money(b.sales)}</td><td>${money(b.avg)}</td><td>${b.cancelRate.toFixed(1)}%</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}


function getWeekStart(dateStr){
  const d = new Date(dateStr + 'T00:00:00');
  if(isNaN(d)) return '';
  const day = d.getDay(); // 0 Sun ... 6 Sat
  const diff = (day === 0 ? -6 : 1 - day); // Monday start
  d.setDate(d.getDate() + diff);
  const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), da = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${da}`;
}
function formatWeekLabel(weekStart){
  if(!weekStart) return '—';
  const start = new Date(weekStart + 'T00:00:00');
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  return `${fmt(start)} → ${fmt(end)}`;
}
function groupByPeriod(rows, mode){
  const map = {};
  rows.forEach(o=>{
    let key = '';
    if(mode==='weekly'){
      key = getWeekStart(o.orderDate||'');
    } else {
      key = (o.orderDate||'').slice(0,7);
    }
    if(!key) return;
    if(!map[key]) map[key] = {key, orders:0, completed:0, canceled:0, sales:0, branches:{}, platforms:{}, drivers:{}, preparers:{}};
    const g = map[key];
    g.orders++;
    if(o.branch) g.branches[o.branch] = (g.branches[o.branch]||0)+1;
    if(o.platform) g.platforms[o.platform] = (g.platforms[o.platform]||0)+1;
    if(o.driver) g.drivers[o.driver] = (g.drivers[o.driver]||0)+1;
    if(o.preparer) g.preparers[o.preparer] = (g.preparers[o.preparer]||0)+1;
    if(o.status==="تم"){ g.completed++; g.sales += Number(o.amount||0); }
    else if(o.status==="ملغي / مرتجع"){ g.canceled++; }
  });
  return Object.values(map).map(g=>{
    const avg = g.completed ? g.sales/g.completed : 0;
    const cancelRate = g.orders ? (g.canceled/g.orders)*100 : 0;
    const topBranch = Object.entries(g.branches).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topPlatform = Object.entries(g.platforms).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topDriver = Object.entries(g.drivers).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topPreparer = Object.entries(g.preparers).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    return {...g, avg, cancelRate, topBranch, topPlatform, topDriver, topPreparer, label: mode==='weekly' ? formatWeekLabel(g.key) : g.key};
  }).sort((a,b)=>b.key.localeCompare(a.key));
}
function setPeriodRange(fromStr,toStr){
  state.periodFrom = fromStr || '';
  state.periodTo = toStr || '';
  render();
}
function buildPeriodAlerts(selected, others, labelWord){
  const avgSales = others.length ? others.reduce((s,x)=>s+Number(x.sales||0),0)/others.length : Number(selected.sales||0);
  const avgCompleted = others.length ? others.reduce((s,x)=>s+Number(x.completed||0),0)/others.length : Number(selected.completed||0);
  const avgCancel = others.length ? others.reduce((s,x)=>s+Number(x.cancelRate||0),0)/others.length : Number(selected.cancelRate||0);
  return [
    (selected.cancelRate > Math.max(10, avgCancel+2))
      ? {level:'red', title:`إلغاء ${labelWord} مرتفع`, text:`نسبة الإلغاء ${selected.cancelRate.toFixed(1)}% أعلى من المعتاد.`}
      : {level:'green', title:`الإلغاء ${labelWord} جيد`, text:`نسبة الإلغاء ${selected.cancelRate.toFixed(1)}% في الحدود الطبيعية.`},
    (selected.sales < avgSales*0.8)
      ? {level:'yellow', title:`مبيعات ${labelWord} أقل من المتوسط`, text:`المبيعات أقل من متوسط باقي الفترات.`}
      : {level:'green', title:`مبيعات ${labelWord} جيدة`, text:`المبيعات مساوية أو أعلى من متوسط باقي الفترات.`},
    (selected.completed < avgCompleted*0.8)
      ? {level:'yellow', title:`تنفيذ ${labelWord} أقل من المتوسط`, text:`الطلبات المكتملة أقل من متوسط باقي الفترات.`}
      : {level:'green', title:`تنفيذ ${labelWord} جيد`, text:`الطلبات المكتملة في مستوى جيد.`}
  ];
}
function periodAnalysisHtml(c){
  const rows = c.data || [];
  const from = state.periodFrom || '';
  const to = state.periodTo || '';
  const filteredRows = rows.filter(o=>{
    const d = o.orderDate || '';
    if(!d) return false;
    if(from && d < from) return false;
    if(to && d > to) return false;
    return true;
  });

  if(!filteredRows.length){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل الفترات</h3></div><div class="note">اختر فترة تحتوي على بيانات لعرض التحليل.</div></div>`;
  }

  const mode = from && to && from.slice(0,7)===to.slice(0,7) ? 'weekly' : 'monthly';
  const labelWord = mode==='weekly' ? 'أسبوعي' : 'شهري';
  const groups = groupByPeriod(filteredRows, mode);
  const selected = groups[0];
  const others = groups.slice(1);

  return `
    <div class="card no-print" style="margin-bottom:16px">
      <div class="section-title"><h3 style="margin:0">تحليل الفترات</h3></div>
      <div class="toolbar">
        <input class="input" type="date" id="periodFrom" value="${from}">
        <input class="input" type="date" id="periodTo" value="${to}">
        <button class="btn btn-primary" onclick="setPeriodRange(document.getElementById('periodFrom').value, document.getElementById('periodTo').value)">تطبيق الفترة</button>
        <button class="btn btn-outline" onclick="setPeriodRange('', '')">مسح الفترة</button>
      </div>
      <div class="small muted" style="margin-top:8px">اختار من وإلى، ولو الفترة داخل نفس الشهر هيظهر لك التحليل بشكل أسبوعي، ولو أكبر هيظهر شهري.</div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص الفترة الحالية</h3></div>
        <div class="note">
          • نوع التجميع: <strong>${mode==='weekly' ? 'أسبوعي' : 'شهري'}</strong><br>
          • عدد الفترات داخل النطاق: <strong>${groups.length}</strong><br>
          • إجمالي الطلبات: <strong>${filteredRows.length}</strong><br>
          • إجمالي الطلبات المكتملة: <strong>${filteredRows.filter(x=>x.status==="تم").length}</strong><br>
          • إجمالي المبيعات: <strong>${money(filteredRows.filter(x=>x.status==="تم").reduce((s,x)=>s+Number(x.amount||0),0))}</strong><br>
          • أفضل فرع: <strong>${selected.topBranch}</strong><br>
          • أفضل منصة: <strong>${selected.topPlatform}</strong>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص ${labelWord}</h3></div>
        <div class="note">
          • الفترة الأحدث: <strong>${selected.label}</strong><br>
          • الطلبات: <strong>${selected.orders}</strong><br>
          • المكتمل: <strong>${selected.completed}</strong><br>
          • المبيعات: <strong>${money(selected.sales)}</strong><br>
          • متوسط الأوردر: <strong>${money(selected.avg)}</strong><br>
          • نسبة الإلغاء: <strong>${selected.cancelRate.toFixed(1)}%</strong>
        </div>
      </div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">تنبيهات ${labelWord}</h3></div>
        ${alertsHtml(buildPeriodAlerts(selected, others, labelWord))}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل عناصر الفترة الأحدث</h3></div>
        <div class="note">
          • أفضل فرع: <strong>${selected.topBranch}</strong><br>
          • أفضل منصة: <strong>${selected.topPlatform}</strong><br>
          • أفضل طيار: <strong>${selected.topDriver}</strong><br>
          • أفضل محضّر: <strong>${selected.topPreparer}</strong>
        </div>
      </div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">المبيعات حسب ${mode==='weekly' ? 'الأسابيع' : 'الشهور'}</h3></div>
        ${renderBars(groups.slice(0,12).map(g=>({name:g.label,sales:g.sales})),'name','sales')}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">الطلبات المكتملة حسب ${mode==='weekly' ? 'الأسابيع' : 'الشهور'}</h3></div>
        ${renderBars(groups.slice(0,12).map(g=>({name:g.label,completed:g.completed})),'name','completed',' طلب')}
      </div>
    </div>

    <div class="card">
      <div class="section-title"><h3 style="margin:0">الجدول التفصيلي</h3></div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>${mode==='weekly' ? 'الأسبوع' : 'الشهر'}</th>
              <th>إجمالي الطلبات</th>
              <th>المكتمل</th>
              <th>الملغي / المرتجع</th>
              <th>المبيعات</th>
              <th>متوسط الأوردر</th>
              <th>نسبة الإلغاء</th>
              <th>أفضل فرع</th>
              <th>أفضل منصة</th>
              <th>أفضل طيار</th>
              <th>أفضل محضّر</th>
            </tr>
          </thead>
          <tbody>
            ${groups.map(g=>`<tr><td>${g.label}</td><td>${g.orders}</td><td>${g.completed}</td><td>${g.canceled}</td><td>${money(g.sales)}</td><td>${money(g.avg)}</td><td>${g.cancelRate.toFixed(1)}%</td><td>${g.topBranch}</td><td>${g.topPlatform}</td><td>${g.topDriver}</td><td>${g.topPreparer}</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function dailyReportHtml(c){
  const byDate = {};
  (c.data||[]).forEach(o=>{
    const d = o.orderDate || '';
    if(!d) return;
    if(!byDate[d]) byDate[d] = {date:d, orders:0, completed:0, canceled:0, sales:0, branches:{}, platforms:{}, drivers:{}, preparers:{}};
    byDate[d].orders++;
    if(o.branch) byDate[d].branches[o.branch] = (byDate[d].branches[o.branch]||0)+1;
    if(o.platform) byDate[d].platforms[o.platform] = (byDate[d].platforms[o.platform]||0)+1;
    if(o.driver) byDate[d].drivers[o.driver] = (byDate[d].drivers[o.driver]||0)+1;
    if(o.preparer) byDate[d].preparers[o.preparer] = (byDate[d].preparers[o.preparer]||0)+1;
    if(o.status==="تم"){ byDate[d].completed++; byDate[d].sales += Number(o.amount||0); }
    else if(o.status==="ملغي / مرتجع"){ byDate[d].canceled++; }
  });
  const days = Object.values(byDate).map(d=>{
    const avg = d.completed ? d.sales/d.completed : 0;
    const cancelRate = d.orders ? (d.canceled/d.orders)*100 : 0;
    const topBranch = Object.entries(d.branches).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topPlatform = Object.entries(d.platforms).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topDriver = Object.entries(d.drivers).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    const topPreparer = Object.entries(d.preparers).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
    return {...d, avg, cancelRate, topBranch, topPlatform, topDriver, topPreparer};
  }).sort((a,b)=>b.date.localeCompare(a.date));

  if(!days.length){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">تقرير يومي منفصل</h3></div><div class="note">No Data يومية متاحة.</div></div>`;
  }

  const selectedDate = state.dailySelectedDate && days.some(x=>x.date===state.dailySelectedDate) ? state.dailySelectedDate : days[0].date;
  const selected = days.find(x=>x.date===selectedDate) || days[0];

  return `
    <div class="card no-print" style="margin-bottom:16px">
      <div class="section-title"><h3 style="margin:0">اختيار يوم محدد</h3></div>
      <div class="toolbar">
        <select class="select" id="dailyDateSelect" onchange="setDailyDate(this.value)">
          ${days.map(d=>`<option value="${d.date}" ${d.date===selectedDate?'selected':''}>${d.date}</option>`).join('')}
        </select>
        <button class="btn btn-outline" onclick="setDailyDate('${days[0].date}')">آخر يوم</button>
      </div>
      <div class="small muted" style="margin-top:8px">اختر أي يوم من الأيام الموجودة في الداتا لعرض تقريره بشكل منفصل.</div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص اليوم المختار</h3></div>
        <div class="note">
          • التاريخ: <strong>${selected?.date||'—'}</strong><br>
          • الطلبات: <strong>${selected?.orders||0}</strong><br>
          • المكتمل: <strong>${selected?.completed||0}</strong><br>
          • الملغي / المرتجع: <strong>${selected?.canceled||0}</strong><br>
          • المبيعات: <strong>${money(selected?.sales||0)}</strong><br>
          • متوسط الأوردر: <strong>${money(selected?.avg||0)}</strong><br>
          • نسبة الإلغاء: <strong>${(selected?.cancelRate||0).toFixed(1)}%</strong>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل عناصر اليوم المختار</h3></div>
        <div class="note">
          • أفضل فرع: <strong>${selected?.topBranch||'—'}</strong><br>
          • أفضل منصة: <strong>${selected?.topPlatform||'—'}</strong><br>
          • أفضل طيار: <strong>${selected?.topDriver||'—'}</strong><br>
          • أفضل محضّر: <strong>${selected?.topPreparer||'—'}</strong>
        </div>
      </div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">التنبيهات الذكية لليوم المختار</h3></div>
        ${alertsHtml(buildSelectedDayAlerts(selected, days))}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">مقارنة اليوم المختار بالمتوسط</h3></div>
        <div class="note">
          • مبيعات اليوم: <strong>${money(selected?.sales||0)}</strong><br>
          • متوسط مبيعات باقي الأيام: <strong>${money(days.filter(d=>d.date!==selectedDate).length ? days.filter(d=>d.date!==selectedDate).reduce((s,d)=>s+Number(d.sales||0),0)/days.filter(d=>d.date!==selectedDate).length : (selected?.sales||0))}</strong><br>
          • مكتمل اليوم: <strong>${selected?.completed||0}</strong><br>
          • متوسط المكتمل لباقي الأيام: <strong>${(days.filter(d=>d.date!==selectedDate).length ? days.filter(d=>d.date!==selectedDate).reduce((s,d)=>s+Number(d.completed||0),0)/days.filter(d=>d.date!==selectedDate).length : (selected?.completed||0)).toFixed(1)}</strong>
        </div>
      </div>
    </div>
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">مبيعات الأيام</h3></div>
        ${renderBars(days.slice(0,10),'date','sales')}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">الطلبات المكتملة اليومية</h3></div>
        ${renderBars(days.slice(0,10),'date','completed',' طلب')}
      </div>
    </div>
    <div class="card">
      <div class="section-title"><h3 style="margin:0">التقرير اليومي التفصيلي</h3></div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>إجمالي الطلبات</th>
              <th>المكتمل</th>
              <th>الملغي / المرتجع</th>
              <th>المبيعات</th>
              <th>متوسط الأوردر</th>
              <th>نسبة الإلغاء</th>
              <th>أفضل فرع</th>
              <th>أفضل منصة</th>
              <th>أفضل طيار</th>
              <th>أفضل محضّر</th>
            </tr>
          </thead>
          <tbody>
            ${days.map(d=>`<tr ${d.date===selectedDate?'style="background:rgba(255,255,255,.05)"':''}><td>${d.date}</td><td>${d.orders}</td><td>${d.completed}</td><td>${d.canceled}</td><td>${money(d.sales)}</td><td>${money(d.avg)}</td><td>${d.cancelRate.toFixed(1)}%</td><td>${d.topBranch}</td><td>${d.topPlatform}</td><td>${d.topDriver}</td><td>${d.topPreparer}</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
    ${selectedDayBranchComparisonHtml(selectedDate, c.data)}

  `;
}
function setDailyDate(dateStr){
  state.dailySelectedDate = dateStr || '';
  render();
}


function buildBranchDailyHourlyData(rows){
  const branchMap = {};
  (rows||[]).forEach(o=>{
    if(!o.branch || !o.orderDate) return;
    const hour = hourLabel(o.orderTime);
    if(!hour || hour===':00') return;
    if(!branchMap[o.branch]) branchMap[o.branch] = {};
    if(!branchMap[o.branch][o.orderDate]) branchMap[o.branch][o.orderDate] = {};
    branchMap[o.branch][o.orderDate][hour] = (branchMap[o.branch][o.orderDate][hour] || 0) + 1;
  });
  return branchMap;
}
function setBranchPeakSelected(branch){
  state.branchPeakSelected = branch || '';
  render();
}
function branchPeakByDayHtml(rows){
  const branchMap = buildBranchDailyHourlyData(rows);
  const allBranches = Object.keys(branchMap).sort();
  if(!allBranches.length){
    return '<div class="note">No Data كافية لحساب وقت الذروة.</div>';
  }

  const options = ['كل الفروع', ...allBranches];
  const selectedBranch = state.branchPeakSelected && options.includes(state.branchPeakSelected) ? state.branchPeakSelected : 'كل الفروع';

  let dailyMap = {};
  if(selectedBranch === 'كل الفروع'){
    (rows||[]).forEach(o=>{
      if(!o.orderDate) return;
      const hour = hourLabel(o.orderTime);
      if(!hour || hour===':00') return;
      if(!dailyMap[o.orderDate]) dailyMap[o.orderDate] = {};
      dailyMap[o.orderDate][hour] = (dailyMap[o.orderDate][hour] || 0) + 1;
    });
  } else {
    dailyMap = branchMap[selectedBranch] || {};
  }

  const days = Object.keys(dailyMap).sort().reverse();
  const hours = [...new Set(days.flatMap(d=>Object.keys(dailyMap[d]||{})))].sort((a,b)=>a.localeCompare(b));

  const selectedSummary = days.map(day=>{
    const entries = Object.entries(dailyMap[day] || {});
    const peak = entries.sort((a,b)=>b[1]-a[1] || a[0].localeCompare(b[0]))[0] || ['—',0];
    return {day, peakHour: peak[0], peakCount: peak[1]};
  });

  const hourTotals = hours.map(hour=>{
    const count = days.reduce((sum, day)=>sum + Number((dailyMap[day] || {})[hour] || 0), 0);
    return {hour, count};
  }).filter(x=>x.count > 0).sort((a,b)=>a.hour.localeCompare(b.hour));

  const summaryRows = allBranches.map(branch=>{
    const daysMap = branchMap[branch] || {};
    const branchDays = Object.keys(daysMap).sort().reverse();
    const peaks = branchDays.map(day=>{
      const entries = Object.entries(daysMap[day] || {});
      const peak = entries.sort((a,b)=>b[1]-a[1] || a[0].localeCompare(b[0]))[0] || ['—',0];
      return {day, peakHour: peak[0], peakCount: peak[1]};
    });
    const best = peaks.sort((a,b)=>b.peakCount-a.peakCount || a.day.localeCompare(b.day))[0] || {day:'—', peakHour:'—', peakCount:0};
    return {branch, day:best.day, peakHour:best.peakHour, peakCount:best.peakCount};
  }).sort((a,b)=>b.peakCount-a.peakCount || a.branch.localeCompare(b.branch));

  const bestSelectedPeak = selectedSummary.slice().sort((a,b)=>b.peakCount-a.peakCount || a.day.localeCompare(b.day))[0] || {day:'—', peakHour:'—', peakCount:0};
  const totalOrders = hourTotals.reduce((s,x)=>s + x.count, 0);
  const bestHour = hourTotals.slice().sort((a,b)=>b.count-a.count || a.hour.localeCompare(b.hour))[0] || {hour:'—', count:0};
  const avgDailyPeak = selectedSummary.length ? (selectedSummary.reduce((s,x)=>s+x.peakCount,0) / selectedSummary.length) : 0;
  const repeatedPeakHour = (() => {
    const freq = {};
    selectedSummary.forEach(x => { if(x.peakHour && x.peakHour !== '—') freq[x.peakHour] = (freq[x.peakHour] || 0) + 1; });
    const entries = Object.entries(freq).sort((a,b)=>b[1]-a[1] || a[0].localeCompare(b[0]));
    return entries[0] ? {hour: entries[0][0], days: entries[0][1]} : {hour:'—', days:0};
  })();
  const topHours = hourTotals.slice().sort((a,b)=>b.count-a.count || a.hour.localeCompare(b.hour)).slice(0,5);
  const topDays = selectedSummary.slice().sort((a,b)=>b.peakCount-a.peakCount || a.day.localeCompare(b.day)).slice(0,5);

  return `
    <div class="card" style="margin-bottom:16px">
      <div class="section-title"><h3 style="margin:0">اختيار الفرع</h3></div>
      <div class="toolbar">
        <select class="select" onchange="setBranchPeakSelected(this.value)">
          ${options.map(b=>`<option value="${b}" ${b===selectedBranch?'selected':''}>${b}</option>`).join('')}
        </select>
      </div>
    </div>

    <div class="peak-kpi-grid">
      <div class="peak-kpi"><div class="label">الفرع المختار</div><div class="value">${selectedBranch}</div></div>
      <div class="peak-kpi"><div class="label">عدد الأيام</div><div class="value">${days.length}</div></div>
      <div class="peak-kpi"><div class="label">أعلى يوم ذروة</div><div class="value">${bestSelectedPeak.day}</div><div class="sub">${bestSelectedPeak.peakHour} — ${bestSelectedPeak.peakCount} طلب</div></div>
      <div class="peak-kpi"><div class="label">أفضل ساعة إجمالاً</div><div class="value">${bestHour.hour}</div><div class="sub">${bestHour.count} طلب</div></div>
      <div class="peak-kpi"><div class="label">إجمالي الطلبات بالساعات</div><div class="value">${totalOrders}</div></div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">رسم ذروة كل يوم</h3></div>
        ${renderLineChart(selectedSummary.map(x=>({name:x.day, value:x.peakCount})), 'name', 'value', ' طلب')}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">رسم الطلبات حسب الساعة</h3></div>
        ${renderColumnChart(hourTotals.map(x=>({name:x.hour, value:x.count})), 'name', 'value', ' طلب')}
      </div>
    </div>

    <div class="peak-report-grid">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">تقرير مختصر لوقت الذروة</h3></div>
        <div class="report-list">
          <div class="report-item"><div class="t">أعلى يوم ذروة</div><div class="v">${bestSelectedPeak.day}</div><div class="sub">${bestSelectedPeak.peakHour} — ${bestSelectedPeak.peakCount} طلب</div></div>
          <div class="report-item"><div class="t">متوسط الذروة اليومية</div><div class="v">${avgDailyPeak.toFixed(1)} طلب</div></div>
          <div class="report-item"><div class="t">الساعة الأكثر تكرارًا كوقت ذروة</div><div class="v">${repeatedPeakHour.hour}</div><div class="sub">تكررت في ${repeatedPeakHour.days} يوم</div></div>
          <div class="report-item"><div class="t">أفضل 5 ساعات ضغط</div><div class="sub">${topHours.map(x=>`${x.hour} (${x.count})`).join(' — ') || 'No Data'}</div></div>
          <div class="report-item"><div class="t">أفضل 5 أيام من حيث الذروة</div><div class="sub">${topDays.map(x=>`${x.day} (${x.peakCount})`).join(' — ') || 'No Data'}</div></div>
        </div>
      </div>

      <div class="card">
        <div class="section-title"><h3 style="margin:0">تحليل مرئي مختصر</h3></div>
        ${renderColumnChart(topHours.map(x=>({name:x.hour, value:x.count})), 'name', 'value', ' طلب')}
        <div class="small muted" style="margin-top:8px">يعرض الرسم أعلى 5 ساعات ضغط ${selectedBranch === 'كل الفروع' ? 'لكل الفروع معًا' : 'للفرع المختار'}.</div>
      </div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">جدول ذروة الأيام ${selectedBranch === 'كل الفروع' ? 'لكل الفروع' : 'للفرع المختار'}</h3></div>
        <div class="tableWrap">
          <table>
            <thead><tr><th>اليوم</th><th>وقت الذروة</th><th>عدد الطلبات</th></tr></thead>
            <tbody>
              ${selectedSummary.map(r=>`<tr><td class="num-cell">${r.day}</td><td class="num-cell">${r.peakHour}</td><td class="num-cell">${r.peakCount}</td></tr>`).join('') || '<tr><td colspan="3">No Data</td></tr>'}
            </tbody>
          </table>
        </div>
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل ذروة لكل فرع</h3></div>
        <div class="tableWrap">
          <table>
            <thead><tr><th>الفرع</th><th>اليوم</th><th>وقت الذروة</th><th>عدد الطلبات</th></tr></thead>
            <tbody>
              ${summaryRows.map(r=>`<tr><td>${r.branch}</td><td class="num-cell">${r.day}</td><td class="num-cell">${r.peakHour}</td><td class="num-cell">${r.peakCount}</td></tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}



function to12HourLabel(h){
  if(!h || typeof h !== 'string' || !h.includes(':')) return h || '—';
  const hour24 = parseInt(h.split(':')[0], 10);
  if(isNaN(hour24)) return h;
  const suffix = hour24 >= 12 ? 'PM' : 'AM';
  let hour12 = hour24 % 12;
  if(hour12 === 0) hour12 = 12;
  return `${String(hour12).padStart(2,'0')}:00 ${suffix}`;
}

function branchPeaksSelectedData(rows){
  const branchMap = buildBranchDailyHourlyData(rows);
  const allBranches = Object.keys(branchMap).sort();
  const options = ['كل الفروع', ...allBranches];
  const selectedBranch = state.branchPeakSelected && options.includes(state.branchPeakSelected) ? state.branchPeakSelected : 'كل الفروع';

  let dailyMap = {};
  if(selectedBranch === 'كل الفروع'){
    (rows||[]).forEach(o=>{
      if(!o.orderDate) return;
      const hour = hourLabel(o.orderTime);
      if(!hour || hour===':00') return;
      if(!dailyMap[o.orderDate]) dailyMap[o.orderDate] = {};
      dailyMap[o.orderDate][hour] = (dailyMap[o.orderDate][hour] || 0) + 1;
    });
  } else {
    dailyMap = branchMap[selectedBranch] || {};
  }

  const allHours = Array.from({length:24}, (_,i)=>String(i).padStart(2,'0')+':00');
  const dayRows = Object.entries(dailyMap).map(([date,hoursObj])=>{
    const entries = Object.entries(hoursObj||{});
    const total = entries.reduce((s,[,v])=>s+Number(v||0),0);
    const peak = entries.sort((a,b)=>b[1]-a[1])[0] || ['—',0];
    return {date,total,peakHour:peak[0],peakCount:Number(peak[1]||0),hoursObj};
  }).sort((a,b)=>a.date.localeCompare(b.date));

  const hourTotals = {};
  dayRows.forEach(d=>{
    allHours.forEach(h=>hourTotals[h]=(hourTotals[h]||0)+Number(d.hoursObj[h]||0));
  });

  const overallPeakHourEntry = Object.entries(hourTotals).sort((a,b)=>b[1]-a[1])[0] || ['—',0];
  const topDayByTotal = [...dayRows].sort((a,b)=>b.total-a.total)[0] || null;
  const topDayByPeak = [...dayRows].sort((a,b)=>b.peakCount-a.peakCount)[0] || null;
  const avgDaily = dayRows.length ? dayRows.reduce((s,d)=>s+d.total,0)/dayRows.length : 0;
  const totalOrders = dayRows.reduce((s,d)=>s+d.total,0);

  return {
    allBranches, options, selectedBranch, dayRows, hourTotals, allHours,
    overallPeakHour: overallPeakHourEntry[0],
    overallPeakCount: Number(overallPeakHourEntry[1]||0),
    topDayByTotal, topDayByPeak, avgDaily, totalOrders
  };
}

function branchPeaksProfessionalHtml(rows){
  const p = branchPeaksSelectedData(rows);
  if(!p.dayRows.length){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">ذروة الفروع اليومية</h3></div><div class="branchpeaks-empty">No Data كافية لعرض تحليل الذروة.</div></div>`;
  }

  const topHours = p.allHours.map(h=>({label:h,value:p.hourTotals[h]||0})).filter(x=>x.value>0);
  const maxHourValue = Math.max(1, ...topHours.map(x=>x.value));
  const topDays = [...p.dayRows].sort((a,b)=>b.total-a.total).slice(0,10);
  const maxDayValue = Math.max(1, ...topDays.map(x=>x.total));
  const recentDays = [...p.dayRows].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,14);

  return `
  <div class="branchpeaks-wrap">
    <div class="card">
      <div class="branchpeaks-head">
        <div class="branchpeaks-title">
          <h3>تحليل ذروة الفروع</h3>
          <p class="note">تابع أعلى ساعات الضغط وأقوى الأيام للفرع المختار بشكل واضح ومنظم.</p>
        </div>
        <div class="branchpeaks-filter">
          <div class="small muted" style="margin-bottom:6px">اختيار الفرع</div>
          <select class="select" onchange="setBranchPeakSelected(this.value)">
            ${p.options.map(b=>`<option value="${b}" ${b===p.selectedBranch?'selected':''}>${b}</option>`).join('')}
          </select>
        </div>
      </div>
    </div>

    <div class="branchpeaks-kpis">
      <div class="branchpeaks-kpi">
        <div class="k">الفرع المختار</div>
        <div class="v">${p.selectedBranch}</div>
        <div class="s">يمكن اختيار كل الفروع أو فرع محدد</div>
      </div>
      <div class="branchpeaks-kpi">
        <div class="k">إجمالي الطلبات بالفترة</div>
        <div class="v">${p.totalOrders}</div>
        <div class="s">إجمالي الطلبات المستخدمة في التحليل</div>
      </div>
      <div class="branchpeaks-kpi">
        <div class="k">أقوى ساعة إجمالًا</div>
        <div class="v">${to12HourLabel(p.overallPeakHour)}</div>
        <div class="s">${p.overallPeakCount} طلب خلال هذه الساعة</div>
      </div>
      <div class="branchpeaks-kpi">
        <div class="k">أعلى يوم ضغط</div>
        <div class="v">${p.topDayByTotal ? p.topDayByTotal.date : '—'}</div>
        <div class="s">${p.topDayByTotal ? `${p.topDayByTotal.total} طلب` : '—'}</div>
      </div>
      <div class="branchpeaks-kpi">
        <div class="k">متوسط الطلبات اليومي</div>
        <div class="v">${p.avgDaily.toFixed(1)}</div>
        <div class="s">${p.dayRows.length} يوم داخل التحليل</div>
      </div>
    </div>

    <div class="branchpeaks-grid">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">توزيع الطلبات حسب الساعة</h3></div><div class="note" style="margin-bottom:12px">هذا الرسم يوضح إجمالي الطلبات في كل ساعة خلال كامل الفترة المختارة، وليس ليوم واحد فقط.</div>
        <div class="branchpeaks-bars">
          ${topHours.map(h=>`
            <div class="branchpeaks-row">
              <div class="branchpeaks-label">${to12HourLabel(h.label)}</div>
              <div class="branchpeaks-track"><div class="branchpeaks-fill" style="width:${(h.value/maxHourValue)*100}%"></div></div>
              <div class="branchpeaks-num">${h.value}</div>
            </div>`).join('')}
        </div>
      </div>

      <div class="card">
        <div class="section-title"><h3 style="margin:0">أعلى الأيام خلال الفترة</h3></div>
        <div class="branchpeaks-bars">
          ${topDays.map(d=>`
            <div class="branchpeaks-row day">
              <div class="branchpeaks-label">${d.date}</div>
              <div class="branchpeaks-track"><div class="branchpeaks-fill" style="width:${(d.total/maxDayValue)*100}%"></div></div>
              <div class="branchpeaks-num">${d.total} طلب</div>
            </div>`).join('')}
        </div>
      </div>
    </div>

    <div class="card branchpeaks-table">
      <div class="section-title"><h3 style="margin:0">تقرير الذروة اليومي</h3></div><div class="note" style="margin-bottom:12px">يعرض هذا التقرير كل يوم على حدة: إجمالي الطلبات، ساعة الذروة لذلك اليوم، وعدد الطلبات داخل ساعة الذروة.</div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>إجمالي الطلبات</th>
              <th>ساعة الذروة</th>
              <th>طلبات ساعة الذروة</th>
              <th>ملاحظة</th>
            </tr>
          </thead>
          <tbody>
            ${recentDays.map(d=>`
              <tr>
                <td>${d.date}</td>
                <td>${d.total}</td>
                <td>${to12HourLabel(d.peakHour)}</td>
                <td>${d.peakCount}</td>
                <td>${d.peakCount >= Math.max(5, p.avgDaily*0.25) ? 'ضغط واضح' : 'طبيعي'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  </div>`;
}



function renderAreasProfessional(c){
  const rows = (c.areas || []).slice(0, 50);
  if(!rows.length){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل المناطق</h3></div><div class="empty">No Data مناطق لعرضها.</div></div>`;
  }

  const totalCompleted = rows.reduce((s,a)=>s + Number(a.completed || a.orders || 0), 0);
  const totalSales = rows.reduce((s,a)=>s + Number(a.sales || 0), 0);
  const topArea = rows[0] || {};
  const avgCancel = rows.length ? rows.reduce((s,a)=>s + Number(a.cancelRate || 0), 0) / rows.length : 0;

  return `
    <div class="card">
      <div class="areas-compact-head">
        <div>
          <div class="section-title"><h3 style="margin:0">تحليل المناطق</h3></div>
          <div class="note">تم تنظيم الجدول بشكل أوضح بدون فراغ كبير في المنتصف، مع أهم المؤشرات لكل منطقة.</div>
        </div>
      </div>

      <div class="areas-summary-grid">
        <div class="areas-summary-card">
          <div class="k">عدد المناطق الظاهرة</div>
          <div class="v">${rows.length}</div>
        </div>
        <div class="areas-summary-card">
          <div class="k">إجمالي الطلبات المكتملة</div>
          <div class="v">${totalCompleted}</div>
        </div>
        <div class="areas-summary-card">
          <div class="k">أعلى منطقة</div>
          <div class="v">${topArea.name || '—'}</div>
        </div>
        <div class="areas-summary-card">
          <div class="k">متوسط الإلغاء</div>
          <div class="v">${avgCancel.toFixed(1)}%</div>
        </div>
      </div>

      <div class="tableWrap areas-compact-table">
        <table>
          <thead>
            <tr>
              <th class="area-name">المنطقة</th>
              <th>الطلبات المكتملة</th>
              <th>إجمالي المبيعات</th>
              <th>متوسط الأوردر</th>
              <th>نسبة الإلغاء</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(a=>`
              <tr>
                <td class="area-name">${a.name || 'غير محدد'}</td>
                <td>${Number(a.completed || a.orders || 0)}</td>
                <td>${money(a.sales || 0)}</td>
                <td>${money(a.avgOrder || 0)}</td>
                <td>${Number(a.cancelRate || 0).toFixed(1)}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}




function renderKpiTab(c){
  const rows = getActiveKpiRows();
  const best = getBestKpiPreparer();
  const month = getActiveKpiMonth() || '—';
  const source = getActiveKpiSourceName();
  const top5 = rows.slice(0,5);
  const avgPct = rows.length ? rows.reduce((s,r)=>s+Number(r.pct||0),0)/rows.length : 0;

  return `
    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">ربط KPI</h3></div>
        <div class="note" style="margin-bottom:12px">
          التبويب ده مخصص لقراءة KPI من <strong>تبويب Final</strong> داخل ملف Excel فقط.
          ارفع ملف الـ KPI بصيغة <strong>Excel</strong> وسيتم قراءة آخر شهر تلقائيًا وعرض أفضل محضّر بناءً على <strong>النسبة</strong>.
        </div>
        ${isAdmin() ? `
          <div class="toolbar" style="margin-bottom:10px" data-admin-only="true">
            <input id="kpiFileInput" type="file" accept=".xlsx,.xls" class="input" style="max-width:320px" onchange="document.getElementById('kpiMsg').textContent=this.files[0] ? ((state.lang==='en' ? 'Selected file: ' : 'تم اختيار الملف: ') + this.files[0].name) : ''">
            <button class="btn btn-primary" onclick="importKpiFromFile()">قراءة KPI من Excel</button>
            <button class="btn btn-outline" onclick="clearKpiStore();render();document.getElementById('msg').textContent=(state.lang==='en' ? 'Current KPI file cleared' : 'تم مسح ملف KPI الحالي');">مسح ملف KPI</button>
          </div>
        ` : `
          <div class="note">عرض فقط — نفس التبويبات متاحة هنا، لكن بدون رفع ملفات أو تعديل إعدادات.</div>
        `}
        <div id="kpiMsg" class="small muted" style="margin-top:8px"></div><div class="small muted" style="margin-top:8px">المصدر الحالي: ${source} — الشهر الحالي: ${month}</div>
      </div>

      <div class="card">
        <div class="section-title"><h3 style="margin:0">ملخص KPI</h3></div>
        <div class="peak-kpi-grid">
          <div class="peak-kpi"><div class="label">الشهر</div><div class="value">${month}</div></div>
          <div class="peak-kpi"><div class="label">أفضل محضّر</div><div class="value">${best?.name || '—'}</div><div class="sub">${best ? best.pct.toFixed(1)+'%' : '—'}</div></div>
          <div class="peak-kpi"><div class="label">عدد المحضّرين</div><div class="value">${rows.length}</div></div>
          <div class="peak-kpi"><div class="label">متوسط النسبة</div><div class="value">${avgPct.toFixed(1)}%</div></div>
          <div class="peak-kpi"><div class="label">مصدر القراءة</div><div class="value">${state.kpi?.sourceType === 'file' ? 'Excel' : 'لا يوجد'}</div></div>
        </div>
      </div>
    </div>

    <div class="grid2eq">
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل 5 محضّرين من KPI</h3></div>
        ${top5.length ? renderBars(top5.map(x=>({name:x.name, pct:x.pct})), 'name', 'pct', '%') : '<div class="note">No Data KPI حالياً</div>'}
      </div>
      <div class="card">
        <div class="section-title"><h3 style="margin:0">أفضل محضّر في الداشبورد</h3></div>
        <div class="note">
          • الاسم: <strong>${best?.name || '—'}</strong><br>
          • النسبة: <strong>${best ? best.pct.toFixed(1)+'%' : '—'}</strong><br>
          • الشهر: <strong>${month}</strong><br>
          • المصدر: <strong>${source}</strong>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="section-title"><h3 style="margin:0">جدول KPI الكامل</h3></div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>الترتيب</th>
              <th>الاسم</th>
              <th>النسبة</th>
              <th>الدرجة النهائية</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map((r, i)=>`<tr><td class="num-cell">${i+1}</td><td>${r.name}</td><td class="num-cell">${Number(r.pct||0).toFixed(1)}%</td><td class="num-cell">${Number(r.final||0).toFixed(2)}</td></tr>`).join('') || '<tr><td colspan="4">No Data KPI</td></tr>'}
          </tbody>
        </table>
      </div>
    </div>
  `;
}





function monthlyVsPreviousHtml(c){
  const months = c.monthlyOrders || [];
  if(months.length < 2){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Monthly vs Previous Month':'مقارنة الشهر الحالي بالسابق'}</h3></div><div class="note">${state.lang==='en'?'Not enough data for monthly comparison.':'لا توجد بيانات كافية للمقارنة الشهرية.'}</div></div>`;
  }

  const cur = months[months.length - 1];
  const prev = months[months.length - 2];

  const pctChange = (current, previous) => {
    const c = Number(current || 0), p = Number(previous || 0);
    if(!p && !c) return 0;
    if(!p) return 100;
    return ((c - p) / p) * 100;
  };

  const salesDiff = pctChange(cur.sales, prev.sales);
  const ordersDiff = pctChange(cur.completed, prev.completed);
  const cancelDiff = pctChange(cur.cancelRate, prev.cancelRate);

  const trendBadge = (v, inverse=false) => {
    const good = inverse ? v < 0 : v > 0;
    const bad = inverse ? v > 0 : v < 0;
    if(good) return `<span class="badge ok">${v > 0 ? '▲' : '▼'} ${Math.abs(v).toFixed(1)}%</span>`;
    if(bad) return `<span class="badge bad">${v > 0 ? '▲' : '▼'} ${Math.abs(v).toFixed(1)}%</span>`;
    return `<span class="badge warn">0.0%</span>`;
  };

  return `
    <div class="card">
      <div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Monthly vs Previous Month':'مقارنة الشهر الحالي بالسابق'}</h3></div>
      <div class="note" style="margin-bottom:12px">
        ${state.lang==='en'
          ? `Current: <strong>${cur.month}</strong> vs Previous: <strong>${prev.month}</strong>`
          : `الحالي: <strong>${cur.month}</strong> مقابل السابق: <strong>${prev.month}</strong>`}
      </div>
      <div class="grid3">
        <div class="card" style="padding:14px">
          <div class="small muted">${state.lang==='en'?'Sales':'المبيعات'}</div>
          <div style="font-size:24px;font-weight:800;margin-top:8px">${money(cur.sales)}</div>
          <div class="small muted" style="margin-top:6px">${state.lang==='en'?'Previous':'السابق'}: ${money(prev.sales)}</div>
          <div style="margin-top:8px">${trendBadge(salesDiff)}</div>
        </div>
        <div class="card" style="padding:14px">
          <div class="small muted">${state.lang==='en'?'Completed Orders':'الطلبات المكتملة'}</div>
          <div style="font-size:24px;font-weight:800;margin-top:8px">${cur.completed||0}</div>
          <div class="small muted" style="margin-top:6px">${state.lang==='en'?'Previous':'السابق'}: ${prev.completed||0}</div>
          <div style="margin-top:8px">${trendBadge(ordersDiff)}</div>
        </div>
        <div class="card" style="padding:14px">
          <div class="small muted">${state.lang==='en'?'Cancellation Rate':'نسبة الإلغاء'}</div>
          <div style="font-size:24px;font-weight:800;margin-top:8px">${Number(cur.cancelRate||0).toFixed(1)}%</div>
          <div class="small muted" style="margin-top:6px">${state.lang==='en'?'Previous':'السابق'}: ${Number(prev.cancelRate||0).toFixed(1)}%</div>
          <div style="margin-top:8px">${trendBadge(cancelDiff, true)}</div>
        </div>
      </div>
    </div>
  `;
}


function boardRankingSummaryHtml(c){
  const rows = (c.branches || []).slice(0, 8);
  if(!rows.length){
    return `<div class="card"><div class="section-title"><h3 style="margin:0">Board Ranking Summary</h3></div><div class="note">No Data</div></div>`;
  }
  return `
    <div class="card">
      <div class="section-title"><h3 style="margin:0">${state.lang === 'en' ? 'Board Ranking Summary' : 'ملخص موحد للإدارة'}</h3></div>
      <div class="note" style="margin-bottom:12px">${state.lang === 'en' ? 'A quick branch ranking snapshot for top management.' : 'ملخص سريع يوضح ترتيب الفروع للإدارة العليا بدون الحاجة للتنقل بين التبويبات.'}</div>
      <div class="tableWrap">
        <table style="min-width:760px">
          <thead>
            <tr>
              <th>${state.lang === 'en' ? 'Branch' : 'الفرع'}</th>
              <th>${state.lang === 'en' ? 'Sales' : 'المبيعات'}</th>
              <th>${state.lang === 'en' ? 'Target %' : 'الهدف %'}</th>
              <th>${state.lang === 'en' ? 'Cancel %' : 'الإلغاء %'}</th>
              <th>${state.lang === 'en' ? 'Status' : 'الحالة'}</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(b=>`
              <tr>
                <td>${b.name || '—'}</td>
                <td class="num-cell">${money(b.sales || 0)}</td>
                <td class="num-cell">${Number(b.targetPct || 0).toFixed(1)}%</td>
                <td class="num-cell">${Number(b.cancelRate || 0).toFixed(1)}%</td>
                <td class="num-cell">${targetStatusBadge(b.targetPct,b.target)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}


function renderTabContent(c){
  if(state.activeTab==='overview') return `${boardRankingSummaryHtml(c)}${monthlyVsPreviousHtml(c)}<div class="grid3"><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Branch Performance':'أداء الفروع'}</h3></div>${progressList(c.branches.slice(0,6),'score','','branches')}</div><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Areas by Completed Orders':'المناطق حسب الطلبات المكتملة'}</h3></div>${progressList(c.areas.slice(0,6),'orders',' طلب','areas')}</div><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top Platforms':'أفضل المنصات'}</h3></div>${progressList(c.platforms.slice(0,6),'sales','','platforms')}</div></div><div class="grid3"><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Daily KPI for Management':'KPI يومي للإدارة'}</h3></div><div class="row"><span>Last Day Orders</span><strong>${c.daily.today.orders||0}</strong></div><div class="row"><span>Last Day Sales</span><strong>${money(c.daily.today.sales||0)}</strong></div><div class="row"><span>Last Day Cancellation</span><strong>${(c.daily.today.cancelRate||0).toFixed(1)}%</strong></div><div class="row"><span>Last Day Peak Hour</span><strong>${c.daily.today.peakHour||'—'}</strong></div></div>
      <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Target Summary':'ملخص التارجت'}</h3></div><div class="row"><span>Total Target</span><strong>${money(c.targetSummary.target)}</strong></div><div class="row"><span>Total Achieved</span><strong>${money(c.targetSummary.sales)}</strong></div><div class="row"><span>Achievement Rate</span><strong>${c.targetSummary.pct.toFixed(1)}%</strong></div><div class="row"><span>Remaining</span><strong>${money(c.targetSummary.remaining)}</strong></div></div><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Monthly Comparison':'المقارنة الشهرية'}</h3></div>${monthlyCompareCard(c)}</div><div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Color Alerts':'التنبيهات بالألوان'}</h3></div>${alertsHtml(buildColorAlerts(c))}</div></div>
    <div class="grid3">
      <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top 5 Branches':'أفضل 5 فروع'}</h3></div>${progressList(c.branches.slice(0,5),'score','%')}</div>
      <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Lowest 5 Platforms':'أقل 5 منصات'}</h3></div>${progressList([...c.platforms].sort((a,b)=>(a.completed||0)-(b.completed||0)).slice(0,5),'completed',' طلب')}</div>
      <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Lowest 5 Drivers':'أقل 5 طيارين'}</h3></div>${progressList([...c.drivers].sort((a,b)=>(a.score||0)-(b.score||0)).slice(0,5),'score','%')}</div>
    </div>`;
  if(state.activeTab==='daily') return dailyReportHtml(c);
  if(state.activeTab==='period') return periodAnalysisHtml(c);
  if(state.activeTab==='charts') return `<div class="grid2eq">
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Branch Sales':'مبيعات الفروع'}</h3></div>${renderBars(c.branches.slice(0,10),'name','sales')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Target Achievement by Branch':'نسبة تحقيق التارجت حسب الفرع'}</h3></div>${renderBars(c.branches.filter(x=>x.target).slice(0,10).map(x=>({name:x.name,pct:Math.min(100,x.targetPct||0)})),'name','pct','%')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Completed Orders by Branch':'عدد الطلبات المكتملة حسب اسم الفرع'}</h3></div>${renderBars(c.branches.slice(0,10),'name','completed',' طلب')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top Platforms':'أفضل المنصات'} من علامات التصنيف</h3></div>${renderBars(c.platforms.slice(0,10),'name','sales')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Completed Orders by Platform':'عدد الطلبات المكتملة حسب المنصات'}</h3></div>${renderBars(c.platforms.slice(0,10),'name','completed',' طلب')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top Preparers':'أفضل المحضّرين'}</h3></div>${renderBars(c.preparers.slice(0,10),'name','completed',' تم')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top Drivers':'أفضل الطيارين'}</h3></div>${renderBars(c.drivers.slice(0,10),'name','completed',' تم')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Top Areas':'أعلى المناطق'}</h3></div>${renderBars(c.areas.slice(0,10),'name','completed',' طلب')}</div>
    <div class="card"><div class="section-title"><h3 style="margin:0">${state.lang==='en'?'Monthly Sales':'المبيعات الشهرية'}</h3></div>${renderBars(c.monthlyOrders.slice(-12),'month','sales')}</div>
  </div>`;
  if(state.activeTab==='branchpeaks') return branchPeaksProfessionalHtml(c.data);
  if(state.activeTab==='branches') return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل الفروع بالتفصيل</h3></div><div class="tableWrap"><table><thead><tr><th>الفرع</th><th>الطلبات المكتملة</th><th>المبيعات</th><th>الهدف</th><th>Achievement Rate</th><th>Remaining</th><th>الحالة</th><th>متوسط الأوردر</th><th>المكتمل</th><th>غير مكتمل</th><th>نسبة الإلغاء</th><th>معدل النجاح</th><th>معدل الأداء</th><th>أفضل منطقة</th><th>أكثر عميل</th><th>أعلى ساعة</th><th>أفضل محضّر</th><th>أفضل مندوب</th></tr></thead><tbody>${c.branches.map(b=>`<tr><td>${b.name}</td><td class="num-cell">${b.orders}</td><td class="num-cell">${money(b.sales)}</td><td class="num-cell">${money(b.target||0)}</td><td class="num-cell">${(b.targetPct||0).toFixed(1)}%</td><td class="num-cell">${money(Math.max(0,(b.target||0)-(b.sales||0)))}</td><td class="num-cell">${targetStatusBadge(b.targetPct,b.target)}</td><td class="num-cell">${money(b.avgOrder)}</td><td class="num-cell">${b.completed}</td><td class="num-cell">${b.canceled}</td><td class="num-cell">${b.cancelRate.toFixed(1)}%</td><td class="num-cell">${b.successRate.toFixed(1)}%</td><td class="num-cell"><span class="badge-wrap">${b.score.toFixed(1)}% ${scoreBadge(b.score)}</span></td><td>${b.topArea||''}</td><td>${b.topCustomer||''}</td><td class="num-cell">${b.peak||''}</td><td>${b.bestPreparer||''}</td><td>${b.bestDriver||''}</td></tr>`).join('')}</tbody></table></div></div>`;
  if(state.activeTab==='drivers') return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل الطيارين / المندوبين</h3></div><div class="tableWrap"><table><thead><tr><th>الطيار</th><th>الطلبات</th><th>المكتمل</th><th>غير مكتمل</th><th>المبيعات</th><th>متوسط الأوردر</th><th>معدل النجاح</th><th>معدل الأداء</th></tr></thead><tbody>${c.drivers.map(d=>`<tr><td>${d.name}</td><td class="num-cell">${d.orders}</td><td class="num-cell">${d.completed}</td><td class="num-cell">${d.canceled}</td><td class="num-cell">${money(d.sales)}</td><td class="num-cell">${money(d.avgOrder)}</td><td class="num-cell">${d.successRate.toFixed(1)}%</td><td class="num-cell"><span class="badge-wrap">${d.score.toFixed(1)}% ${scoreBadge(d.score)}</span></td></tr>`).join('')}</tbody></table></div></div>`;
  if(state.activeTab==='platforms') return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل المنصات</h3></div><div class="tableWrap"><table><thead><tr><th>المنصة</th><th>الطلبات</th><th>المكتمل</th><th>غير مكتمل</th><th>المبيعات</th><th>متوسط الأوردر</th><th>معدل النجاح</th><th>نسبة الإلغاء</th></tr></thead><tbody>${c.platforms.map(p=>`<tr><td>${p.name}</td><td class="num-cell">${p.orders}</td><td class="num-cell">${p.completed}</td><td class="num-cell">${p.canceled}</td><td class="num-cell">${money(p.sales)}</td><td class="num-cell">${money(p.avgOrder)}</td><td class="num-cell">${p.successRate.toFixed(1)}%</td><td class="num-cell">${p.cancelRate.toFixed(1)}%</td></tr>`).join('')}</tbody></table></div></div>`;
  if(state.activeTab==='areas') return renderAreasProfessional(c);
  if(state.activeTab==='customers') return `<div class="card"><div class="section-title"><h3 style="margin:0">تحليل العملاء</h3></div><div class="tableWrap"><table><thead><tr><th>العميل</th><th>الهاتف</th><th>عدد الطلبات</th><th>إجمالي الإنفاق</th><th>متوسط الأوردر</th><th>أكثر فرع</th></tr></thead><tbody>${c.customers.slice(0,250).map(x=>`<tr><td>${x.name||''}</td><td class="num-cell">${x.phone||''}</td><td class="num-cell">${x.orders}</td><td class="num-cell">${money(x.spent)}</td><td class="num-cell">${money(x.avgOrder)}</td><td>${x.topBranch||''}</td></tr>`).join('')}</tbody></table></div></div>`;
  if(state.activeTab==='kpi') return renderKpiTab(c);
  if(state.activeTab==='settings') return renderSettingsTab(c);
  return `<div class="card"><div class="section-title"><h3 style="margin:0">نظرة عامة</h3></div><div class="note">اختر تبويبًا من القائمة الجانبية لعرض التفاصيل.</div></div>`;
}


function brLang(){ return state.lang === 'en' ? 'en' : 'ar'; }
function brDir(){ return brLang() === 'en' ? 'ltr' : 'rtl'; }
function brText(ar, en){ return brLang() === 'en' ? en : ar; }

function reportBars(items, labelKey, valueKey, suffix=''){
  if(!items || !items.length) return `<div class="a4-small">${brText('لا توجد بيانات','No Data')}</div>`;
  const max = Math.max(...items.map(x=>Number(x[valueKey]||0)), 1);
  return `<div class="a4-chart">${items.map(it=>{
    const v=Number(it[valueKey]||0), w=(v/max)*100;
    return `<div class="a4-chart-row"><div class="a4-chart-label">${it[labelKey]||'—'}</div><div class="a4-bar-wrap"><div class="a4-bar" style="width:${w}%"></div></div><div class="a4-chart-value">${money(v)}${suffix}</div></div>`;
  }).join('')}</div>`;
}


function boardReportHeader(title, subtitle){
  const isEn = brLang() === 'en';
  return `<div class="a4-header">
    <div class="a4-brand">
      <img class="a4-logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAn8AAAJ6CAYAAABZgyRIAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAP+lSURBVHhe7P13sCVXnt+JfY5Jc92z9cqjqlDwwDTQDaD9AD3T3TPT4zg0uzukyNUuzUZIitAy9If+UWgZ2lAopI2QuLHa2JDEYGiX3OWK5JCc4cz02DbTDg20h2vYAspXvarnr83MY/THOfneRTVQrxtvqquqKz8VGffVvTfzZp48meebv/Mzoqoqz3sgpbz2rZ8IIcS1b/2l4py79q2fiL3vn4/L+0HE5cbh/fvdt9uD3frn7sd/Y8/f3vvXjWWv18/tzs0+P7v9/p1+fm40u7X/7tzY+8du7H5/2xt7bZ/d1r+V+ne9r9P7vHv77uX833yEMeY99363k3ez2f3kXJ+9H99eTv7eL/7d2Gv73Or8rJ+/G83Pev/Yjb33nxvLnX5+bjR7P/839/5xo/vH3tvn+tzo/f9JuBPF3/VNJw0NDQ0NDQ0NDT9TCGvtDZOuuyvnm8ven2z2ovz3/uS3G7d6+99odj+/N/b83entf6uze/+4sTT94+ay9/N/Y+8fu3Gj+89e2+dG799fJo3lr6GhoaGhoaGh4WeaRvw1NDQ0NDQ0NNxB3NBp372yu9n1xrK72XsvZt+9m/1342a3383mZp+/O739b3d27z/Xpzn/tzZ7Pb93+v1jt/a71fd/mmbat6GhoaGhoaGh4WeaRvw1NDQ0NDQ0NNxBNOKvoaGhoaGhoeEOohF/DQ0NDQ0NDQ13EI34a2hoaGhoaGi4g2jEX0NDQ0NDQ0PDHUST6uU67BbKvvcw7922vzdudvvdbG72+bvT2/92Z/f+c32a839rs9fze6ffP3Zrv1t9/6d5/6lebiZ7+33hnPNMHejuB9xwqyDl9Q23zrlr32poaGho+DFo7q93Njf6/O+2/etrsfqz633n+lz/1xsaGhoaGhoaGm4hrm91/XFoxF9DQ0NDQ0NDwx1EM+17G/Oz5HPR0NDQcCvR3F/vbG70+d9t+7vjm2nfhoaGhoaGhoaGH4/G8tfQ0NDQ0NDQcAuxu2Wwsfw1NDQ0NDQ0NDT8mDTir6GhoaGhoaHhDkIYY96/3bChoaGh4WeS3aadGhehhoabR7g83/812Fj+GhoaGhoaGhruIBrx19DQ0NDQ0NBwB9GIv4aGhoaGhoaGO4hG/DU0NDQ0NDQ03EE04q+hoaGhoaGh4Q6iEX8NDQ0NDQ0NDXcQjfhraGhoaGhoaLiDaPL8NTQ0NDQ0NDTcRjR5/hoaGhoaGhoaGn5sGvHX0NDQ0NDQ0H">
      <div>
        <div class="a4-title">${title}</div>
        <div class="a4-sub">${subtitle}</div>
        <div class="a4-badge">${brText('تقرير تنفيذي جاهز للإدارة','Board-ready executive report')}</div>
      </div>
    </div>
    <div class="a4-small">${brText('تاريخ الإنشاء','Generated')}: ${new Date().toLocaleString(isEn ? 'en-GB' : 'ar-EG')}</div>
  </div>`;
}

function buildBoardReportHtml(c){
  const bestBranchesSorted = [...(c.branches||[])].sort((a,b)=>(b.score||0)-(a.score||0));
  const isEn = brLang() === 'en';
  const weakBranch = [...(c.branches||[])].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const bestBranch = bestBranchesSorted[0];
  const secondBranch = bestBranchesSorted[1];
  const bestDriver = c.drivers?.[0];
  const bestPlatform = c.platforms?.[0];
  const bestPreparer = c.bestKpiPreparer || c.preparers?.[0];
  const months = c.monthlyOrders || [];
  const cur = months.length ? months[months.length-1] : null;
  const prev = months.length > 1 ? months[months.length-2] : null;
  const currentMonthName = cur?.month || '—';
  const previousMonthName = prev?.month || '—';
  const branchGap = bestBranch && secondBranch ? Number(bestBranch.sales||0) - Number(secondBranch.sales||0) : null;
  const changePct = (a,b) => {
    const x = Number(a||0), y = Number(b||0);
    if(!x && !y) return 0;
    if(!y) return 100;
    return ((x-y)/y)*100;
  };
  const salesChange = cur && prev ? changePct(cur.sales, prev.sales) : null;
  const ordersChange = cur && prev ? changePct(cur.completed, prev.completed) : null;
  const cancelChange = cur && prev ? changePct(cur.cancelRate, prev.cancelRate) : null;

  const statusLabel = (pct) => {
    const n = Number(pct||0);
    if(n >= 100) return {text: brText('ممتاز','Excellent'), cls:'good'};
    if(n >= 80) return {text: brText('مستقر','Stable'), cls:'warn'};
    return {text: brText('يحتاج متابعة','Needs Attention'), cls:'bad'};
  };
  const statusPill = (pct) => {
    const s = statusLabel(pct);
    return `<span class="status-pill ${s.cls}">${s.text}</span>`;
  };
  const deltaBadge = (value, inverse=false) => {
    if(value == null || Number.isNaN(Number(value))) return '—';
    const v = Number(value);
    if(v === 0) return `<span class="delta-badge neutral">0.0%</span>`;
    const good = inverse ? v < 0 : v > 0;
    const cls = good ? 'up' : 'down';
    const arrow = v > 0 ? '▲' : '▼';
    return `<span class="delta-badge ${cls}">${arrow} ${Math.abs(v).toFixed(1)}%</span>`;
  };

  const risks = [];
  if((c.cancelRate||0) > Number(state.settings.alertThresholds.cancelRate||10)) risks.push(brText(`نسبة الإلغاء الحالية عند <strong>${Number(c.cancelRate||0).toFixed(1)}%</strong> وتتجاوز الحد المستهدف، ما يستدعي تدخلًا تشغيليًا مباشرًا.`,`The current cancellation rate stands at <strong>${Number(c.cancelRate||0).toFixed(1)}%</strong>, above the target threshold, and requires direct operational intervention.`));
  if(weakBranch) risks.push(brText(`يظهر <strong>${weakBranch.name}</strong> كأضعف فرع في الأداء الحالي، ويحتاج إلى خطة تصحيح مركزة خلال الفترة المقبلة.`,`<strong>${weakBranch.name}</strong> is currently the weakest branch and requires a focused recovery plan in the coming period.`));
  if(cur && prev && cur.sales < prev.sales) risks.push(brText(`المبيعات في <strong>${currentMonthName}</strong> جاءت أقل من <strong>${previousMonthName}</strong>، وهو ما يشير إلى تباطؤ يجب تفسيره سريعًا.`,`Sales in <strong>${currentMonthName}</strong> trailed <strong>${previousMonthName}</strong>, indicating a slowdown that should be explained promptly.`));
  if(!risks.length) risks.push(brText('لا توجد مخاطر حرجة ظاهرة حاليًا، مع استمرار الحاجة إلى المتابعة الدورية لمؤشرات التنفيذ والإلغاء.','No critical risks are currently flagged, while routine monitoring of execution and cancellation indicators should continue.'));

  const opps = [];
  if(bestBranch) opps.push(brText(`يُظهر <strong>${bestBranch.name}</strong> أفضل أداء تشغيلي حاليًا، ويمكن استخدامه كنموذج مرجعي لباقي الفروع.`,`<strong>${bestBranch.name}</strong> is currently the strongest branch and can serve as the operating benchmark for the rest of the network.`));
  if(bestPlatform) opps.push(brText(`تُعد <strong>${bestPlatform.name}</strong> المنصة الأقوى أداءً، ما يخلق فرصة واضحة لتعظيم المبيعات من خلالها.`,`<strong>${bestPlatform.name}</strong> is the best performing platform, creating a clear route for incremental sales growth.`));
  if(cur && prev && cur.completed > prev.completed) opps.push(brText(`الطلبات المكتملة في <strong>${currentMonthName}</strong> تحسنت مقارنة بالشهر السابق، وهو تطور إيجابي يمكن البناء عليه.`,`Completed orders in <strong>${currentMonthName}</strong> improved versus the prior month, providing a positive operating base to build upon.`));
  if(!opps.length) opps.push(brText('لا توجد فرصة توسع بارزة في هذه اللحظة، لكن يمكن رفع الكفاءة عبر ضبط الإلغاء وتعزيز أفضل الممارسات التشغيلية.','No standout growth opportunity is highlighted at the moment, though efficiency can still be improved through tighter cancellation control and replication of best practices.'));

  const rankingRows = (c.branches||[]).slice(0,8).map(b=>`
    <tr>
      <td>${b.name||'—'}</td>
      <td>${money(b.sales||0)}</td>
      <td>${Number(b.targetPct||0).toFixed(1)}%</td>
      <td>${Number(b.cancelRate||0).toFixed(1)}%</td>
      <td>${statusPill(b.targetPct)}</td>
    </tr>
  `).join('');

  const executiveNarrative = brText(
    `يعرض هذا التقرير التنفيذي صورة مركزة لأداء شبكة سي ستور في الفترة الحالية أمام مجلس الإدارة. بلغ إجمالي المبيعات <strong>${money(c.sales||0)}</strong> من خلال <strong>${c.total||0}</strong> طلبًا مكتملًا، بمتوسط أوردر <strong>${money(c.avg||0)}</strong>. وقد سجلت الشبكة نسبة إلغاء قدرها <strong>${Number(c.cancelRate||0).toFixed(1)}%</strong>، فيما بلغت نسبة تحقيق الأهداف <strong>${Number(c.targetSummary?.pct||0).toFixed(1)}%</strong>. ويتصدر المشهد التشغيلي فرع <strong>${bestBranch?.name||'—'}</strong>، بينما يحتاج فرع <strong>${weakBranch?.name||'—'}</strong> إلى متابعة أكثر قربًا. ويبرز كذلك <strong>${bestPreparer?.name||'—'}</strong> كأفضل محضّر وفق قراءات KPI الحالية.` ,
    `This executive report provides a concise board-level view of current network performance at C Store. Total sales reached <strong>${money(c.sales||0)}</strong> from <strong>${c.total||0}</strong> completed orders, with an average order value of <strong>${money(c.avg||0)}</strong>. The network recorded a cancellation rate of <strong>${Number(c.cancelRate||0).toFixed(1)}%</strong>, while overall target achievement reached <strong>${Number(c.targetSummary?.pct||0).toFixed(1)}%</strong>. <strong>${bestBranch?.name||'—'}</strong> is the leading branch at present, whereas <strong>${weakBranch?.name||'—'}</strong> requires closer management attention. KPI readings also identify <strong>${bestPreparer?.name||'—'}</strong> as the top preparer currently.`
  );

  const chairmanNote = brText(
    `الأولوية الإدارية خلال الفترة المقبلة تتمثل في تثبيت مستوى الأداء في <strong>${bestBranch?.name||'الفرع الأفضل'}</strong>، وتسريع خطة التحسين في <strong>${weakBranch?.name||'الفرع الأضعف'}</strong>، مع تكثيف المتابعة اليومية لمعدل الإلغاء. كما يُوصى بالاستفادة من زخم <strong>${bestPlatform?.name||'المنصة الأقوى'}</strong> لتعزيز النمو، وتعميم أفضل الممارسات التنفيذية على مستوى الفروع.` ,
    `Management priority for the next period should be to sustain performance at <strong>${bestBranch?.name||'the leading branch'}</strong>, accelerate the recovery plan for <strong>${weakBranch?.name||'the weakest branch'}</strong>, and intensify daily oversight of cancellation. The business should also capitalize on the momentum of <strong>${bestPlatform?.name||'the strongest platform'}</strong> and replicate best operating practices across the branch network.`
  );

  return `
    <div class="a4-page board-report ${brLang()}-report" dir="${brDir()}">
      ${boardReportHeader(brText('التقرير التنفيذي لمجلس الإدارة','Board Executive Report'), brText('ملخص استراتيجي وتشغيلي رفيع المستوى','Strategic and Operational Snapshot'))}
      <div class="a4-grid">
        <div class="a4-card"><div class="label">${brText('إجمالي المبيعات','Total Sales')}</div><div class="value">${money(c.sales||0)}</div></div>
        <div class="a4-card"><div class="label">${brText('الطلبات المكتملة','Completed Orders')}</div><div class="value">${c.total||0}</div></div>
        <div class="a4-card"><div class="label">${brText('نسبة الإلغاء','Cancellation Rate')}</div><div class="value">${Number(c.cancelRate||0).toFixed(1)}%</div></div>
        <div class="a4-card"><div class="label">${brText('تحقيق الأهداف','Target Achievement')}</div><div class="value">${Number(c.targetSummary?.pct||0).toFixed(1)}%</div></div>
        <div class="a4-card"><div class="label">${brText('أفضل فرع','Leading Branch')}</div><div class="value">${bestBranch?.name||'—'}</div></div>
        <div class="a4-card"><div class="label">${brText('أضعف فرع','Priority Branch')}</div><div class="value">${weakBranch?.name||'—'}</div></div>
        <div class="a4-card"><div class="label">${brText('أفضل طيار','Top Driver')}</div><div class="value">${bestDriver?.name||'—'}</div></div>
        <div class="a4-card"><div class="label">${brText('أفضل محضّر','Top Preparer')}</div><div class="value">${bestPreparer?.name||'—'}</div></div>
      </div>
      <div class="board-grid-2 a4-section">
        <div>
          <div class="a4-section-title">${brText('الرسالة التنفيذية','Executive Narrative')}</div>
          <div class="a4-lead">${executiveNarrative}</div>
        </div>
        <div>
          <div class="a4-section-title">${brText('مقارنة الشهر الحالي بالسابق','Monthly vs Previous Month')}</div>
          <table class="a4-table">
            <thead><tr><th>${brText('المؤشر','Metric')}</th><th>${brText('الحالي','Current')}</th><th>${brText('السابق','Previous')}</th><th>${brText('التغير','Change')}</th></tr></thead>
            <tbody>
              <tr><td>${brText('المبيعات','Sales')}</td><td>${cur ? money(cur.sales||0) : '—'}</td><td>${prev ? money(prev.sales||0) : '—'}</td><td>${deltaBadge(salesChange)}</td></tr>
              <tr><td>${brText('الطلبات المكتملة','Completed Orders')}</td><td>${cur?.completed ?? '—'}</td><td>${prev?.completed ?? '—'}</td><td>${deltaBadge(ordersChange)}</td></tr>
              <tr><td>${brText('نسبة الإلغاء','Cancellation Rate')}</td><td>${cur ? Number(cur.cancelRate||0).toFixed(1)+'%' : '—'}</td><td>${prev ? Number(prev.cancelRate||0).toFixed(1)+'%' : '—'}</td><td>${deltaBadge(cancelChange, true)}</td></tr>
            </tbody>
          </table>
          <div class="a4-small" style="margin-top:8px">${brText(`الفجوة البيعية بين أفضل فرعين: <strong>${branchGap==null?'—':money(branchGap)}</strong>.`,`Sales gap between the top two branches: <strong>${branchGap==null?'—':money(branchGap)}</strong>.`)}</div>
        </div>
      </div>
      <div class="a4-footer">${brText('تقرير سي ستور التنفيذي - صفحة 1','C Store Executive Report - Page 1')}</div>
    </div>

    <div class="a4-page board-report ${brLang()}-report" dir="${brDir()}">
      ${boardReportHeader(brText('الترتيب الموحد للفروع','Branch Ranking Summary'), brText('نظرة موحدة لاتخاذ القرار الإداري','Unified decision-making view'))}
      <div class="a4-section">
        <div class="a4-section-title">${brText('الترتيب الموحد للفروع','Unified Branch Ranking')}</div>
        <table class="a4-table">
          <thead><tr><th>${brText('الفرع','Branch')}</th><th>${brText('المبيعات','Sales')}</th><th>${brText('تحقيق الهدف %','Target %')}</th><th>${brText('الإلغاء %','Cancel %')}</th><th>${brText('الحالة','Status')}</th></tr></thead>
          <tbody>${rankingRows}</tbody>
        </table>
      </div>
      <div class="board-grid-2 a4-section">
        <div>
          <div class="a4-section-title">${brText('المخاطر الرئيسية','Key Risks')}</div>
          <ul class="board-bullets">${risks.map(x=>`<li>${x}</li>`).join('')}</ul>
        </div>
        <div>
          <div class="a4-section-title">${brText('أبرز فرص التحسين والنمو','Opportunities & Upside')}</div>
          <ul class="board-bullets">${opps.map(x=>`<li>${x}</li>`).join('')}</ul>
        </div>
      </div>
      <div class="a4-section">
        <div class="a4-section-title">${brText('لقطة القيادات التشغيلية','Leadership Snapshot')}</div>
        <table class="a4-table">
          <thead><tr><th>${brText('البند','Item')}</th><th>${brText('القيمة','Value')}</th></tr></thead>
          <tbody>
            <tr><td>${brText('أفضل فرع','Leading Branch')}</td><td>${bestBranch?.name||'—'}</td></tr>
            <tr><td>${brText('أفضل منصة','Leading Platform')}</td><td>${bestPlatform?.name||'—'}</td></tr>
            <tr><td>${brText('أفضل طيار','Top Driver')}</td><td>${bestDriver?.name||'—'}</td></tr>
            <tr><td>${brText('أفضل محضّر','Top Preparer')}</td><td>${bestPreparer?.name||'—'}</td></tr>
            <tr><td>${brText('أفضل عميل','Top Customer')}</td><td>${c.topCustomer?.name||'—'}</td></tr>
          </tbody>
        </table>
      </div>
      <div class="a4-footer">${brText('تقرير سي ستور التنفيذي - صفحة 2','C Store Executive Report - Page 2')}</div>
    </div>

    <div class="a4-page board-report ${brLang()}-report" dir="${brDir()}">
      ${boardReportHeader(brText('المؤشرات التشغيلية والتوصيات التنفيذية','Operational Performance & Management Actions'), brText('رسائل مجلس الإدارة الختامية','Board concluding page'))}
      <div class="board-grid-2">
        <div class="a4-section">
          <div class="a4-section-title">${brText('مبيعات الفروع','Branch Sales')}</div>
          ${reportBars((c.branches||[]).slice(0,6),'name','sales')}
        </div>
        <div class="a4-section">
          <div class="a4-section-title">${brText('الطلبات المكتملة حسب المنصة','Completed Orders by Platform')}</div>
          ${reportBars((c.platforms||[]).slice(0,6),'name','completed','')}
        </div>
      </div>
      <div class="board-grid-2">
        <div class="a4-section">
          <div class="a4-section-title">${brText('أفضل الطيارين','Top Drivers')}</div>
          ${reportBars((c.drivers||[]).slice(0,6),'name','completed','')}
        </div>
        <div class="a4-section">
          <div class="a4-section-title">${brText('المبيعات الشهرية','Monthly Sales')}</div>
          ${reportBars((c.monthlyOrders||[]).slice(-6),'month','sales')}
        </div>
      </div>
      <div class="a4-section">
        <div class="a4-section-title">${brText('توصية الإدارة العليا','Chairman Recommendation')}</div>
        <div class="a4-lead">${chairmanNote}</div>
      </div>
      <div class="a4-footer">${brText('تقرير سي ستور التنفيذي - صفحة 3','C Store Executive Report - Page 3')}</div>
    </div>
  `;
}

function updateBoardReport(c){
  const holder = document.querySelector('.print-report');
  if(holder) holder.innerHTML = buildBoardReportHtml(c);
}


function openBoardReportPrint(){
  const c = calc();
  updateBoardReport(c);
  const reportHtml = document.querySelector('.print-report')?.innerHTML || '';
  if(!reportHtml) return;

  const styles = `<style>
    body{font-family:Tahoma,Arial,sans-serif;background:#fff;color:#111;margin:0}
    .a4-page{width:210mm;min-height:297mm;margin:0 auto;background:#fff;color:#1c1c1c;padding:11mm 10mm;page-break-after:always}
    .a4-page:last-child{page-break-after:auto}
    .a4-header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #ead7e0;padding-bottom:12px;margin-bottom:16px;gap:16px}
    .a4-brand{display:flex;align-items:center;gap:14px}
    .a4-logo{width:74px;height:74px;object-fit:cover;border-radius:18px;box-shadow:0 8px 20px rgba(213,51,114,.10)}
    .a4-title{font-size:28px;font-weight:900;line-height:1.2;letter-spacing:.1px}
    .a4-sub{font-size:13px;color:#6e5b65;margin-top:3px}
    .a4-badge{display:inline-block;padding:6px 10px;border:1px solid #efcbd9;border-radius:999px;font-size:11px;color:#933665;background:#fff4f8;margin-top:7px;font-weight:700}
    .a4-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:12px}
    .a4-card{border:1px solid #ebd8e0;border-radius:16px;padding:10px 12px;background:linear-gradient(180deg,#fff 0%,#fbf4f7 100%);box-shadow:0 4px 12px rgba(213,51,114,.05);min-height:78px}
    .a4-card .label{font-size:11px;color:#6f5d66;font-weight:700}
    .a4-card .value{font-size:21px;font-weight:900;margin-top:5px;line-height:1.3}
    .a4-section{margin-top:14px}
    .a4-section-title{font-size:17px;font-weight:900;margin:0 0 9px 0;padding-bottom:6px;border-bottom:2px solid #f2d9e4;color:#8b1d52}
    .a4-lead{border:1px solid #ecdbe2;border-radius:16px;padding:13px 15px;background:linear-gradient(180deg,#fff 0%,#fcf7fa 100%);line-height:2;font-size:12.5px}
    .a4-table{width:100%;border-collapse:separate;border-spacing:0;font-size:11px;overflow:hidden;border-radius:14px}
    .a4-table th,.a4-table td{border-bottom:1px solid #e9dfe4;padding:8px 7px;vertical-align:top;word-break:break-word;overflow-wrap:anywhere}
    .a4-table thead th{background:#fff3f8;font-weight:900;color:#7f2450;border-top:1px solid #ead6df}
    .a4-table tr td:first-child,.a4-table tr th:first-child{border-inline-start:1px solid #ead6df}
    .a4-table tr td:last-child,.a4-table tr th:last-child{border-inline-end:1px solid #ead6df}
    .a4-table tbody tr:last-child td{border-bottom:1px solid #ead6df}
    .a4-small{font-size:11px;color:#65545d}
    .a4-two-col,.board-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .a4-footer{margin-top:14px;padding-top:8px;border-top:1px solid #ead7e0;font-size:10px;color:#7a6870;text-align:center}
    .board-bullets{margin:0;line-height:1.9;font-size:12px}.board-bullets li{margin:2px 0}.board-bullets li::marker{color:#d53372}
    .a4-chart-row{display:grid;grid-template-columns:minmax(150px,1.35fr) minmax(170px,1fr) minmax(95px,.6fr);gap:10px;align-items:center;margin:9px 0}.a4-chart-label{line-height:1.45;word-break:break-word;overflow-wrap:anywhere;font-size:11px;font-weight:700;color:#42343d}.a4-chart-value{text-align:end;white-space:nowrap;font-variant-numeric:tabular-nums;font-size:11px;font-weight:800;color:#5d4250}
    .a4-bar-wrap{height:10px;background:#f0e6eb;border-radius:999px;overflow:hidden}
    .a4-bar{height:100%;background:linear-gradient(90deg,#d53372,#ef7aa9)}
    .board-report{position:relative}.board-report::before{content:'';position:absolute;top:0;inset-inline:0;height:5px;background:linear-gradient(90deg,#d53372,#ef7aa9)}
    .board-report.ar-report{direction:rtl;text-align:right;font-family:Tahoma,'Segoe UI',Arial,sans-serif}
    .board-report.en-report{direction:ltr;text-align:left;font-family:'Segoe UI',Arial,sans-serif}
    .board-report.ar-report .a4-header{flex-direction:row-reverse}
    .board-report.en-report .a4-header{flex-direction:row}
    .board-report.en-report .a4-table th,.board-report.en-report .a4-table td{text-align:left}
    .board-report.ar-report .a4-table th,.board-report.ar-report .a4-table td{text-align:right}
    .board-report.en-report .board-bullets{padding-inline-start:18px;padding-inline-end:0}
    .board-report.ar-report .board-bullets{padding-inline-start:0;padding-inline-end:18px}
    .board-report.en-report .a4-small{text-align:right}
    .board-report.ar-report .a4-small{text-align:left}
    .status-pill{display:inline-block;padding:4px 10px;border-radius:999px;font-size:10px;font-weight:800}.status-pill.good{background:#eaf8f1;color:#157347}.status-pill.warn{background:#fff4dd;color:#9a6700}.status-pill.bad{background:#fdecef;color:#b42318}
    .delta-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 8px;border-radius:999px;font-size:10px;font-weight:800}.delta-badge.up{background:#eaf8f1;color:#157347}.delta-badge.down{background:#fdecef;color:#b42318}.delta-badge.neutral{background:#f4f4f5;color:#52525b}
    @page{size:A4;margin:12mm}
    @media print{body{margin:0}}
  </style>`;

  const w = window.open('', '_blank');
  if(!w) return;
  w.document.open();
  w.document.write(`<!DOCTYPE html><html lang="${brLang()}" dir="${brDir()}"><head><meta charset="UTF-8"><title>${brText('تقرير مجلس الإدارة - سي ستور','C Store Board Report')}</title>${styles}</head><body>${reportHtml}</body></html>`);
  w.document.close();
  setTimeout(()=>{
    w.focus();
    w.print();
  }, 700);
}


function downloadBoardReportPDF(){
  openBoardReportPrint();
}


function downloadReport(){
  updateBoardReport(calc());
  const reportHtml = document.querySelector('.print-report').innerHTML;
  const styles = `<style>
  body{font-family:Tahoma,Arial,sans-serif;background:#fff;color:#111;margin:0}
  .a4-page{width:210mm;min-height:297mm;margin:0 auto;background:#fff;color:#111;padding:10mm}
  .a4-header{display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #ddd;padding-bottom:8px;margin-bottom:10px}
  .a4-logo{width:70px;height:70px;object-fit:cover;border-radius:12px}
  .a4-title{font-size:22px;font-weight:700}.a4-sub{font-size:12px;color:#555}
  .a4-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}
  .a4-card{border:1px solid #ddd;border-radius:10px;padding:8px}
  .a4-card .label{font-size:11px;color:#666}.a4-card .value{font-size:18px;font-weight:700;margin-top:4px}
  .a4-section{margin-top:10px}.a4-table{width:100%;border-collapse:collapse;font-size:11px}
  .a4-table th,.a4-table td{border:1px solid #ddd;padding:5px;text-align:right}
  .a4-table th{background:#f5f5f5}
  
.alert-box{padding:14px;border-radius:16px;border:1px solid var(--border);margin-bottom:10px}
.alert-red{background:rgba(178,42,84,.18);border-color:#b22a54}
.alert-yellow{background:rgba(161,95,0,.18);border-color:#a15f00}
.alert-green{background:rgba(14,122,104,.14);border-color:#0e7a68}
.alert-title{font-weight:700;margin-bottom:6px}

</style>`;
  const doc = `<!DOCTYPE html><html lang="${brLang()}" dir="${brDir()}"><head><meta charset="UTF-8"><title>${brText('تقرير سي ستور','C Store Report')}</title>${styles}</head><body>${reportHtml}</body></html>`;
  const blob = new Blob([doc], {type:'text/html;charset=utf-8'});
  const a = document.createElement('a');
  const d = new Date();
  const stamp = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  a.href = URL.createObjectURL(blob);
  a.download = `CStore_Report_${stamp}.html`;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 1500);
}


function buildColorAlerts(c){
  const alerts = [];
  const threshold = Number(state.settings.alertThresholds.cancelRate || 10);

  if(c.cancelRate > threshold){
    alerts.push({level:'red', title:'تنبيه إلغاء مرتفع', text:`Current cancellation rate ${c.cancelRate.toFixed(1)}% وتجاوزت الحد ${threshold}%`});
  } else if(c.cancelRate > Math.max(5, threshold * 0.7)){
    alerts.push({level:'yellow', title:'مراقبة الإلغاء', text:`Current cancellation rate ${c.cancelRate.toFixed(1)}% وقريبة من الحد ${threshold}%`});
  } else {
    alerts.push({level:'green', title:'Cancellation Under Control', text:`Current cancellation rate ${c.cancelRate.toFixed(1)}%`});
  }

  const weakBranches = c.branches.filter(b => b.target && b.targetPct < 70).slice(0,3);
  weakBranches.forEach(b=>{
    alerts.push({level:'red', title:'فرع أقل من الهدف', text:`${b.name} حقق ${b.targetPct.toFixed(1)}% فقط من الهدف`});
  });

  const midBranches = c.branches.filter(b => b.target && b.targetPct >= 70 && b.targetPct < 100).slice(0,2);
  midBranches.forEach(b=>{
    alerts.push({level:'yellow', title:'فرع يحتاج دفعة', text:`${b.name} عند ${b.targetPct.toFixed(1)}% من الهدف`});
  });

  const topBranch = c.branches[0];
  if(topBranch){
    alerts.push({level:'green', title:'أفضل فرع حاليًا', text:`${topBranch.name} بمعدل أداء ${topBranch.score.toFixed(1)}%`});
  }

  if(c.monthlyOrders.length >= 2){
    const cur = c.monthlyOrders[c.monthlyOrders.length-1];
    const prev = c.monthlyOrders[c.monthlyOrders.length-2];
    if(cur.sales < prev.sales){
      alerts.push({level:'yellow', title:'انخفاض مبيعات شهري', text:`مبيعات ${cur.month} أقل من ${prev.month}`});
    } else if(cur.sales > prev.sales){
      alerts.push({level:'green', title:'تحسن مبيعات شهري', text:`مبيعات ${cur.month} أعلى من ${prev.month}`});
    }
    if(cur.cancelRate > prev.cancelRate){
      alerts.push({level:'red', title:'زيادة إلغاء شهرية', text:`إلغاء ${cur.month} أعلى من ${prev.month}`});
    }
  }

  return alerts.slice(0,6);
}
function alertsHtml(items){
  if(!items.length) return '<div class="note">لا توجد تنبيهات حالية.</div>';
  return items.map(a=>`<div class="alert-box alert-${a.level}"><div class="alert-title">${a.title}</div><div>${a.text}</div></div>`).join('');
}



function exportReportPDF(){
  updatePrintReport(calc());
  const el = document.querySelector('.print-report .a4-page');
  const d = new Date();
  const stamp = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  if(window.html2pdf){
    const opt = {
      margin: 0,
      filename: `CStore_Report_${stamp}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    html2pdf().set(opt).from(el).save();
  } else {
    const oldTitle = document.title;
    document.title = 'CStore_Report';
    window.print();
    setTimeout(()=>{ document.title = oldTitle; }, 500);
  }
}

function exportReportExcel(){
  const c = calc();
  const summary = [
    ['البند','القيمة'],
    ['إجمالي الطلبات', c.total],
    ['إجمالي المبيعات', Number(c.sales||0)],
    ['متوسط الأوردر', Number(c.avg||0)],
    ['الملغي / المرتجع', c.badCount],
    ['نسبة الإلغاء %', Number((c.cancelRate||0).toFixed(2))],
    ['أعلى ساعة ضغط', c.peakHour || '—'],
    ['معدل الأداء العام', c.total>0 ? Number((((100-c.cancelRate)*0.6 + Math.min(100,c.total/100)*0.4).toFixed(2))) : '—']
  ];

  const leaders = [
    ['البند','الاسم','القيمة'],
    ['أفضل محضّر', c.preparers[0]?.name||'—', c.preparers[0]?.completed||0],
    ['أفضل طيار', c.drivers[0]?.name||'—', c.drivers[0]?.completed||0],
    ['أفضل عميل', c.topCustomer?.name||'—', c.topCustomer?.orders||0],
    ['أفضل منصة', c.platforms[0]?.name||'—', c.platforms[0]?.orders||0]
  ];

  const branches = [['الفرع','Band','الطلبات','المبيعات','الهدف','Achievement Rate %','متوسط الأوردر','المكتمل','غير مكتمل','نسبة الإلغاء %','معدل النجاح %','معدل الأداء %','أفضل منطقة','أكثر عميل','أعلى ساعة','أفضل محضّر','أفضل مندوب']]
    .concat(c.branches.map(b=>[
      b.name||'', b.band||'', b.orders||0, Number(b.sales||0), Number(b.target||0), Number((b.targetPct||0).toFixed(2)),
      Number(b.avgOrder||0), b.completed||0, b.canceled||0, Number((b.cancelRate||0).toFixed(2)),
      Number((b.successRate||0).toFixed(2)), Number((b.score||0).toFixed(2)), b.topArea||'', b.topCustomer||'', b.peak||'', b.bestPreparer||'', b.bestDriver||''
    ]));

  const preparers = [['المحضّر','Band','الطلبات','المكتمل','غير مكتمل','المبيعات','متوسط الأوردر','معدل النجاح %','معدل الأداء %']]
    .concat(c.preparers.map(p=>[
      p.name||'', p.band||'', p.orders||0, p.completed||0, p.canceled||0, Number(p.sales||0), Number(p.avgOrder||0),
      Number((p.successRate||0).toFixed(2)), Number((p.score||0).toFixed(2))
    ]));

  const drivers = [['الطيار','Band','الطلبات','المكتمل','غير مكتمل','المبيعات','متوسط الأوردر','معدل النجاح %','معدل الأداء %']]
    .concat(c.drivers.map(d=>[
      d.name||'', d.band||'', d.orders||0, d.completed||0, d.canceled||0, Number(d.sales||0), Number(d.avgOrder||0),
      Number((d.successRate||0).toFixed(2)), Number((d.score||0).toFixed(2))
    ]));

  const platforms = [['المنصة','Band','الطلبات','المكتمل','غير مكتمل','المبيعات','متوسط الأوردر','معدل النجاح %','نسبة الإلغاء %']]
    .concat(c.platforms.map(p=>[
      p.name||'', p.band||'', p.orders||0, p.completed||0, p.canceled||0, Number(p.sales||0), Number(p.avgOrder||0),
      Number((p.successRate||0).toFixed(2)), Number((p.cancelRate||0).toFixed(2))
    ]));

  const customers = [['العميل','Band','الهاتف','عدد الطلبات','إجمالي الإنفاق','متوسط الأوردر','أكثر فرع']]
    .concat(c.customers.map(x=>[
      x.name||'', x.band||'', x.phone||'', x.orders||0, Number(x.spent||0), Number(x.avgOrder||0), x.topBranch||''
    ]));

  const monthly = [['الشهر','الطلبات','المبيعات','المكتمل','الملغي / المرتجع','نسبة الإلغاء %','متوسط الأوردر']]
    .concat((c.monthlyOrders||[]).map(m=>[
      m.month||'', m.orders||0, Number(m.sales||0), m.completed||0, m.canceled||0, Number((m.cancelRate||0).toFixed(2)), Number(m.avgOrder||0)
    ]));

  const orders = [['رقم الأوردر','التاريخ','الوقت','الفرع','المنصة','العميل','الهاتف','المنطقة','القيمة','الدفع','الحالة','المحضّر','الطيار']]
    .concat((c.data||[]).map(o=>[
      o.orderNumber||'', o.orderDate||'', o.orderTime||'', o.branch||'', o.platform||'', o.customerName||'', o.customerPhone||'', o.area||'',
      Number(o.amount||0), o.paymentMethod||'', o.status||'', o.preparer||'', o.driver||''
    ]));

  const wb = XLSX.utils.book_new();
  const addSheet = (name, rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  addSheet('Summary', summary);
  addSheet('Leaders', leaders);
  addSheet('Branches', branches);
  addSheet('Preparers', preparers);
  addSheet('Drivers', drivers);
  addSheet('Platforms', platforms);
  addSheet('Customers', customers);
  addSheet('Monthly', monthly);
  addSheet('Orders', orders);

  const d = new Date();
  const stamp = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `CStore_Report_${stamp}.xlsx`);
}


function viewerHeaderHtml(c){
  const bestBranch = c.branches[0];
  const bestPlatform = c.platforms[0];
  const bestDriver = c.drivers[0];
  const totalBranches = c.branches.length;
  const periodText = (() => {
    const dates = (state.orders||[]).map(o=>o.orderDate).filter(Boolean).sort();
    if(!dates.length) return 'No Data مرفوعة حتى الآن';
    return dates[0] === dates[dates.length-1] ? dates[0] : `${dates[0]} → ${dates[dates.length-1]}`;
  })();
  return `<div class="viewer-hero">
    <div class="viewer-hero-head">
      <div>
        <div class="viewer-hero-title">لوحة متابعة الإدارة</div>
        <div class="viewer-hero-sub">عرض مختصر وواضح لأهم المؤشرات خلال الفترة الحالية: ${periodText}</div>
      </div>
      <div class="viewer-pill">Viewer Mode • عرض فقط</div>
    </div>
    <div class="viewer-mini-grid">
      <div class="viewer-mini">
        <div class="k">إجمالي الطلبات المكتملة</div>
        <div class="v">${c.total}</div>
        <div class="s">المعتمد: مفوتر بالكامل</div>
      </div>
      <div class="viewer-mini">
        <div class="k">إجمالي المبيعات</div>
        <div class="v">${moneyShort(c.sales)}</div>
        <div class="s">متوسط الأوردر ${moneyShort(c.avgOrder)}</div>
      </div>
      <div class="viewer-mini">
        <div class="k">أفضل فرع حاليًا</div>
        <div class="v">${bestBranch ? bestBranch.name : '—'}</div>
        <div class="s">${bestBranch ? `أداء ${(bestBranch.score||0).toFixed(1)}%` : 'No Data'}</div>
      </div>
      <div class="viewer-mini">
        <div class="k">أفضل منصة / طيار</div>
        <div class="v">${bestPlatform ? bestPlatform.name : '—'}</div>
        <div class="s">${bestDriver ? `أفضل طيار: ${bestDriver.name}` : 'No Data'}</div>
      </div>
    </div>
  </div>`;
}


function executiveSummaryHtml(c){
  const bestBranch = c.branches[0];
  const weakBranch = [...c.branches].sort((a,b)=>(a.score||0)-(b.score||0))[0];
  const bestPlatform = c.platforms[0];
  const bestDriver = c.drivers[0];
  const bestPreparer = c.bestKpiPreparer || c.preparers[0];
  if(state.lang === 'en'){
    return `<div class="exec-strip">
      <div class="exec-main">
        <div class="exec-title">Executive Summary</div>
        <div class="exec-line">
          • Completed Orders: <strong>${c.total}</strong><br>
          • Total Sales: <strong>${money(c.sales)}</strong><br>
          • Avg Order: <strong>${money(c.avg)}</strong><br>
          • Cancellation Rate: <strong>${c.cancelRate.toFixed(1)}%</strong><br>
          • Top Branch: <strong>${bestBranch?.name||'—'}</strong><br>
          • Weakest Branch: <strong>${weakBranch?.name||'—'}</strong><br>
          • Top Platform: <strong>${bestPlatform?.name||'—'}</strong><br>
          • Best Driver: <strong>${bestDriver?.name||'—'}</strong><br>
          • Best Preparer: <strong>${bestPreparer?.name||'—'}</strong> ${c.kpiPreparerMonth ? `<span class="small muted">(${translateTextValue(c.kpiPreparerMonth)})</span>` : ''}
        </div>
      </div>
      <div class="exec-mini">
        <div class="small muted">Target</div>
        <div class="v">${c.targetSummary.pct.toFixed(1)}%</div>
        <div class="small muted">Achievement Rate</div>
      </div>
      <div class="exec-mini">
        <div class="small muted">Sales</div>
        <div class="v">${money(c.sales)}</div>
        <div class="small muted">Current Total</div>
      </div>
      <div class="exec-mini">
        <div class="small muted">Cancellation</div>
        <div class="v">${c.cancelRate.toFixed(1)}%</div>
        <div class="small muted">Quick Indicator</div>
      </div>
    </div>`;
  }
  return `<div class="exec-strip">
    <div class="exec-main">
      <div class="exec-title">ملخص جاهز للإدارة</div>
      <div class="exec-line">
        • الطلبات المكتملة: <strong>${c.total}</strong><br>
        • إجمالي المبيعات: <strong>${money(c.sales)}</strong><br>
        • متوسط الأوردر: <strong>${money(c.avg)}</strong><br>
        • نسبة الإلغاء: <strong>${c.cancelRate.toFixed(1)}%</strong><br>
        • أفضل فرع: <strong>${bestBranch?.name||'—'}</strong><br>
        • الفرع الأضعف: <strong>${weakBranch?.name||'—'}</strong><br>
        • أفضل منصة: <strong>${bestPlatform?.name||'—'}</strong><br>
        • أفضل طيار: <strong>${bestDriver?.name||'—'}</strong><br>
        • أفضل محضّر: <strong>${bestPreparer?.name||'—'}</strong> ${c.kpiPreparerMonth ? `<span class="small muted">(${c.kpiPreparerMonth})</span>` : ''}
      </div>
    </div>
    <div class="exec-mini">
      <div class="small muted">التارجت</div>
      <div class="v">${c.targetSummary.pct.toFixed(1)}%</div>
      <div class="small muted">نسبة التحقيق</div>
    </div>
    <div class="exec-mini">
      <div class="small muted">المبيعات</div>
      <div class="v">${money(c.sales)}</div>
      <div class="small muted">الإجمالي الحالي</div>
    </div>
    <div class="exec-mini">
      <div class="small muted">الإلغاء</div>
      <div class="v">${c.cancelRate.toFixed(1)}%</div>
      <div class="small muted">مؤشر سريع</div>
    </div>
  </div>`;
}

function render(){
  if(state.activeTab==='manager' || state.activeTab==='insights' || state.activeTab==='rankings' || state.activeTab==='payments' || state.activeTab==='preparers') state.activeTab='overview';
  const c=calc(), years=[...new Set(state.orders.map(o=>ymd(o.orderDate).y))].sort((a,b)=>a-b), months=[...new Set(state.orders.map(o=>ymd(o.orderDate).m))].sort((a,b)=>a-b).map(m=>({value:m,label:monthNameLocal(m)})), days=[...new Set(state.orders.map(o=>ymd(o.orderDate).day))].filter(x=>!Number.isNaN(x)).sort((a,b)=>a-b);
  document.getElementById("filters").innerHTML=renderMultiSelect(tr("السنة"),years,"years")+renderMultiSelect(tr("الشهر"),months,"months")+renderMultiSelect(tr("اليوم"),days,"days")+renderMultiSelect(tr("الفرع"),[...new Set(state.orders.map(o=>o.branch).filter(Boolean))].sort(),"branches")+renderMultiSelect(tr("المنطقة"),[...new Set(state.orders.map(o=>o.area).filter(Boolean))].sort(),"areas")+renderMultiSelect(tr("الحالة"),[...new Set(state.orders.map(o=>o.status).filter(Boolean))].sort().map(x=>({value:x,label:statusLocal(x)})),"statuses")+renderMultiSelect(tr("المحضّر"),[...new Set(state.orders.map(o=>o.preparer).filter(Boolean))].sort(),"preparers")+renderMultiSelect(tr("الطيار"),[...new Set(state.orders.map(o=>o.driver).filter(Boolean))].sort(),"drivers")+renderMultiSelect(tr("المنصة"),[...new Set(state.orders.map(o=>o.platform).filter(Boolean))].sort(),"platforms")+renderMultiSelect(tr("الدفع"),[...new Set(state.orders.map(o=>o.paymentMethod).filter(Boolean))].sort(),"payments");
  applyRoleUI();
  const overallPerf = c.total>0 ? ((100-c.cancelRate)*0.6 + Math.min(100,c.total/100)*0.4).toFixed(1) : '—';
  document.getElementById("viewerHeader").innerHTML = isViewer() ? viewerHeaderHtml(c) : "";
  document.getElementById("executiveSummary").innerHTML=executiveSummaryHtml(c);
  document.getElementById("stats").innerHTML=`<div class="stat"><div class="small muted">إجمالي الطلبات</div><div class="v">${c.total}</div></div><div class="stat"><div class="small muted">إجمالي المبيعات</div><div class="v">${moneyShort(c.sales)}</div></div><div class="stat"><div class="small muted">متوسط الأوردر</div><div class="v">${moneyShort(c.avg)}</div></div><div class="stat"><div class="small muted">الملغي / المرتجع</div><div class="v">${c.badCount}</div></div><div class="stat"><div class="small muted">نسبة الإلغاء</div><div class="v">${c.cancelRate.toFixed(1)}%</div></div><div class="stat"><div class="small muted">أعلى ساعة ضغط</div><div class="v">${c.peakHour}</div></div><div class="stat"><div class="small muted">معدل الأداء العام</div><div class="v">${overallPerf==="—" ? "—" : overallPerf+"%"}</div></div><div class="stat"><div class="small muted">Last Day Orders</div><div class="v">${c.daily.today.orders||0}</div></div>`;
  document.getElementById("highlights").innerHTML=`<div class="card" style="cursor:pointer"><div class="small muted">أفضل محضّر</div><div style="font-size:22px;font-weight:800;margin-top:8px">${c.bestKpiPreparer?.name||"—"}</div><div class="small muted" style="margin-top:6px">${c.bestKpiPreparer ? ('KPI ' + (c.kpiPreparerMonth||'') + ': ' + Number(c.bestKpiPreparer.pct||0).toFixed(1) + '%') : 'لا يوجد KPI'}</div></div><div class="card" style="cursor:pointer" onclick="quickFilter('drivers', ${JSON.stringify(c.drivers[0]?.name||'')})"><div class="small muted">أفضل طيار</div><div style="font-size:22px;font-weight:800;margin-top:8px">${c.drivers[0]?.name||"—"}</div><div class="small muted" style="margin-top:6px">أداء: ${c.drivers[0]?.score?.toFixed(1)||0}%</div></div><div class="card" style="cursor:pointer" onclick="setSearchValue(${JSON.stringify(c.topCustomer?.name||'')})"><div class="small muted">أعلى عميل</div><div style="font-size:22px;font-weight:800;margin-top:8px">${c.topCustomer?.name||"—"}</div><div class="small muted" style="margin-top:6px">${c.topCustomer?.orders||0} طلب</div></div><div class="card" style="cursor:pointer" onclick="quickFilter('branches', ${JSON.stringify(c.branches[0]?.name||'')})"><div class="small muted">أفضل فرع</div><div style="font-size:22px;font-weight:800;margin-top:8px">${c.branches[0]?.name||"—"}</div><div class="small muted" style="margin-top:6px">أداء: ${c.branches[0]?.score?.toFixed(1)||0}%</div></div>`;
  document.getElementById("tabs").innerHTML=tabsHtml();
  document.getElementById("tabContent").innerHTML=renderTabContent(c);
  document.getElementById("countText").textContent=(state.lang==='en' ? "Records Count: " : "عدد السجلات: ")+c.data.length;
  updatePrintReport(c);
  updateBoardReport(c);
  document.getElementById("ordersCard").classList.toggle('hidden', !state.showOrders);
  document.getElementById("toggleOrdersBtn").textContent=state.showOrders?'إخفاء جدول الأوردرات':'إظهار جدول الأوردرات';
  document.getElementById("tableBody").innerHTML=c.data.map(o=>`<tr><td>${o.orderNumber}</td><td>${o.orderDate}</td><td>${o.orderTime}</td><td>${o.branch||''}</td><td>${''||''}</td><td>${o.platform||''}</td><td>${o.customerName||''}</td><td>${o.customerPhone||''}</td><td>${o.area||''}</td><td>${money(o.amount)}</td><td>${o.paymentMethod||''}</td><td><span class="badge ${o.status==="تم"?"ok":"bad"}">${statusLocal(o.status||'')}</span></td><td>${o.preparer||''}</td><td>${o.driver||''}</td></tr>`).join('');
  applyLanguage();
}
document.getElementById("search").oninput=(e)=>{state.search=e.target.value; render()};
document.getElementById("toggleOrdersBtn").onclick=()=>{state.showOrders=!state.showOrders; render()};
document.getElementById("themeBtn").onclick=()=>{document.body.classList.toggle('light'); document.getElementById('themeBtn').textContent=document.body.classList.contains('light')?'Dark Mode':'Light Mode'};
document.getElementById("clearDataBtn").onclick=clearAllData;
document.getElementById("resetFiltersBtn").onclick=()=>{state.filters={years:[], months:[], days:[], branches:[], statuses:[], preparers:[], drivers:[], areas:[], platforms:[], payments:[]}; state.search=''; document.getElementById('search').value=''; render()};

function requiredExcelColumns(){
  return ["مرجع الطلب","تاريخ الطلب","المستودع","علامات التصنيف","الإجمالي","حالة الفاتورة"];
}
function optionalExcelColumns(){
  return ["عنوان الفاتورة","العميل/الهاتف المحمول","العميل/رقم الهاتف","العميل/الشارع","العميل/الشارع 2","Payment Method(POS)(1)","Payment Method(POS)(2)","Prepared By","Delivery Man/اسم العرض"];
}
function renderUploadStatus(type, title, lines){
  const cls = type==='ok' ? 'upload-status upload-ok' : type==='warn' ? 'upload-status upload-warn' : 'upload-status upload-bad';
  const body = Array.isArray(lines) && lines.length ? `<ul class="upload-list">${lines.map(x=>`<li>${x}</li>`).join('')}</ul>` : '';
  document.getElementById('uploadStatus').innerHTML = `<div class="${cls}"><strong>${title}</strong>${body}</div>`;
}
function validateExcelColumns(columns){
  const required = requiredExcelColumns();
  const optional = optionalExcelColumns();
  const missingRequired = required.filter(c => !columns.includes(c));
  const missingOptional = optional.filter(c => !columns.includes(c));
  return {missingRequired, missingOptional};
}


function dedupeOrdersByNumber(rows){
  const map = new Map();
  for(const row of rows){
    const key = String(row.orderNumber || '').trim();
    if(!key){
      map.set('no_key_'+map.size, row);
      continue;
    }
    if(!map.has(key)){
      map.set(key, row);
    } else {
      const prev = map.get(key);
      const prevScore = Number(!!prev.amount) + Number(!!prev.customerName) + Number(!!prev.branch) + Number(!!prev.orderDate);
      const curScore = Number(!!row.amount) + Number(!!row.customerName) + Number(!!row.branch) + Number(!!row.orderDate);
      if(curScore >= prevScore){
        map.set(key, row);
      }
    }
  }
  return Array.from(map.values());
}
async function parseExcelFiles(files){
  const parsed = [];
  const infoLines = [];
  let hadMissingOptional = false;
  for(const f of files){
    const arr = await f.arrayBuffer();
    const wb = XLSX.read(arr,{type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws,{defval:''});
    const columns = json.length ? Object.keys(json[0]) : [];
    const validation = validateExcelColumns(columns);

    if(!json.length){
      throw new Error('الملف فارغ: ' + f.name);
    }
    if(validation.missingRequired.length){
      throw new Error('أعمدة أساسية ناقصة في ' + f.name + ': ' + validation.missingRequired.join(' | '));
    }

    const transformed = transformExcelRows(json);
    parsed.push(...transformed);
    infoLines.push(`اسم الملف: ${f.name} — عدد السجلات: ${transformed.length}`);
    if(validation.missingOptional.length){
      hadMissingOptional = true;
    }
  }
  return {parsed, infoLines, hadMissingOptional};
}
async function handleExcelUpload(mode){
  const files = Array.from(document.getElementById('excelInput').files || []);
  if(!files.length){
    document.getElementById('msg').textContent='اختر ملف Excel واحد على الأقل';
    renderUploadStatus('bad','لم يتم رفع ملف بعد',['اختر ملف Excel أو أكثر من جهازك ثم اضغط الزر المناسب.']);
    return;
  }
  try{
    document.getElementById('msg').textContent='جاري قراءة الملف/الملفات...';
    const beforeCount = state.orders.length;
    const {parsed, infoLines, hadMissingOptional} = await parseExcelFiles(files);

    let merged = [];
    let duplicatesRemoved = 0;

    if(mode === 'replace'){
      merged = dedupeOrdersByNumber(parsed);
      duplicatesRemoved = parsed.length - merged.length;
      state.orders = merged;
    } else {
      const combined = [...state.orders, ...parsed];
      merged = dedupeOrdersByNumber(combined);
      duplicatesRemoved = combined.length - merged.length;
      state.orders = merged;
    }

    const fileNames = files.map(f => f.name).join(' | ');
    const finalCount = state.orders.length;

    const statusLines = [
      'أسماء الملفات: ' + fileNames,
      'عدد الملفات: ' + files.length,
      'عدد السجلات المقروءة من الملفات الجديدة: ' + parsed.length,
      'عدد السجلات قبل العملية: ' + beforeCount,
      'عدد المكررات التي تم حذفها: ' + Math.max(0, duplicatesRemoved),
      'العدد النهائي بعد الدمج: ' + finalCount,
      ...infoLines
    ];

    if(mode === 'replace'){
      statusLines.splice(3, 0, 'نوع العملية: استبدال كامل للبيانات الحالية');
    } else {
      statusLines.splice(3, 0, 'نوع العملية: إضافة ودمج مع البيانات الحالية');
    }

    if(hadMissingOptional){
      renderUploadStatus('warn','تمت معالجة الملفات بنجاح مع بعض الأعمدة الاختيارية الناقصة', statusLines);
    } else {
      renderUploadStatus('ok', mode === 'replace' ? 'تم استبدال البيانات بنجاح' : 'تم دمج الملفات بنجاح', statusLines);
    }

    document.getElementById('msg').textContent = mode === 'replace'
      ? 'تم استبدال البيانات الحالية بنجاح: ' + finalCount + ' سجل'
      : 'تمت إضافة ودمج الملفات بنجاح: ' + finalCount + ' سجل';

    render();
  }catch(err){
    document.getElementById('msg').textContent='حدث خطأ أثناء قراءة ملفات Excel.';
    renderUploadStatus('bad','تعذر معالجة الملف/الملفات',[String(err.message || err)]);
  }
}


document.getElementById("loginBtn").onclick=()=>{};
document.getElementById("loginPass").addEventListener('keydown', e=>{});

document.getElementById("addExcelBtn").onclick=async()=>{ await handleExcelUpload('add'); };
document.getElementById("replaceExcelBtn").onclick=async()=>{ await handleExcelUpload('replace'); };

document.getElementById('appView').classList.add('hidden');
document.getElementById('loginView').classList.remove('hidden');
function attemptLogin(){
  const selectedRole = document.getElementById('loginRole').value;
  const pass = (document.getElementById('loginPass').value || '').trim();
  const adminPasses = [String(state.settings.password || ''), String(DEFAULT_PASSWORD || '17512'), '17512'];
  const viewerPasses = [String(state.settings.viewerPassword || ''), '12345'];
  const ok = (selectedRole === 'admin' && adminPasses.includes(pass)) || (selectedRole === 'viewer' && viewerPasses.includes(pass));
  if(!ok){
    document.getElementById('loginMsg').innerHTML = `<span class="badge bad">${state.lang === 'en' ? 'Incorrect password' : 'كلمة المرور غير صحيحة'}</span>`;
    return;
  }
  state.sessionRole = selectedRole;
  if(!isAdmin()){ state.activeTab = 'overview'; }
  document.getElementById('loginMsg').textContent = '';
  document.getElementById('loginView').classList.add('hidden');
  document.getElementById('appView').classList.remove('hidden');
  render();
}
setLoginRole('admin');
document.getElementById('loginBtn').onclick = attemptLogin;
document.getElementById('loginPass').addEventListener('keydown', function(e){ if(e.key==='Enter') attemptLogin(); });

render();