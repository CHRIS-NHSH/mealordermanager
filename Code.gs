// ==========================================
// 臺北市立內湖高級中學 餐盒訂購系統
// Code.gs
// ==========================================

// ── Script Properties ──
var PROP_SS_ID     = 'spreadsheet_id';
var PROP_ACTIVE_YR = 'active_year';

// ── 分頁名稱 ──
var SH = {
  OFFICES:         'offices',
  VENDORS:         'vendors',
  BUYERS:          'buyers',
  LOCATIONS:       'locations',
  ADMINS:          'admins',
  PURPOSES:        'purposes',
  CONFIG:          'config',
};

// ── 訂單欄位 (0-indexed) ──
var O = {
  ID:0, OFFICE:1, EXT:2, VENDOR:3, BUYER:4,
  DATE:5, TIME:6, LOCATION:7,
  MEAT:8, VEG:9, TOTAL:10, PRICE:11, AMOUNT:12,
  STATUS:13,
  STATUS_PENDING_AT:14, STATUS_ORDERED_AT:15,
  STATUS_SETTLED_AT:16, STATUS_CANCELLED_AT:17, STATUS_NOTIFIED_AT:18,
  MODIFIED:19,
  ORDER_PAY_METHOD:20, ORDER_NOTE:21, ORDER_PURPOSE:22,
  REQ_NO:23,      // 請購單號
  REQ_DATE:24,    // 請購日期
  INVOICE:25, PAY_DATE:26, PAY_METHOD:27, NOTE:28,
  CREATED:29,
};

var ORDER_HEADERS = [
  'ID','處室','組別分機','廠商','訂購人','送餐日期','送餐時間','送餐地點',
  '葷食','素食','總數','單價','總計','狀態',
  '待確認時間','已訂購時間','已核銷時間','已取消時間','已通知出納時間',
  '已修改','付款方式','備註','用途',
  '請購單號','請購日期',
  '發票號碼','付款日期','核銷付款方式','核銷備註','建立時間'
];

var STATUS = {
  PENDING:   '待確認',
  ORDERED:   '已訂購',
  SETTLED:   '已核銷',
  CANCELLED: '已取消',
  NOTIFIED:  '待通知出納',  // 發票已填，等通知出納匯款
};

// ==========================================
// Web App Entry Point
// ==========================================
// 暫存列印資料（用 Script Properties）
function setPrintData(orders) {
  try {
    var key = 'print_' + Session.getActiveUser().getEmail().replace(/[^a-z0-9]/gi,'_');
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(orders));
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPrintData() {
  try {
    var key = 'print_' + Session.getActiveUser().getEmail().replace(/[^a-z0-9]/gi,'_');
    var raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return { success:false, error:'找不到列印資料' };
    return { success:true, orders: JSON.parse(raw) };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPurchaser() {
  try {
    var nc = _getNotifyConfig();
    return { success:true, name: nc.purchaser || '' };
  } catch(e) { return { success:false, error:e.message, name:'' }; }
}

function getPrintPageUrl() {
  try {
    var url = ScriptApp.getService().getUrl() + '?page=print';
    return { success:true, url:url };
  } catch(e) { return { success:false, error:e.message }; }
}

function doGet(e) {
  try {
    // 列印頁
    if (e && e.parameter && e.parameter.page === 'print') {
      return HtmlService.createHtmlOutputFromFile('Print')
        .setTitle('餐盒訂購單｜內湖高中')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var email   = Session.getActiveUser().getEmail();
    var ssId    = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);

    if (!ssId) {
      var t = HtmlService.createTemplateFromFile('Admin');
      t.userEmail = email || ''; t.isAdmin = true; t.connected = false;
      return t.evaluate()
        .setTitle('餐盒訂購系統｜內湖高中')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport','width=device-width,initial-scale=1');
    }

    var isAdmin = _isAdmin(email, SpreadsheetApp.openById(ssId));
    var t = HtmlService.createTemplateFromFile(isAdmin ? 'Admin' : 'View');
    t.userEmail = email || ''; t.isAdmin = isAdmin; t.connected = true;
    return t.evaluate()
      .setTitle('餐盒訂購系統｜內湖高中')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width,initial-scale=1');
  } catch(err) {
    return _errPage('系統錯誤', err.message);
  }
}

function _errPage(title, msg) {
  return HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f0f4f8;">'
    +'<div style="text-align:center;padding:40px;background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.1);">'
    +'<div style="font-size:48px;margin-bottom:12px;">⚠️</div>'
    +'<h2 style="color:#1a3a5c;margin-bottom:8px;">'+title+'</h2>'
    +'<p style="color:#6b7280;">'+msg+'</p>'
    +'</div></body></html>').setTitle(title);
}

// ==========================================
// 部署前執行（授權 + 檢查）
// ==========================================

function setup() {
  var results = [];
  try { DriveApp.getRootFolder();           results.push('✅ Google Drive'); }
  catch(e) { results.push('❌ Google Drive：'+e.message); }
  try { GmailApp.getDrafts();               results.push('✅ Gmail'); }
  catch(e) { results.push('❌ Gmail：'+e.message); }
  try { SpreadsheetApp.getActiveSpreadsheet(); results.push('✅ Google Sheets'); }
  catch(e) { results.push('⚠️ Google Sheets（連結 Sheet 後自動授權）'); }
  try { Session.getActiveUser().getEmail(); results.push('✅ 使用者資訊'); }
  catch(e) { results.push('❌ 使用者資訊：'+e.message); }
  try { ScriptApp.getProjectTriggers();     results.push('✅ 排程觸發器'); }
  catch(e) { results.push('❌ 排程觸發器：'+e.message); }
  var msg = '=== 授權檢查結果 ===\n\n'+results.join('\n')+'\n\n全部 ✅ 後請重新部署 Web App。';
  Logger.log(msg);
}

// ==========================================
// Sheet 連結
// ==========================================
function connectSheet(ssId) {
  try {
    var m = ssId.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) ssId = m[1];
    var ss   = SpreadsheetApp.openById(ssId);
    var name = ss.getName();
    PropertiesService.getScriptProperties().setProperty(PROP_SS_ID, ssId);

    // 建立所有設定分頁
    _initAllSheets(ss);

    // 設定今年為預設年度
    var year = new Date().getFullYear().toString();
    PropertiesService.getScriptProperties().setProperty(PROP_ACTIVE_YR, year);
    _ensureOrderSheet(ss, year);

    // 自動把部署者加為管理員
    var email = Session.getActiveUser().getEmail();
    if (email) {
      var admSheet = ss.getSheetByName(SH.ADMINS);
      var data = admSheet.getDataRange().getValues();
      var found = data.some(function(r){ return String(r[0]).trim()===email.trim(); });
      if (!found) admSheet.appendRow([email]);
    }
    return { success: true, name: name, year: year };
  } catch(e) {
    return { success: false, error: '無法開啟試算表：' + e.message };
  }
}

function getSheetStatus() {
  try {
    var props = PropertiesService.getScriptProperties();
    var ssId  = props.getProperty(PROP_SS_ID);
    var year  = props.getProperty(PROP_ACTIVE_YR);
    if (!ssId) return { success: true, connected: false };
    var ss = SpreadsheetApp.openById(ssId);
    var orderSheets = ss.getSheets()
      .map(function(s){ return s.getName(); })
      .filter(function(n){ return !SH[n.toUpperCase()] && n!==SH.CONFIG; });
    return { success:true, connected:true, name:ss.getName(), url:ss.getUrl(), id:ssId, activeYear:year, years:orderSheets };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 年度管理
// ==========================================
function createYear(year) {
  try {
    if (!_checkAdmin()) return { success: false, error: '無權限' };
    _ensureOrderSheet(_getActiveSS(), year);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function switchYear(year) {
  try {
    if (!_checkAdmin()) return { success: false, error: '無權限' };
    PropertiesService.getScriptProperties().setProperty(PROP_ACTIVE_YR, year);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 首頁統計
// ==========================================
function getDashboardStats() {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
    if (!ssId) return { success: true, connected: false };
    var sheet  = _getActiveOrderSheet();
    var data   = sheet.getDataRange().getValues();
    var today  = _dateStr(new Date());
    var in3    = _addWorkdays(today, 3); // 今天起三個工作天後

    var todayCnt=0, urgentCnt=0, settleCnt=0, futureCnt=0;

    for (var i=1; i<data.length; i++) {
      var r=data[i]; if(!r[O.ID]) continue;
      var status = String(r[O.STATUS]);
      var date   = _parseDate(r[O.DATE]);
      if (!date || status===STATUS.CANCELLED) continue;

      // 今天送餐
      if (date===today) todayCnt++;

      // 三天內需訂購（今天到三個工作天內，待確認）
      if (status===STATUS.PENDING && date>=today && date<=in3) urgentCnt++;

      // 待核銷：已訂購且送餐日已到（現金或轉帳都要處理）
      if (status===STATUS.ORDERED && date<=today) settleCnt++;
      // 待通知出納也算待辦
      if (status===STATUS.NOTIFIED) settleCnt++;

      // 未來待訂購（三個工作天後，待確認）
      if (status===STATUS.PENDING && date>in3) futureCnt++;
    }
    return { success:true, connected:true,
      today:todayCnt, urgent:urgentCnt, settle:settleCnt, future:futureCnt,
      activeYear:_getActiveYear() };
  } catch(e) { return { success: false, error: e.message }; }
}

// 計算 N 個工作天後的日期（排除週六日）
function _addWorkdays(dateStr, n) {
  var d = new Date(dateStr);
  var added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    var dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return _dateStr(d);
}


// ==========================================
// 訂單 CRUD
// ==========================================
function getOrders(filters) {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
    if (!ssId) return { success:true, data:[], year:'', connected:false };
    var sheet = _getActiveOrderSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length<=1) return { success:true, data:[], year:_getActiveYear(), connected:true };
    var orders=[];
    for (var i=1; i<data.length; i++) {
      try {
        if (!data[i][O.ID]) continue;
        orders.push(_rowToOrder(data[i]));
      } catch(rowErr) {
        Logger.log('Row '+i+' error: '+rowErr.message);
      }
    }
    if (filters) {
      if (filters.dateFrom) orders=orders.filter(function(o){return o.date>=filters.dateFrom;});
      if (filters.dateTo)   orders=orders.filter(function(o){return o.date<=filters.dateTo;});
      if (filters.office)   orders=orders.filter(function(o){return o.office===filters.office;});
      if (filters.vendor)   orders=orders.filter(function(o){return o.vendor===filters.vendor;});
      if (filters.buyer)    orders=orders.filter(function(o){return o.buyer===filters.buyer;});
      if (filters.status)   orders=orders.filter(function(o){return o.status===filters.status;});
    }
    orders.sort(function(a,b){
      if(b.date!==a.date) return b.date.localeCompare(a.date);
      return b.created.localeCompare(a.created);
    });
    return { success:true, data:orders, year:_getActiveYear(), connected:true };
  } catch(e) {
    Logger.log('getOrders error: '+e.message);
    return { success:false, error:e.message, data:[], year:'', connected:false };
  }
}

function addOrder(d) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var id=_genId(), now=_now();
    var meat=Number(d.meat)||0, veg=Number(d.veg)||0, price=Number(d.price)||0;
    var row=new Array(30).fill('');
    row[O.ID]=id; row[O.OFFICE]=d.office||''; row[O.EXT]=d.ext||'';
    row[O.VENDOR]=d.vendor||''; row[O.BUYER]=d.buyer||'';
    row[O.DATE]=d.date||''; row[O.TIME]=d.time||''; row[O.LOCATION]=d.location||'';
    row[O.MEAT]=meat; row[O.VEG]=veg; row[O.TOTAL]=meat+veg;
    row[O.PRICE]=price; row[O.AMOUNT]=(meat+veg)*price;
    row[O.STATUS]=STATUS.PENDING; row[O.STATUS_PENDING_AT]=now;
    row[O.MODIFIED]=false;
    row[O.ORDER_PAY_METHOD]=d.payMethod||'現金';
    row[O.ORDER_NOTE]=d.note||'';
    row[O.ORDER_PURPOSE]=d.purpose||'';
    row[O.REQ_NO]=d.reqNo||'';
    row[O.REQ_DATE]=d.reqDate||'';
    row[O.CREATED]=now;
    sheet.appendRow(row);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateOrder(d) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var idx=_findById(data, d.id);
    if (idx<0) return { success:false, error:'找不到訂單' };
    var meat=Number(d.meat)||0, veg=Number(d.veg)||0, price=Number(d.price)||0;
    var rn=idx+1;
    var currentStatus=String(data[idx][O.STATUS]);
    var wasOrdered=(currentStatus===STATUS.ORDERED||currentStatus===STATUS.NOTIFIED);
    sheet.getRange(rn, O.OFFICE+1, 1, 12).setValues([[
      d.office||'', d.ext||'', d.vendor||'', d.buyer||'',
      d.date||'', d.time||'', d.location||'',
      meat, veg, meat+veg, price, (meat+veg)*price
    ]]);
    sheet.getRange(rn, O.ORDER_PAY_METHOD+1).setValue(d.payMethod||'現金');
    sheet.getRange(rn, O.ORDER_NOTE+1).setValue(d.note||'');
    sheet.getRange(rn, O.ORDER_PURPOSE+1).setValue(d.purpose||'');
    sheet.getRange(rn, O.REQ_NO+1).setValue(d.reqNo||'');
    sheet.getRange(rn, O.REQ_DATE+1).setValue(d.reqDate||'');
    if (wasOrdered) {
      sheet.getRange(rn, O.STATUS+1).setValue(STATUS.PENDING);
      sheet.getRange(rn, O.STATUS_PENDING_AT+1).setValue(_now());
      sheet.getRange(rn, O.MODIFIED+1).setValue(true);
    }
    return { success:true, wasOrdered:wasOrdered };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteOrder(id) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var idx=_findById(data, id);
    if (idx<0) return { success:false, error:'找不到訂單' };
    sheet.deleteRow(idx+1);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function markAsOrdered(ids) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var now=_now();
    var ordered=[];
    ids.forEach(function(id) {
      var idx=_findById(data, id); if(idx<0) return;
      if (data[idx][O.STATUS]===STATUS.PENDING) {
        sheet.getRange(idx+1, O.STATUS+1).setValue(STATUS.ORDERED);
        sheet.getRange(idx+1, O.STATUS_ORDERED_AT+1).setValue(now);
        sheet.getRange(idx+1, O.MODIFIED+1).setValue(false);
        data[idx][O.STATUS]=STATUS.ORDERED;
        ordered.push(_rowToOrder(data[idx]));
      }
    });
    var docResult = { skipped: true };
    try { docResult = saveOrdersToPdf(ordered); } catch(e) { Logger.log('saveOrdersToPdf:'+e.message); docResult = { success:false, error:e.message }; }
    return { success:true, orders:ordered, docResult:docResult };
  } catch(e) { return { success:false, error:e.message }; }
}

function cancelOrder(id) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var idx=_findById(data, id);
    if (idx<0) return { success:false, error:'找不到訂單' };
    sheet.getRange(idx+1, O.STATUS+1).setValue(STATUS.CANCELLED);
    sheet.getRange(idx+1, O.STATUS_CANCELLED_AT+1).setValue(_now());
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function settleOrder(id, sd) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var idx=_findById(data, id);
    if (idx<0) return { success:false, error:'找不到訂單' };
    var rn=idx+1;
    // 轉帳付款→待通知出納；現金→已核銷
    var payMethod = String(data[idx][O.ORDER_PAY_METHOD]||'');
    var isTransfer = (payMethod==='轉帳'||payMethod==='代墊－轉帳償還');
    var newStatus = isTransfer ? STATUS.NOTIFIED : STATUS.SETTLED;
    sheet.getRange(rn, O.STATUS+1).setValue(newStatus);
    if (!isTransfer) {
      sheet.getRange(rn, O.STATUS_SETTLED_AT+1).setValue(_now());
    }
    sheet.getRange(rn, O.INVOICE+1).setValue(sd.invoice||'');
    sheet.getRange(rn, O.PAY_DATE+1).setValue(sd.payDate||'');
    sheet.getRange(rn, O.PAY_METHOD+1).setValue(sd.payMethod||'');
    sheet.getRange(rn, O.NOTE+1).setValue(sd.note||'');
    return { success:true, newStatus:newStatus };
  } catch(e) { return { success:false, error:e.message }; }
}

// ==========================================
// 零用金清單
// ==========================================
function getCashOrders() {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
    if (!ssId) return { success:true, upcoming:[], overdue:[], connected:false };
    var sheet  = _getActiveOrderSheet();
    var data   = sheet.getDataRange().getValues();
    var today  = _dateStr(new Date());
    var in7    = _dateStr(new Date(new Date().getTime()+7*24*60*60*1000));
    var upcoming=[], overdue=[];

    for (var i=1; i<data.length; i++) {
      var r=data[i]; if(!r[O.ID]) continue;
      var o      = _rowToOrder(r);
      var status = o.status;
      var date   = o.date;
      var pay    = o.payMethod;
      if (status===STATUS.CANCELLED||status===STATUS.SETTLED) continue;
      if (pay!=='現金') continue;

      // 即將送餐（未來7天內，待確認或已訂購）
      if ((status===STATUS.PENDING||status===STATUS.ORDERED) && date>today && date<=in7) {
        upcoming.push(o);
      }
      // 已送餐待核銷（已訂購，送餐日≤今天）
      if (status===STATUS.ORDERED && date<=today) {
        overdue.push(o);
      }
    }
    upcoming.sort(function(a,b){return a.date.localeCompare(b.date);});
    overdue.sort(function(a,b){return a.date.localeCompare(b.date);});
    return { success:true, upcoming:upcoming, overdue:overdue, connected:true };
  } catch(e) { return { success:false, error:e.message }; }
}


function createPaymentDraft(ids) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var orders=[];
    ids.forEach(function(id){
      var idx=_findById(data,id);
      if(idx>=0) orders.push(_rowToOrder(data[idx]));
    });
    if (!orders.length) return { success:false, error:'沒有選取訂單' };

    var nc=_getNotifyConfig();
    var cashierEmail=nc?nc.cashierEmail:'';
    var byVendor={};
    orders.forEach(function(o){
      if(!byVendor[o.vendor]) byVendor[o.vendor]=[];
      byVendor[o.vendor].push(o);
    });

    var today=_dateStr(new Date());
    var subject='【內湖高中訂餐】匯款明細 '+today;
    var body='您好，\n\n以下為餐盒費用匯款明細，請協助處理，謝謝。\n\n';
    var grandTotal=0;

    Object.keys(byVendor).forEach(function(vendor){
      var vOrders=byVendor[vendor];
      var vTotal=vOrders.reduce(function(s,o){return s+(o.amount||0);},0);
      grandTotal+=vTotal;
      body+='═══════════════════════════\n';
      body+='廠商：'+vendor+'\n應付金額：$'+vTotal+'\n───────────────────────────\n';
      vOrders.forEach(function(o){
        body+='  '+o.date+' '+o.office+' '+o.buyer
          +(o.reqNo?' 【請購單號：'+o.reqNo+'】':'')
          +'｜葷'+o.meat+'+素'+o.veg+'='+o.total+'個'
          +'｜$'+o.price+'×'+o.total+'=$'+o.amount+'\n';
      });
      body+='\n';
    });

    body+='═══════════════════════════\n合計應付：$'+grandTotal+'\n共 '+orders.length+' 筆\n\n請確認後回覆，謝謝。\n\n臺北市立內湖高中 總務處';
    GmailApp.createDraft(cashierEmail||'', subject, body);

    var now=_now();
    ids.forEach(function(id){
      var idx=_findById(data,id); if(idx<0) return;
      if(data[idx][O.STATUS]===STATUS.NOTIFIED){
        sheet.getRange(idx+1,O.STATUS+1).setValue(STATUS.SETTLED);
        sheet.getRange(idx+1,O.STATUS_SETTLED_AT+1).setValue(now);
        sheet.getRange(idx+1,O.STATUS_NOTIFIED_AT+1).setValue(now);
        data[idx][O.STATUS]=STATUS.SETTLED;
      }
    });
    return { success:true, total:grandTotal, count:orders.length };
  } catch(e) { return { success:false, error:e.message }; }
}

// ==========================================
// 設定資料讀取（各自獨立分頁）
// ==========================================
function getConfig() {
  try {
    var ssId=PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
    if (!ssId) return { success:false, error:'no_sheet', offices:[], vendors:[], buyers:[], locations:[], admins:[], connected:false };
    var ss=SpreadsheetApp.openById(ssId);

    var offices   = _readCol(ss, SH.OFFICES,   0);
    var locations = _readCol(ss, SH.LOCATIONS,  0);
    var admins    = _readCol(ss, SH.ADMINS,     0);
    var purposes  = _readCol(ss, SH.PURPOSES,   0);
    var vendors   = _readSheet(ss, SH.VENDORS).map(function(r){
      return {
        name:        String(r[0]||'').trim(),
        price:       String(r[1]||'').trim(),
        orderMethod: String(r[2]||'紙本').trim(),
        phone:       String(r[3]||'').trim(),
        url:         String(r[4]||'').trim(),
        bank:        String(r[5]||'').trim(),
        bankCode:    String(r[6]||'').trim(),
        account:     String(r[7]||'').trim(),
        accountName: String(r[8]||'').trim(),
      };
    }).filter(function(v){return v.name;});
    var buyers    = _readSheet(ss, SH.BUYERS).map(function(r){
      return { name:String(r[0]||'').trim(), office:String(r[1]||'').trim(), ext:String(r[2]||'').trim() };
    }).filter(function(b){return b.name;});

    // 時間從 locations 分頁的 B 欄讀取
    var times = _readSheet(ss, SH.LOCATIONS).map(function(r){
      return String(r[1]||'').trim();
    }).filter(Boolean);
    // 去重
    times = times.filter(function(t,i){return times.indexOf(t)===i;});

    // vendorAccounts 從 vendors 資料建構（向後相容）
    var vendorAccounts = {};
    vendors.forEach(function(v){
      if(v.account) vendorAccounts[v.name]={bank:v.bank,code:v.bankCode,account:v.account,name:v.accountName};
    });

    return { success:true, offices:offices, vendors:vendors, buyers:buyers,
             locations:locations, times:times, admins:admins, purposes:purposes,
             vendorAccounts:vendorAccounts, connected:true };
  } catch(e) {
    return { success:false, error:e.message, offices:[], vendors:[], buyers:[],
             locations:[], times:[], admins:[], connected:false };
  }
}

// ── 各設定分頁 CRUD ──
// 批次儲存所有設定資料（一次寫入）
function saveAllConfig(data) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var ss = _getActiveSS();

    // 重建各分頁
    var maps = {
      office:   { sheet: SH.OFFICES,   rows: (data.offices||[]).map(function(v){return [v];}) },
      location: { sheet: SH.LOCATIONS, rows: (data.locations||[]).map(function(v){return [v];}) },
      purpose:  { sheet: SH.PURPOSES,  rows: (data.purposes||[]).map(function(v){return [v];}) },
      admin:    { sheet: SH.ADMINS,    rows: (data.admins||[]).map(function(v){return [v];}) },
    };

    Object.keys(maps).forEach(function(key) {
      var m = maps[key];
      var sheet = ss.getSheetByName(m.sheet);
      if (!sheet) return;
      // 保留標題行，清除其餘
      var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      sheet.clearContents();
      sheet.appendRow(header);
      m.rows.forEach(function(r){ sheet.appendRow(r); });
    });

    // 廠商（name + price）
    var vendorSheet = ss.getSheetByName(SH.VENDORS);
    if (vendorSheet) {
      var vh = vendorSheet.getRange(1,1,1,vendorSheet.getLastColumn()).getValues()[0];
      vendorSheet.clearContents();
      vendorSheet.appendRow(vh);
      (data.vendors||[]).forEach(function(v){ vendorSheet.appendRow([v.name, v.price||'', v.orderMethod||'紙本', v.phone||'', v.url||'', v.bank||'', v.bankCode||'', v.account||'', v.accountName||'']); });
    }

    // 訂購人（name + office + ext）
    var buyerSheet = ss.getSheetByName(SH.BUYERS);
    if (buyerSheet) {
      var bh = buyerSheet.getRange(1,1,1,buyerSheet.getLastColumn()).getValues()[0];
      buyerSheet.clearContents();
      buyerSheet.appendRow(bh);
      (data.buyers||[]).forEach(function(b){ buyerSheet.appendRow([b.name, b.office||'', b.ext||'']); });
    }

    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function saveConfigItem(type, itemData) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var ss=_getActiveSS();
    if (type==='office')   { _appendToSheet(ss, SH.OFFICES,   [itemData.name]); }
    if (type==='vendor')   { _appendToSheet(ss, SH.VENDORS,   [itemData.name, itemData.price||'', itemData.orderMethod||'紙本', itemData.phone||'', itemData.url||'']); }
    if (type==='buyer')    { _appendToSheet(ss, SH.BUYERS,    [itemData.name, itemData.office||'', itemData.ext||'']); }
    if (type==='location') { _appendToSheet(ss, SH.LOCATIONS, [itemData.name, '']); }
    if (type==='purpose')  { _appendToSheet(ss, SH.PURPOSES,  [itemData.name]); }
    if (type==='time')     {
      // 時間存在 locations 分頁 B 欄
      var sheet=ss.getSheetByName(SH.LOCATIONS)||_ensureSheet(ss,SH.LOCATIONS,['送餐地點','常用時間']);
      var data=sheet.getDataRange().getValues();
      // 找第一個 B 欄空白的行補入，或新增一行
      for (var i=1;i<data.length;i++){
        if (!data[i][1]) { sheet.getRange(i+1,2).setValue(itemData.value); return {success:true}; }
      }
      sheet.appendRow(['', itemData.value]);
    }
    if (type==='admin')    { _appendToSheet(ss, SH.ADMINS,    [itemData.email]); }

    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteConfigItem(type, value) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var ss=_getActiveSS();
    var sheetMap={
      office:SH.OFFICES, vendor:SH.VENDORS, buyer:SH.BUYERS,
      location:SH.LOCATIONS, admin:SH.ADMINS, purpose:SH.PURPOSES
    };
    if (type==='time') {
      // 刪除 locations 分頁 B 欄的時間
      var sheet=ss.getSheetByName(SH.LOCATIONS); if(!sheet) return {success:false,error:'找不到分頁'};
      var data=sheet.getDataRange().getValues();
      for (var i=1;i<data.length;i++){
        if (String(data[i][1]).trim()===String(value).trim()) {
          sheet.getRange(i+1,2).setValue('');
          return {success:true};
        }
      }
      return {success:false,error:'找不到項目'};
    }
    var sheetName=sheetMap[type];
    if (!sheetName) return {success:false,error:'未知類型'};
    var sheet=ss.getSheetByName(sheetName); if(!sheet) return {success:false,error:'找不到分頁'};
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++){
      if (String(data[i][0]).trim()===String(value).trim()) {
        sheet.deleteRow(i+1); return {success:true};
      }
    }
    return {success:false,error:'找不到項目'};
  } catch(e) { return {success:false,error:e.message}; }
}

// ==========================================
// 通知設定（存在 config 分頁，key-value 格式）
// ==========================================
function _getNotifyConfig() {
  try {
    var ss=_getActiveSS();
    var sheet=ss.getSheetByName(SH.CONFIG);
    if (!sheet) return _defaultNotify();
    var data=sheet.getDataRange().getValues();
    var cfg=_defaultNotify();
    data.forEach(function(r){
      var k=String(r[0]||'').trim(), v=String(r[1]||'').trim();
      if (k==='notify_daily')    cfg.daily        = v!=='false';
      if (k==='notify_print')    cfg.print        = v!=='false';
      if (k==='notify_monthly')  cfg.monthly      = v!=='false';
      if (k==='notify_hour')     cfg.hour         = parseInt(v)||8;
      if (k==='notify_email')    cfg.email        = v;
      if (k==='cashier_email')   cfg.cashierEmail = v;
      if (k==='purchaser')       cfg.purchaser    = v;
      if (k==='save_drive')      cfg.drive        = v==='true';
      if (k==='drive_folder_id') cfg.folderId     = v;
    });
    // 預設收件人用第一個管理員
    if (!cfg.email) {
      var admSheet=ss.getSheetByName(SH.ADMINS);
      if (admSheet) {
        var admData=admSheet.getDataRange().getValues();
        for (var i=1;i<admData.length;i++){
          if (admData[i][0]) { cfg.email=String(admData[i][0]); break; }
        }
      }
    }
    return cfg;
  } catch(e) { return _defaultNotify(); }
}

function _defaultNotify() {
  return { daily:true, print:true, monthly:true, hour:8, email:'', cashierEmail:'', purchaser:'', drive:false, folderId:'' };
}

function getNotifyConfig() {
  try { return { success:true, data:_getNotifyConfig() }; }
  catch(e) { return { success:false, error:e.message }; }
}

function saveNotifyConfig(data) {
  try {
    if (!_checkAdmin()) return { success:false, error:'無權限' };
    var ss=_getActiveSS();
    var sheet=ss.getSheetByName(SH.CONFIG);
    if (!sheet) {
      sheet=ss.insertSheet(SH.CONFIG);
      sheet.appendRow(['key','value']);
      sheet.getRange(1,1,1,2).setBackground('#374151').setFontColor('#fff').setFontWeight('bold');
    }
    var vals={
      notify_daily:    String(data.daily),
      notify_print:    String(data.print),
      notify_monthly:  String(data.monthly),
      notify_hour:     String(data.hour||8),
      notify_email:    String(data.email||''),
      cashier_email:   String(data.cashierEmail||''),
      purchaser:       String(data.purchaser||''),
      save_drive:      String(data.drive),
      drive_folder_id: String(data.folderId||''),
    };
    var rows=sheet.getDataRange().getValues();
    var found={};
    for (var i=1;i<rows.length;i++){
      var k=String(rows[i][0]).trim();
      if (vals.hasOwnProperty(k)) { sheet.getRange(i+1,2).setValue(vals[k]); found[k]=true; }
    }
    Object.keys(vals).forEach(function(k){ if(!found[k]) sheet.appendRow([k,vals[k]]); });
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ==========================================
// 每日提醒
// ==========================================
function sendDailyReminder() {
  try {
    var nc=_getNotifyConfig();
    if (!nc||!nc.email) return;
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var today=_dateStr(new Date());
    var in3=_dateStr(new Date(new Date().getTime()+3*24*60*60*1000));
    var todayOrders=[], printReminder=[];
    for (var i=1;i<data.length;i++){
      var r=data[i]; if(!r[O.ID]) continue;
      var o=_rowToOrder(r);
      if (o.status===STATUS.CANCELLED) continue;
      if (nc.daily&&o.date===today) todayOrders.push(o);
      if (nc.print&&o.date===in3&&o.status===STATUS.PENDING) printReminder.push(o);
    }
    if (!todayOrders.length&&!printReminder.length) return;
    var subject='【內湖高中訂餐】每日提醒 '+today;
    var body='';
    if (todayOrders.length){
      body+='📦 今日送餐訂單（'+todayOrders.length+' 筆）\n──────────────────\n';
      todayOrders.forEach(function(o){
        body+='• '+o.office+' '+o.buyer+'｜'+o.vendor+'｜葷'+o.meat+'+素'+o.veg+'='+o.total+'個｜$'+o.amount+'｜'+o.location+'｜'+o.time+'\n';
      });
      body+='\n';
    }
    if (printReminder.length){
      body+='⚠️ 三天後送餐，尚未訂購（'+printReminder.length+' 筆）\n──────────────────\n';
      printReminder.forEach(function(o){
        body+='• '+o.date+' '+o.office+' '+o.buyer+'｜'+o.vendor+'｜共'+o.total+'個｜'+o.location+'\n';
      });
      body+='\n請記得送出訂購單給餐廳！\n';
    }
    body+='\n系統網址：'+ScriptApp.getService().getUrl();
    GmailApp.sendEmail(nc.email, subject, body);
  } catch(e) { Logger.log('sendDailyReminder:'+e.message); }
}

// ==========================================
// 月底未通知出納提醒
// ==========================================
function sendMonthlyReminder() {
  try {
    var nc=_getNotifyConfig();
    if (!nc||!nc.monthly||!nc.email) return;
    if (!_isLastWorkdayOfMonth()) return;
    var sheet=_getActiveOrderSheet();
    var data=sheet.getDataRange().getValues();
    var ym=_dateStr(new Date()).substring(0,7);
    var unnotified=[];
    for (var i=1;i<data.length;i++){
      var r=data[i]; if(!r[O.ID]) continue;
      var o=_rowToOrder(r);
      if (o.date.substring(0,7)===ym&&o.status===STATUS.ORDERED&&o.payMethod==='轉帳') unnotified.push(o);
    }
    if (!unnotified.length) return;
    var subject='【內湖高中訂餐】月底提醒：尚有匯款未通知出納 '+ym;
    var body='⚠️ 本月有 '+unnotified.length+' 筆轉帳訂單尚未通知出納，請盡快處理。\n\n';
    unnotified.forEach(function(o){ body+='• '+o.date+' '+o.office+' '+o.buyer+'｜'+o.vendor+'｜$'+o.amount+'\n'; });
    var total=unnotified.reduce(function(s,o){return s+(o.amount||0);},0);
    body+='\n未通知出納總金額：$'+total+'\n\n系統網址：'+ScriptApp.getService().getUrl();
    GmailApp.sendEmail(nc.email, subject, body);
  } catch(e) { Logger.log('sendMonthlyReminder:'+e.message); }
}

function _isLastWorkdayOfMonth() {
  var today=new Date();
  var last=new Date(today.getFullYear(),today.getMonth()+1,0);
  while(last.getDay()===0||last.getDay()===6) last.setDate(last.getDate()-1);
  return _dateStr(today)===_dateStr(last);
}

// ==========================================
// ==========================================
// 訂購後存 HTML 到 Google Drive
// ==========================================
function saveOrdersToPdf(orders) {
  try {
    var nc=_getNotifyConfig();
    if (!nc||!nc.drive||!nc.folderId) return {success:true,skipped:true};
    if (!orders||!orders.length) return {success:true,skipped:true};

    var folder    = DriveApp.getFolderById(nc.folderId);
    var purchaser = nc.purchaser || '';

    // 依送餐日期分組
    var byDate = {};
    orders.forEach(function(o){
      var d = o.date || _dateStr(new Date());
      if (!byDate[d]) byDate[d] = [];
      byDate[d].push(o);
    });

    // 也把 Sheet 裡同一送餐日已訂購的訂單合入
    var sheet = _getActiveOrderSheet();
    var data  = sheet.getDataRange().getValues();
    Object.keys(byDate).forEach(function(deliveryDate){
      for (var i=1;i<data.length;i++){
        var o=_rowToOrder(data[i]);
        if(!o.id) continue;
        if(o.date===deliveryDate && (o.status===STATUS.ORDERED||o.status===STATUS.NOTIFIED||o.status===STATUS.SETTLED)){
          var found=byDate[deliveryDate].some(function(x){return x.id===o.id;});
          if(!found) byDate[deliveryDate].push(o);
        }
      }
    });

    var fileNames = [];
    Object.keys(byDate).sort().forEach(function(deliveryDate){
      var fileName = '訂購單_'+deliveryDate+'.html';
      var html = _buildPrintHtml(byDate[deliveryDate], purchaser);
      var files = folder.getFilesByName(fileName);
      if (files.hasNext()) {
        files.next().setContent(html);
      } else {
        folder.createFile(fileName, html, 'text/html');
      }
      fileNames.push(fileName);
    });

    return {success:true, fileName:fileNames.join('、')};
  } catch(e) {
    Logger.log('saveOrdersToHtml error: '+e.message);
    return {success:false, error:e.message};
  }
}

function _buildPrintHtml(orders, purchaser) {
  var css='*{margin:0;padding:0;box-sizing:border-box;}'
    +'body{font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',sans-serif;color:#000;}'
    +'.a4-page{width:210mm;height:297mm;display:grid;grid-template-rows:repeat(4,1fr);page-break-after:always;}'
    +'.a4-page:last-child{page-break-after:auto;}'
    +'.slip{border-bottom:2px dashed #999;padding:5mm 8mm 4mm;display:flex;flex-direction:column;gap:5px;overflow:hidden;}'
    +'.slip:last-child{border-bottom:none;}'
    +'.slip-title{font-size:16px;font-weight:900;text-align:center;letter-spacing:2px;margin-bottom:2px;}'
    +'.slip-line{border:none;border-top:2px solid #000;margin:2px 0 4px;}'
    +'.info-grid{display:grid;grid-template-columns:auto 1fr auto 1fr;gap:4px 10px;align-items:baseline;font-size:14px;}'
    +'.lbl{font-weight:700;white-space:nowrap;font-size:13px;}'
    +'.val{border-bottom:1px solid #666;padding-bottom:1px;font-size:14px;font-weight:500;}'
    +'.qty-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:4px;border:1.5px solid #888;border-radius:4px;padding:6px 8px;margin:4px 0;}'
    +'.qty-item{display:flex;flex-direction:column;align-items:center;gap:2px;}'
    +'.qty-lbl{font-size:12px;font-weight:700;color:#333;}'
    +'.qty-val{font-size:22px;font-weight:900;line-height:1;border-bottom:2px solid #333;min-width:40px;text-align:center;padding:0 4px 2px;}'
    +'.qty-unit{font-size:11px;color:#666;}'
    +'.loc-grid{display:grid;grid-template-columns:auto 1fr auto 1fr;gap:4px 10px;align-items:baseline;font-size:14px;}'
    +'.slip-empty{border-bottom:2px dashed #ddd;display:flex;align-items:center;justify-content:center;color:#ddd;font-size:13px;}'
    +'.slip-empty:last-child{border-bottom:none;}';

  var body='';
  for (var i=0;i<orders.length;i+=4){
    var group=orders.slice(i,i+4);
    while(group.length<4) group.push(null);
    body+='<div class="a4-page">';
    group.forEach(function(o){
      if(!o){
        body+='<div class="slip-empty">（空白）</div>';
      } else {
        body+='<div class="slip">';
        body+='<div class="slip-title">臺北市立內湖高級中學 餐盒訂購單</div>';
        body+='<hr class="slip-line">';
        body+='<div class="info-grid">'
          +'<span class="lbl">請購人及分機</span><span class="val">'+_he(o.buyer)+(o.ext?' '+_he(o.ext):'')+'</span>'
          +'<span class="lbl">採購人</span><span class="val">'+_he(purchaser)+'</span>'
          +'</div>';
        body+='<div class="info-grid" style="margin-top:4px;">'
          +'<span class="lbl">送餐日期</span><span class="val">'+_he(o.date)+'</span>'
          +'<span class="lbl">送餐時間</span><span class="val">'+_he(o.time)+'</span>'
          +'</div>';
        body+='<div class="qty-grid">'
          +_qItem('葷食',o.meat,'個')+_qItem('素食',o.veg,'個')+_qItem('共',o.total,'個')
          +_qItem('單價',o.price,'元')+_qItem('總價',o.amount,'元')
          +'</div>';
        body+='<div class="loc-grid">'
          +'<span class="lbl">廠商</span><span class="val">'+_he(o.vendor)+'</span>'
          +'<span class="lbl">送餐地點</span><span class="val">'+_he(o.location)+'</span>'
          +'</div>';
        body+='</div>';
      }
    });
    body+='</div>';
  }

  return '<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">'
    +'<title>訂購單_'+_he(orders[0]&&orders[0].date||'')+'</title>'
    +'<style>'+css+'</style></head><body>'+body+'</body></html>';
}

function _he(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function _qItem(label,val,unit){
  return '<div class="qty-item"><div class="qty-lbl">'+label+'</div>'
    +'<div class="qty-val">'+_he(String(val||0))+'</div>'
    +'<div class="qty-unit">'+unit+'</div></div>';
}
function testDriveFolder(folderId) {
  try {
    if (!_checkAdmin()) return {success:false,error:'無權限'};
    return {success:true, name:DriveApp.getFolderById(folderId).getName()};
  } catch(e) { return {success:false,error:'無法存取資料夾：'+e.message}; }
}

// ==========================================
// 觸發器管理
// ==========================================
function setupTriggers() {
  try {
    if (!_checkAdmin()) return {success:false,error:'無權限'};
    var nc=_getNotifyConfig(); var hour=nc?nc.hour:8;
    _removeTriggers();
    ScriptApp.newTrigger('sendDailyReminder').timeBased().everyDays(1).atHour(hour).create();
    ScriptApp.newTrigger('sendMonthlyReminder').timeBased().everyDays(1).atHour(hour).create();
    return {success:true};
  } catch(e) { return {success:false,error:e.message}; }
}

function removeTriggers() {
  try {
    if (!_checkAdmin()) return {success:false,error:'無權限'};
    _removeTriggers(); return {success:true};
  } catch(e) { return {success:false,error:e.message}; }
}

function _removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    var fn=t.getHandlerFunction();
    if(fn==='sendDailyReminder'||fn==='sendMonthlyReminder') ScriptApp.deleteTrigger(t);
  });
}

function getTriggerStatus() {
  try {
    var daily=false, monthly=false;
    ScriptApp.getProjectTriggers().forEach(function(t){
      if(t.getHandlerFunction()==='sendDailyReminder') daily=true;
      if(t.getHandlerFunction()==='sendMonthlyReminder') monthly=true;
    });
    return {success:true,daily:daily,monthly:monthly};
  } catch(e) { return {success:false,error:e.message}; }
}

// ==========================================
// 建立測試資料
// ==========================================
function createTestData() {
  try {
    var sheet=_getActiveOrderSheet();
    var now=_now(), today=_dateStr(new Date());
    var in3=_dateStr(new Date(new Date().getTime()+3*24*60*60*1000));
    var yday=_dateStr(new Date(new Date().getTime()-1*24*60*60*1000));
    var tests=[
      {office:'教務處',ext:'106',vendor:'連大莊',buyer:'吳宜樺',date:today,time:'上午 11:30',location:'A4簡報室',meat:10,veg:2,price:80,payMethod:'現金',note:'',status:STATUS.ORDERED,orderedAt:now},
      {office:'學務處',ext:'205',vendor:'御華興',buyer:'楊雅芩',date:today,time:'上午 11:30',location:'校長室',meat:5,veg:1,price:100,payMethod:'轉帳',note:'請準時送達',status:STATUS.PENDING,orderedAt:''},
      {office:'總務處',ext:'301',vendor:'連大莊',buyer:'王鳳儀',date:in3,time:'上午 11:30',location:'總務處',meat:18,veg:0,price:80,payMethod:'現金',note:'',status:STATUS.PENDING,orderedAt:''},
      {office:'教務處',ext:'106',vendor:'御華興',buyer:'吳宜樺',date:in3,time:'下午 01:30',location:'A3語言教室',meat:23,veg:1,price:100,payMethod:'轉帳',note:'',status:STATUS.PENDING,orderedAt:''},
      {office:'輔導室',ext:'501',vendor:'連大莊',buyer:'許蕙萍',date:yday,time:'上午 11:30',location:'圓輔室',meat:6,veg:0,price:80,payMethod:'轉帳',note:'',status:STATUS.ORDERED,orderedAt:now},
      {office:'人事室',ext:'702',vendor:'連大莊',buyer:'張甄珍',date:yday,time:'上午 11:30',location:'A4簡報室',meat:16,veg:0,price:80,payMethod:'現金',note:'',status:STATUS.SETTLED,orderedAt:now},
      {office:'學務處',ext:'205',vendor:'連大莊',buyer:'葉永祺',date:in3,time:'上午 11:30',location:'A4簡報室',meat:6,veg:0,price:80,payMethod:'現金',note:'數量有調整',status:STATUS.PENDING,orderedAt:'',modified:true},
    ];
    tests.forEach(function(d){
      var meat=d.meat,veg=d.veg,price=d.price;
      var row=new Array(27).fill('');
      row[O.ID]=_genId(); row[O.OFFICE]=d.office; row[O.EXT]=d.ext;
      row[O.VENDOR]=d.vendor; row[O.BUYER]=d.buyer;
      row[O.DATE]=d.date; row[O.TIME]=d.time; row[O.LOCATION]=d.location;
      row[O.MEAT]=meat; row[O.VEG]=veg; row[O.TOTAL]=meat+veg;
      row[O.PRICE]=price; row[O.AMOUNT]=(meat+veg)*price;
      row[O.STATUS]=d.status; row[O.STATUS_PENDING_AT]=now;
      row[O.STATUS_ORDERED_AT]=d.orderedAt||'';
      row[O.MODIFIED]=d.modified?true:false;
      row[O.ORDER_PAY_METHOD]=d.payMethod; row[O.ORDER_NOTE]=d.note;
      row[O.CREATED]=now;
      if(d.status===STATUS.SETTLED){
        row[O.STATUS_SETTLED_AT]=now;
        row[O.INVOICE]='ZH82795517'; row[O.PAY_DATE]=yday;
        row[O.PAY_METHOD]=d.payMethod;
      }
      sheet.appendRow(row);
    });
    Logger.log('✅ 建立 '+tests.length+' 筆測試資料完成！');
  } catch(e) { Logger.log('❌ '+e.message); }
}

// ==========================================
// 診斷
// ==========================================
function fixSetup() {
  var ssId = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
  if (!ssId) { Logger.log('❌ 尚未連結 Sheet'); return; }
  var ss = SpreadsheetApp.openById(ssId);

  // 建立所有缺少的分頁
  _initAllSheets(ss);
  Logger.log('✅ 所有設定分頁已建立');

  // ── 補欄位遷移 ──

  // 1. buyers 加分機欄
  var buyerSheet = ss.getSheetByName(SH.BUYERS);
  if (buyerSheet) {
    var bHeaders = buyerSheet.getRange(1,1,1,buyerSheet.getLastColumn()).getValues()[0];
    if (bHeaders.length < 3 || String(bHeaders[2]).trim() !== '分機') {
      buyerSheet.getRange(1,1,1,1).setValues([['訂購人姓名']]);
      if (buyerSheet.getLastColumn() < 2) buyerSheet.insertColumnAfter(1);
      if (buyerSheet.getLastColumn() < 3) buyerSheet.insertColumnAfter(2);
      buyerSheet.getRange(1,3).setValue('分機');
      Logger.log('✅ buyers 已加入分機欄');
    } else {
      Logger.log('✅ buyers 分機欄已存在');
    }
  }

  // 2. vendors 補齊所有欄位
  var vendorSheet = ss.getSheetByName(SH.VENDORS);
  if (vendorSheet) {
    var expectedVendorCols = ['廠商名稱','預設單價','訂購方式','聯絡電話','網址','銀行名稱','分行代碼','帳號','戶名'];
    var vLastCol = vendorSheet.getLastColumn();
    var vHeaders = vendorSheet.getRange(1,1,1,vLastCol).getValues()[0].map(function(h){return String(h).trim();});
    expectedVendorCols.forEach(function(col, idx) {
      if (vHeaders.indexOf(col) < 0) {
        var newCol = vendorSheet.getLastColumn() + 1;
        vendorSheet.getRange(1, newCol).setValue(col);
        // 訂購方式補預設值
        if (col === '訂購方式') {
          var lastRow = vendorSheet.getLastRow();
          if (lastRow > 1) {
            var defaults = [];
            for (var i=1; i<lastRow; i++) defaults.push(['紙本']);
            vendorSheet.getRange(2, newCol, lastRow-1, 1).setValues(defaults);
          }
        }
        Logger.log('✅ vendors 已加入「' + col + '」欄');
      }
    });
  }

  // 3. config 補 purchaser
  var configSheet = ss.getSheetByName(SH.CONFIG);
  if (configSheet) {
    var cData = configSheet.getDataRange().getValues();
    var configKeys = ['notify_daily','notify_print','notify_monthly','notify_hour',
                      'notify_email','cashier_email','purchaser','save_drive','drive_folder_id'];
    var existingKeys = cData.map(function(r){return String(r[0]).trim();});
    configKeys.forEach(function(k){
      if (existingKeys.indexOf(k) < 0) {
        configSheet.appendRow([k,'']);
        Logger.log('✅ config 已加入「' + k + '」');
      }
    });
  }

  // 4. 訂單分頁補欄位（請購單號、請購日期）
  var activeYear = _getActiveYear();
  var orderSheet = ss.getSheetByName(activeYear);
  if (orderSheet && orderSheet.getLastColumn() > 0) {
    var oHeaders = orderSheet.getRange(1,1,1,orderSheet.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
    if (oHeaders.indexOf('請購單號') < 0) {
      // 補在「用途」後面
      var purposeIdx = oHeaders.indexOf('用途');
      if (purposeIdx >= 0) {
        orderSheet.insertColumnAfter(purposeIdx + 1);
        orderSheet.insertColumnAfter(purposeIdx + 2);
        orderSheet.getRange(1, purposeIdx+2).setValue('請購單號');
        orderSheet.getRange(1, purposeIdx+3).setValue('請購日期');
        Logger.log('✅ 訂單分頁已加入請購單號、請購日期欄');
      } else {
        Logger.log('⚠️ 找不到「用途」欄，請購單號欄未加入');
      }
    } else {
      Logger.log('✅ 訂單分頁請購欄位已存在');
    }
  }

  // 把當前使用者加為管理員
  var email = Session.getActiveUser().getEmail();
  if (email) {
    var admSheet = ss.getSheetByName(SH.ADMINS);
    var data = admSheet.getDataRange().getValues();
    var found = data.some(function(r){ return String(r[0]).trim() === email.trim(); });
    if (!found) {
      admSheet.appendRow([email]);
      Logger.log('✅ 已將 ' + email + ' 加入管理員');
    } else {
      Logger.log('✅ ' + email + ' 已在管理員名單中');
    }
  }

  var sheets = ss.getSheets().map(function(s){ return s.getName(); });
  Logger.log('所有分頁：' + sheets.join(', '));
  Logger.log('✅ 修復完成！');
  return { success: true };
}

function cleanConfig() {
  var ss = _getActiveSS();
  var sheet = ss.getSheetByName(SH.CONFIG);
  if (!sheet) { Logger.log('config 分頁不存在'); return; }

  var data  = sheet.getDataRange().getValues();
  var keep  = [data[0]]; // 保留標題行
  var oldTypes = ['office','vendor','buyer','location','time','admin','type'];
  var removed  = 0;

  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    // 舊格式的 type/value/extra 列，key 是 office/vendor/buyer 等的刪掉
    if (oldTypes.indexOf(key) >= 0) {
      Logger.log('移除舊資料：' + JSON.stringify(data[i]));
      removed++;
    } else {
      keep.push(data[i]);
    }
  }

  // 重寫整個分頁
  sheet.clearContents();
  keep.forEach(function(r){ sheet.appendRow(r); });
  Logger.log('✅ 清理完成，移除 ' + removed + ' 筆舊資料，保留 ' + (keep.length-1) + ' 筆設定');
}

function migrateOldData() {
  var ss    = _getActiveSS();
  var year  = _getActiveYear();
  var sheet = ss.getSheetByName(year);
  if (!sheet) { Logger.log('找不到 '+year+' 分頁'); return; }

  var data = sheet.getDataRange().getValues();
  Logger.log('共 '+data.length+' 行（含標題）');

  // 舊欄位對應（0-indexed）
  var OLD = {
    ID:0, OFFICE:1, EXT:2, VENDOR:3, BUYER:4,
    DATE:5, TIME:6, LOCATION:7,
    MEAT:8, VEG:9, TOTAL:10, PRICE:11, AMOUNT:12,
    STATUS:13,
    STATUS_PENDING_AT:14,
    STATUS_PRINTED_AT:15,   // 舊：已列印時間
    STATUS_SETTLED_AT:16,
    STATUS_CANCELLED_AT:17,
    INVOICE:18, PAY_DATE:19, PAY_METHOD:20, NOTE:21,
    MODIFIED:22,            // 舊版有些有這欄，有些沒有
    CREATED:22,             // 有 MODIFIED 時 CREATED 在 23
  };

  // 判斷是否已經是新格式（看標題行）
  var header = data[0];
  if (header && String(header[14]) === '已訂購時間') {
    Logger.log('✅ 已是新格式，不需要遷移');
    return;
  }

  // 備份舊分頁
  var backupName = year + '_backup_' + new Date().getTime().toString().slice(-6);
  ss.duplicateActiveSheet();
  var backup = ss.getActiveSheet();
  backup.setName(backupName);
  Logger.log('✅ 已備份到分頁：' + backupName);

  // 清除原分頁並寫入新標題
  sheet.activate();
  sheet.clearContents();
  sheet.appendRow(ORDER_HEADERS);
  sheet.getRange(1,1,1,ORDER_HEADERS.length).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');

  // 轉換每筆資料
  var converted = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[OLD.ID]) continue;

    // 判斷是否有 MODIFIED 欄（舊版部分資料有）
    var hasModified = (data[0] && String(data[0][22]) === '已修改');
    var created = hasModified ? (r[23]||'') : (r[22]||'');

    // 舊狀態對應新狀態
    var status = String(r[OLD.STATUS]||'');
    if (status === '已列印') status = '已訂購';

    var newRow = new Array(27).fill('');
    newRow[O.ID]       = r[OLD.ID];
    newRow[O.OFFICE]   = r[OLD.OFFICE]||'';
    newRow[O.EXT]      = r[OLD.EXT]||'';
    newRow[O.VENDOR]   = r[OLD.VENDOR]||'';
    newRow[O.BUYER]    = r[OLD.BUYER]||'';
    newRow[O.DATE]     = r[OLD.DATE]?String(r[OLD.DATE]).substring(0,10):'';
    newRow[O.TIME]     = r[OLD.TIME]||'';
    newRow[O.LOCATION] = r[OLD.LOCATION]||'';
    newRow[O.MEAT]     = Number(r[OLD.MEAT])||0;
    newRow[O.VEG]      = Number(r[OLD.VEG])||0;
    newRow[O.TOTAL]    = Number(r[OLD.TOTAL])||0;
    newRow[O.PRICE]    = Number(r[OLD.PRICE])||0;
    newRow[O.AMOUNT]   = Number(r[OLD.AMOUNT])||0;
    newRow[O.STATUS]   = status;
    newRow[O.STATUS_PENDING_AT]  = r[OLD.STATUS_PENDING_AT]||'';
    newRow[O.STATUS_ORDERED_AT]  = r[OLD.STATUS_PRINTED_AT]||''; // 已列印時間 → 已訂購時間
    newRow[O.STATUS_SETTLED_AT]  = r[OLD.STATUS_SETTLED_AT]||'';
    newRow[O.STATUS_CANCELLED_AT]= r[OLD.STATUS_CANCELLED_AT]||'';
    newRow[O.STATUS_NOTIFIED_AT] = '';
    newRow[O.MODIFIED]           = false;
    newRow[O.ORDER_PAY_METHOD]   = r[OLD.PAY_METHOD]||'現金';
    newRow[O.ORDER_NOTE]         = r[OLD.NOTE]||'';
    newRow[O.INVOICE]            = r[OLD.INVOICE]||'';
    newRow[O.PAY_DATE]           = r[OLD.PAY_DATE]||'';
    newRow[O.PAY_METHOD]         = r[OLD.PAY_METHOD]||'';
    newRow[O.NOTE]               = r[OLD.NOTE]||'';
    newRow[O.CREATED]            = created;

    sheet.appendRow(newRow);
    converted++;
  }

  Logger.log('✅ 遷移完成！共轉換 '+converted+' 筆資料');
  Logger.log('舊資料備份在分頁：'+backupName);
}

function fixDateFormat() {
  var sheet = _getActiveOrderSheet();
  var data  = sheet.getDataRange().getValues();
  var fixed = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[O.ID]) continue;
    var changed = false;
    // 修正日期
    var newDate = _parseDate(r[O.DATE]);
    if (newDate && newDate !== String(r[O.DATE]).substring(0,10)) {
      sheet.getRange(i+1, O.DATE+1).setValue(newDate);
      changed = true;
    }
    // 修正時間
    var newTime = _parseTime(r[O.TIME]);
    if (newTime && newTime !== String(r[O.TIME])) {
      sheet.getRange(i+1, O.TIME+1).setValue(newTime);
      changed = true;
    }
    if (changed) fixed++;
  }
  Logger.log('✅ 修正 '+fixed+' 筆資料的日期/時間格式');
}

function testGetOrders() {
  var result = getOrders(null);
  Logger.log('result type: ' + typeof result);
  Logger.log('result: ' + JSON.stringify(result));
  if (result && result.data && result.data.length > 0) {
    Logger.log('first order: ' + JSON.stringify(result.data[0]));
  }
}

function diagFull() {
  var ssId = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
  Logger.log('Sheet ID: ' + ssId);

  if (!ssId) { Logger.log('❌ 尚未連結 Sheet'); return; }

  var ss = SpreadsheetApp.openById(ssId);
  Logger.log('Sheet 名稱: ' + ss.getName());

  // 所有分頁
  var sheets = ss.getSheets().map(function(s){ return s.getName(); });
  Logger.log('所有分頁: ' + sheets.join(', '));

  // 當前使用者
  var email = Session.getActiveUser().getEmail();
  Logger.log('當前使用者 email: ' + email);

  // admins 分頁
  var admSheet = ss.getSheetByName(SH.ADMINS);
  if (!admSheet) {
    Logger.log('❌ admins 分頁不存在');
  } else {
    var admData = admSheet.getDataRange().getValues();
    Logger.log('admins 分頁共 ' + admData.length + ' 行:');
    admData.forEach(function(r,i){ Logger.log('  ['+i+']: "'+r[0]+'"'); });
    Logger.log('isAdmin 結果: ' + _isAdmin(email, ss));
  }

  // active year
  var year = PropertiesService.getScriptProperties().getProperty(PROP_ACTIVE_YR);
  Logger.log('目前年度: ' + year);

  // getConfig 結果
  var cfg = getConfig();
  Logger.log('getConfig.success: ' + cfg.success);
  if (!cfg.success) Logger.log('getConfig.error: ' + cfg.error);
  else {
    Logger.log('offices: ' + cfg.offices.length + ' 筆');
    Logger.log('vendors: ' + cfg.vendors.length + ' 筆');
    Logger.log('buyers:  ' + cfg.buyers.length + ' 筆');
    Logger.log('admins:  ' + cfg.admins.join(', '));
  }
}


function diagConfig() {
  var ssId=PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
  Logger.log('Sheet ID: '+ssId);
  if(!ssId){Logger.log('尚未設定');return;}
  var ss=SpreadsheetApp.openById(ssId);
  var sheetNames=ss.getSheets().map(function(s){return s.getName();});
  Logger.log('所有分頁：'+sheetNames.join(', '));
  Object.values(SH).forEach(function(name){
    var sheet=ss.getSheetByName(name);
    if(!sheet){Logger.log(name+': 不存在');return;}
    var data=sheet.getDataRange().getValues();
    Logger.log(name+': '+data.length+' 行');
    data.slice(0,5).forEach(function(r,i){Logger.log('  ['+i+']: '+JSON.stringify(r));});
  });
  Logger.log('getConfig: '+JSON.stringify(getConfig()));
}

// ==========================================
// Sheet 結構建立
// ==========================================
function _initAllSheets(ss) {
  _initSheet(ss, SH.OFFICES,   ['處室名稱'],
    ['校長室','教務處','學務處','教官室','總務處','輔導室','人事室','會計室','圖書館','國際會議廳','多功能教室(A棟3樓)','語言教室(A棟3樓)','A4簡報室']
      .map(function(o){return [o];}));
  _initSheet(ss, SH.VENDORS,   ['廠商名稱','預設單價','訂購方式','聯絡電話','網址','銀行名稱','分行代碼','帳號','戶名'],
    [['連大莊','80','紙本','',''],['御華興','100','電話','',''],['自行採購','','現金','','']]);
  _initSheet(ss, SH.BUYERS,    ['訂購人姓名','所屬處室','分機'], []);
  _initSheet(ss, SH.LOCATIONS, ['送餐地點','常用時間'],
    [['A4簡報室','上午 11:30'],['校長室','下午 01:30'],['教務處',''],['教官室',''],
     ['總務處',''],['輔導室',''],['人事室',''],['會計室',''],['圖書館',''],
     ['國際會議廳',''],['A3多功能教室',''],['A3語言教室',''],
     ['學務處',''],['圓輔室',''],['高一導','']]);
  _initSheet(ss, SH.PURPOSES,  ['用途名稱'],
    [['會議'],['研習']]);
  _initSheet(ss, SH.ADMINS,    ['管理員email'], []);
  // config 分頁（系統設定）
  if (!ss.getSheetByName(SH.CONFIG)) {
    var cfg=ss.insertSheet(SH.CONFIG);
    cfg.appendRow(['key','value']);
    cfg.getRange(1,1,1,2).setBackground('#374151').setFontColor('#fff').setFontWeight('bold');
  }
}

function _initSheet(ss, name, headers, defaultData) {
  var sheet=ss.getSheetByName(name);
  if (!sheet) {
    sheet=ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (defaultData&&defaultData.length) {
      defaultData.forEach(function(r){sheet.appendRow(r);});
    }
  }
  return sheet;
}

function _ensureOrderSheet(ss, year) {
  var sheet=ss.getSheetByName(year);
  if (!sheet) {
    sheet=ss.insertSheet(year);
    sheet.appendRow(ORDER_HEADERS);
    sheet.getRange(1,1,1,ORDER_HEADERS.length).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ==========================================
// 工具函數
// ==========================================
function _getActiveSS() {
  var id=PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
  if(!id) throw new Error('尚未連結 Google Sheet');
  return SpreadsheetApp.openById(id);
}
function _getActiveYear() {
  return PropertiesService.getScriptProperties().getProperty(PROP_ACTIVE_YR)||new Date().getFullYear().toString();
}
function _getActiveOrderSheet() {
  return _ensureOrderSheet(_getActiveSS(), _getActiveYear());
}
function _isAdmin(email, ss) {
  if(!email||!ss) return false;
  var sheet=ss.getSheetByName(SH.ADMINS); if(!sheet) return false;
  var data=sheet.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase().trim()===email.toLowerCase().trim()) return true;
  }
  return false;
}
function _checkAdmin() {
  try {
    var ssId=PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
    if(!ssId) return false;
    return _isAdmin(Session.getActiveUser().getEmail(), SpreadsheetApp.openById(ssId));
  } catch(e){return false;}
}
function _readCol(ss, sheetName, col) {
  var sheet=ss.getSheetByName(sheetName); if(!sheet) return [];
  var data=sheet.getDataRange().getValues();
  return data.slice(1).map(function(r){return String(r[col]||'').trim();}).filter(Boolean);
}
function _readSheet(ss, sheetName) {
  var sheet=ss.getSheetByName(sheetName); if(!sheet) return [];
  var data=sheet.getDataRange().getValues();
  return data.slice(1).filter(function(r){return r[0];});
}
function _appendToSheet(ss, sheetName, rowData) {
  var sheet=ss.getSheetByName(sheetName); if(!sheet) return;
  sheet.appendRow(rowData);
}
function _findById(data, id) {
  for(var i=1;i<data.length;i++){if(String(data[i][O.ID])===String(id))return i;}
  return -1;
}
function _rowToOrder(r) {
  return {
    id:      String(r[O.ID]||''),
    office:  String(r[O.OFFICE]||''),
    ext:     String(r[O.EXT]||''),
    vendor:  String(r[O.VENDOR]||''),
    buyer:   String(r[O.BUYER]||''),
    date:    _parseDate(r[O.DATE]),
    time:    _parseTime(r[O.TIME]),
    location:String(r[O.LOCATION]||''),
    meat:    Number(r[O.MEAT])||0,
    veg:     Number(r[O.VEG])||0,
    total:   Number(r[O.TOTAL])||0,
    price:   Number(r[O.PRICE])||0,
    amount:  Number(r[O.AMOUNT])||0,
    status:  String(r[O.STATUS]||''),
    modified:String(r[O.MODIFIED])==='true',
    payMethod:   String(r[O.ORDER_PAY_METHOD]||''),
    orderNote:   String(r[O.ORDER_NOTE]||''),
    purpose:     String(r[O.ORDER_PURPOSE]||''),
    reqNo:       String(r[O.REQ_NO]||''),
    reqDate:     _toStr(r[O.REQ_DATE]),
    statusPendingAt:   _toStr(r[O.STATUS_PENDING_AT]),
    statusOrderedAt:   _toStr(r[O.STATUS_ORDERED_AT]),
    statusSettledAt:   _toStr(r[O.STATUS_SETTLED_AT]),
    statusCancelledAt: _toStr(r[O.STATUS_CANCELLED_AT]),
    statusNotifiedAt:  _toStr(r[O.STATUS_NOTIFIED_AT]),
    invoice:        String(r[O.INVOICE]||''),
    payDate:        _toStr(r[O.PAY_DATE]),
    settlePayMethod:String(r[O.PAY_METHOD]||''),
    note:           String(r[O.NOTE]||''),
    created:        _toStr(r[O.CREATED]),
  };
}

// 把任何值（包含 Date 物件）轉成純字串
function _toStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    try { return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
    catch(e) { return ''; }
  }
  var s = String(val).trim();
  // 把 ISO 格式轉成易讀格式
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    try {
      var d = new Date(s);
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    } catch(e) { return s.substring(0,16); }
  }
  return s;
}

// 支援多種日期格式 → yyyy-MM-dd 字串
function _parseDate(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    var s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
    // 嘗試解析其他格式
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      d.setFullYear(new Date().getFullYear());
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return s;
  } catch(e) {
    return String(val).substring(0,10);
  }
}

// 支援多種時間格式 → HH:mm 字串
function _parseTime(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
    }
    var s = String(val).trim();
    // 移除秒數：「上午 11:30:00」→「11:30」
    var m = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
    if (!m) return s;
    var h = parseInt(m[1]), min = m[2];
    if (s.indexOf('下午') >= 0 && h < 12) h += 12;
    if (s.indexOf('上午') >= 0 && h === 12) h = 0;
    return String(h).padStart(2,'0') + ':' + min;
  } catch(e) {
    return String(val);
  }
}
function _genId() {
  return 'ORD'+new Date().getTime().toString().slice(-8)+Math.random().toString(36).slice(-3).toUpperCase();
}
function _now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
function _dateStr(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
