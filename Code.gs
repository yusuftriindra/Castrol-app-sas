var SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  try {
    // Mode bridge dipakai oleh frontend GitHub Pages.
    // Bridge ini membuat GitHub Pages tetap dapat memanggil
    // google.script.run tanpa menaruh banner Apps Script di halaman utama.
    if (e && e.parameter && String(e.parameter.bridge || '') === '1') {
      return HtmlService.createHtmlOutput(getGithubBridgeHtml_())
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle('Castrol API Bridge');
    }

    return HtmlService.createTemplateFromFile('Index')
            .evaluate()
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
            .setTitle('Castrol')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Terjadi Kesalahan Server (doGet):</h3><p>' + err.message + '</p><p>Pastikan file HTML bernama persis <b>Index.html</b></p>');
  }
}

/**
 * Halaman bridge tersembunyi untuk GitHub Pages.
 * Tidak menampilkan UI aplikasi. Fungsinya hanya meneruskan
 * request postMessage -> google.script.run -> postMessage.
 */
function getGithubBridgeHtml_() {
  return '<!doctype html><html><head><base target="_top"></head><body>' +
    '<script>' +
    '(function(){' +
    '  window.addEventListener("message", function(event){' +
    '    var msg=event.data||{};' +
    '    if(!msg || msg.type!=="castrol-gas-request" || !msg.id || !msg.method) return;' +
    '    var runner=google.script.run' +
    '      .withSuccessHandler(function(data){' +
    '        event.source.postMessage({type:"castrol-gas-response",id:msg.id,ok:true,data:data}, "*");' +
    '      })' +
    '      .withFailureHandler(function(err){' +
    '        event.source.postMessage({type:"castrol-gas-response",id:msg.id,ok:false,error:{message:(err&&err.message)?err.message:String(err)}},"*");' +
    '      });' +
    '    try {' +
    '      runner[msg.method].apply(runner, Array.isArray(msg.args)?msg.args:[]);' +
    '    } catch(err) {' +
    '      event.source.postMessage({type:"castrol-gas-response",id:msg.id,ok:false,error:{message:err&&err.message?err.message:String(err)}},"*");' +
    '    }' +
    '  });' +
    '  window.parent.postMessage({type:"castrol-gas-ready"},"*");' +
    '})();' +
    '</script></body></html>';
}

function getOrCreateAccountSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName("ACCOUNT");
  if (!sheet) {
    sheet = ss.insertSheet("ACCOUNT");
    sheet.appendRow(["SALES", "USERNAME", "PASSWORD"]);
  }
  return sheet;
}

function prosesLogin(username, password) {
  try {
    var sheet = getOrCreateAccountSheet();
    var data = sheet.getDataRange().getValues();
    var cleanUser = String(username).trim().toLowerCase();
    var cleanPass = String(password).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === cleanUser && String(data[i][2]).trim() === cleanPass) {
        return { success: true, sales: data[i][0] };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch(e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function prosesRegister(namaSales, username, password) {
  try {
    var sheet = getOrCreateAccountSheet();
    var data = sheet.getDataRange().getValues();
    var cleanUser = String(username).trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === cleanUser) {
        return { success: false, message: "Username sudah digunakan!" };
      }
    }
    sheet.appendRow([String(namaSales).trim().toUpperCase(), cleanUser, String(password).trim()]);
    return { success: true, message: "Pendaftaran berhasil! Silakan login." };
  } catch(e) {
    return { success: false, message: "Error: " + e.message };
  }
}


/* ================= PERFORMANCE CACHE ================= */
var APP_CACHE_TTL = 120; // detik
function _cacheVersion() { try { return PropertiesService.getScriptProperties().getProperty('APP_DATA_VERSION') || '1'; } catch(e){ return '1'; } }
function _cacheKey(key) { return 'V'+_cacheVersion()+'_'+key; }
function _cacheGetJson(key) {
  try { var v=CacheService.getScriptCache().get(_cacheKey(key)); return v ? JSON.parse(v) : null; } catch(e){ return null; }
}
function _cachePutJson(key, obj, ttl) {
  try {
    var text=JSON.stringify(obj);
    if (text.length <= 95000) CacheService.getScriptCache().put(_cacheKey(key), text, ttl || APP_CACHE_TTL);
  } catch(e) {}
}
function clearAppDataCaches() {
  try {
    var p=PropertiesService.getScriptProperties();
    var v=Number(p.getProperty('APP_DATA_VERSION')||'1')+1;
    p.setProperty('APP_DATA_VERSION', String(v));
  } catch(e) {}
  try { CacheService.getScriptCache().removeAll([
    'RIWAYAT_ADMIN','RIWAYAT_SALES_', 'MONITORING_ADMIN','MONITORING_SALES_', 'MONITORING_ADMIN_V3','MONITORING_SALES_V3_',
    'KONTRAK_ADMIN','KONTRAK_SALES_', 'PENCAPAIAN_ADMIN','PENCAPAIAN_SALES_',
    'PRODUK_ORDER','OUTLETS_ADMIN','OUTLETS_SALES_','SALES_OPTIONS','CASHBACK_LIST','PRICELIST_DATA_V2'
  ]); } catch(e) {}
}

function getListSalesOptions() { var k='SALES_OPTIONS', c=_cacheGetJson(k); if(c) return c; var r=getListSalesOptions_UNCACHED(); _cachePutJson(k,r,600); return r; }
function getOutletsBySales(salesName) { var k='OUTLETS_SALES_'+String(salesName||'').trim().toUpperCase(); var c=_cacheGetJson(k); if(c) return c; var r=getOutletsBySales_UNCACHED(salesName); _cachePutJson(k,r,600); return r; }
function getAllOutlets() { var k='OUTLETS_ADMIN', c=_cacheGetJson(k); if(c) return c; var r=getAllOutlets_UNCACHED(); _cachePutJson(k,r,600); return r; }
function getListProduk() { var k='PRODUK_ORDER', c=_cacheGetJson(k); if(c) return c; var r=getListProduk_UNCACHED(); _cachePutJson(k,r,600); return r; }
function getListCashback() { var k='CASHBACK_LIST', c=_cacheGetJson(k); if(c) return c; var r=getListCashback_UNCACHED(); _cachePutJson(k,r,600); return r; }

function refreshDerivedForNewOrder_(sheetDamen, sheetInput, idInput, items, selectedCashback, formattedDate) {
  try {
    if (!sheetDamen) return;
    var lastRow = sheetDamen.getLastRow();
    if (lastRow < 2) return;

    // Hitung turunan hanya untuk baris DAMEN milik ID INPUT ini.
    // AL = cashback rate (CASHBACK kolom F, dicari dari key AK) x P (VOLUME)
    // AM = Q - AL
    // AO = tahun dari J (DATE)
    var data = sheetDamen.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h == null ? '' : h).trim().toUpperCase(); });
    var idxId = headers.indexOf('ID INPUT'); if (idxId < 0) idxId = 0;
    var idxDate = headers.indexOf('DATE'); if (idxDate < 0) idxDate = 9;
    var idxVolume = headers.indexOf('VOLUME'); if (idxVolume < 0) idxVolume = 15;
    var idxValueQ = 16; // Q
    var idxAk = 36;     // AK

    // CASHBACK: key ada di kolom L, rate ada di kolom F.
    var cbMap = {};
    var cbSheet = sheetDamen.getParent().getSheetByName('CASHBACK');
    if (cbSheet && cbSheet.getLastRow() >= 2) {
      var cbData = cbSheet.getDataRange().getValues();
      for (var c = 1; c < cbData.length; c++) {
        var cbKey = String(cbData[c][11] == null ? '' : cbData[c][11]).trim();
        if (cbKey) cbMap[cbKey] = Number(cbData[c][5]) || 0;
      }
    }

    var alRows = [], amRows = [], aoRows = [], inputTotal = 0;
    var targetRows = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId] == null ? '' : data[r][idxId]).trim() !== String(idInput || '').trim()) continue;
      targetRows.push(r + 1);

      var ak = String(data[r][idxAk] == null ? '' : data[r][idxAk]).trim();
      var volume = Number(data[r][idxVolume]) || 0;       // P
      var qValue = Number(data[r][idxValueQ]) || 0;       // Q
      var cashbackRate = cbMap[ak] !== undefined ? cbMap[ak] : 0;
      var al = cashbackRate * volume;
      var am = qValue - al;
      var dt = new Date(data[r][idxDate]);
      var ao = isNaN(dt.getTime()) ? '' : Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy');

      alRows.push([al]);
      amRows.push([am]);
      aoRows.push([ao]);
      inputTotal += am;
    }

    for (var i = 0; i < targetRows.length; i++) {
      sheetDamen.getRange(targetRows[i], 38).setValue(alRows[i][0]);
      sheetDamen.getRange(targetRows[i], 39).setValue(amRows[i][0]);
      sheetDamen.getRange(targetRows[i], 41).setValue(aoRows[i][0]);
    }

    // Update INPUT kolom L untuk ID yang sama, bukan berdasarkan lastRow global.
    if (sheetInput && sheetInput.getLastRow() >= 2) {
      var idData = sheetInput.getRange(2, 1, sheetInput.getLastRow() - 1, 1).getValues();
      for (var x = 0; x < idData.length; x++) {
        if (String(idData[x][0] == null ? '' : idData[x][0]).trim() === String(idInput || '').trim()) {
          sheetInput.getRange(x + 2, 12).setValue(inputTotal);
          break;
        }
      }
    }
  } catch(e) {
    Logger.log('refreshDerivedForNewOrder_: '+e.message);
  }
}
function getRiwayatOrder(salesName, usernameInput) { var admin=String(usernameInput||'').trim().toLowerCase()==='admin'; var k=admin?'RIWAYAT_ADMIN':'RIWAYAT_SALES_'+String(salesName||'').trim().toUpperCase(); var c=_cacheGetJson(k); if(c) return c; var r=getRiwayatOrder_UNCACHED(salesName,usernameInput); _cachePutJson(k,r,APP_CACHE_TTL); return r; }
function getMonitoringOutletData(salesName, usernameInput) { var admin=String(usernameInput||'').trim().toLowerCase()==='admin'; var k=admin?'MONITORING_ADMIN_V3':'MONITORING_SALES_V3_'+String(salesName||'').trim().toUpperCase(); var c=_cacheGetJson(k); if(c) return c; var r=getMonitoringOutletData_UNCACHED(salesName,usernameInput); _cachePutJson(k,r,APP_CACHE_TTL); return r; }


/* ================= ADMIN - DATA OUTLET BWS =================
 * CRUD khusus username ADMIN untuk sheet DATA OUTLET BWS.
 * Struktur kolom mengikuti header baris pertama sheet:
 * A CUST.CODE
 * B OUTLET
 * C SALES
 * D CHANNEL OUTLET
 * E START
 * F END
 * G TOTAL TARGET VOLUME
 * H TOTAL TARGET KARTON
 * I TARGET LITER PER BULAN
 * J NAMA OWNER
 * K ID TVD
 * L TARGET KARTON PER BULAN
 * ========================================================== */
function _assertAdminBWS_(usernameInput) {
  if (String(usernameInput == null ? '' : usernameInput).trim().toLowerCase() !== 'admin') {
    throw new Error('Akses ditolak. Menu Outlet BWS hanya dapat digunakan oleh admin.');
  }
}

function _getDataOutletBWSSheet_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheets = ss.getSheets();
  var target = 'DATA OUTLET BWS'.replace(/\s+/g, ' ').trim().toUpperCase();
  for (var i = 0; i < sheets.length; i++) {
    var n = String(sheets[i].getName() || '').replace(/\s+/g, ' ').trim().toUpperCase();
    if (n === target) return sheets[i];
  }
  throw new Error('Sheet "DATA OUTLET BWS" tidak ditemukan.');
}

function _getBWSHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) {
    throw new Error('Sheet DATA OUTLET BWS belum memiliki header.');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return headers.map(function(h) { return String(h == null ? '' : h).trim(); });
}

function _normalizeBWSHeader_(header) {
  return String(header == null ? '' : header)
    .trim()
    .toUpperCase()
    .replace(/[._-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function _isBWSDateHeader_(header) {
  var h = _normalizeBWSHeader_(header);
  return h === 'START' || h === 'END' || h === 'TANGGAL START' || h === 'TANGGAL END';
}

function _isBWSNumberHeader_(header) {
  var h = _normalizeBWSHeader_(header);
  return h.indexOf('TARGET') !== -1 || h.indexOf('VOLUME') !== -1 || h.indexOf('KARTON') !== -1 || h === 'LITER PER BULAN';
}

function _parseBWSDateInput_(value) {
  if (value === null || value === undefined || String(value).trim() === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var text = String(value).trim();
  var m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(text);
  return isNaN(d.getTime()) ? '' : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function _formatBWSCellForClient_(value, header) {
  if (_isBWSDateHeader_(header)) {
    var d = _parseBWSDateInput_(value);
    if (d) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return '';
  }
  return value === null || value === undefined ? '' : String(value);
}

function _getBWSRowsForClient_(sheet, headers) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();
  return values.map(function(row, index) {
    var item = { rowNumber: index + 2, values: [] };
    for (var c = 0; c < headers.length; c++) {
      item.values.push(_formatBWSCellForClient_(row[c], headers[c]));
    }
    return item;
  }).filter(function(item) {
    return item.values.some(function(v) { return String(v || '').trim() !== ''; });
  });
}

function getDataOutletBWSAdmin(usernameInput, salesName) {
  try {
    var isAdmin = String(usernameInput == null ? '' : usernameInput).trim().toLowerCase() === 'admin';
    var sheet = _getDataOutletBWSSheet_();
    var headers = _getBWSHeaders_(sheet);
    var rows = _getBWSRowsForClient_(sheet, headers);

    // Admin melihat seluruh data. Sales hanya melihat Outlet BWS
    // yang SALES-nya sesuai dengan nama sales yang sedang login.
    if (!isAdmin) {
      var salesIdx = -1;
      for (var i = 0; i < headers.length; i++) {
        var nh = _normalizeBWSHeader_(headers[i]);
        if (nh === 'SALES' || nh === 'DSR' || nh === 'NAMA SALES') { salesIdx = i; break; }
      }
      var loginSales = String(salesName == null ? '' : salesName).trim().toUpperCase();
      if (salesIdx === -1) {
        return { success: false, message: 'Kolom SALES pada sheet DATA OUTLET BWS tidak ditemukan.' };
      }
      rows = rows.filter(function(item) {
        return String((item.values || [])[salesIdx] == null ? '' : (item.values || [])[salesIdx]).trim().toUpperCase() === loginSales;
      });
    }

    return { success: true, headers: headers, rows: rows, isAdmin: isAdmin };
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}


/* ================= ADMIN - MASTER FORM OUTLET BWS =================
 * Sumber pilihan form:
 * OUTLET!D = SALES / DSR
 * OUTLET!B = OUTLET
 * OUTLET!E = CHANNEL OUTLET
 * OUTLET!C = CUST.CODE
 *
 * Perhitungan BWS:
 * TARGET LITER PER BULAN = TOTAL TARGET VOLUME / jumlah bulan kalender inklusif
 * TOTAL TARGET KARTON = TOTAL TARGET VOLUME / 9.6 untuk BP
 *                      = TOTAL TARGET VOLUME / 12 untuk CAS
 * TARGET KARTON PER BULAN = TOTAL TARGET KARTON / jumlah bulan
 * ========================================================== */
function _getOutletBWSSourceMaster_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('OUTLET');
  if (!sheet) throw new Error('Sheet "OUTLET" tidak ditemukan.');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { sales: [], map: {} };

  // Sesuai struktur yang diminta: B=OUTLET, C=CUST.CODE, D=SALES, E=CHANNEL OUTLET.
  var salesMap = {};
  var salesOrder = [];
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var outlet = String(row[1] == null ? '' : row[1]).trim();       // B
    var custCode = String(row[2] == null ? '' : row[2]).trim();    // C
    var sales = String(row[3] == null ? '' : row[3]).trim();      // D
    var channel = String(row[4] == null ? '' : row[4]).trim();    // E
    if (!sales && !outlet) continue;

    var salesKey = sales.toUpperCase();
    if (sales && !salesMap[salesKey]) {
      salesMap[salesKey] = sales;
      salesOrder.push(sales);
    }

    if (!map[salesKey]) map[salesKey] = [];
    if (outlet) {
      map[salesKey].push({
        outlet: outlet,
        custCode: custCode,
        channel: channel
      });
    }
  }

  salesOrder.sort(function(a,b){ return String(a).localeCompare(String(b)); });
  Object.keys(map).forEach(function(k){
    var seen={};
    map[k]=map[k].filter(function(x){
      var key=String(x.outlet||'').trim().toUpperCase();
      if (!key || seen[key]) return false;
      seen[key]=true;
      return true;
    }).sort(function(a,b){ return String(a.outlet).localeCompare(String(b.outlet)); });
  });

  return { sales: salesOrder, map: map };
}

function getBWSFormMasterAdmin(usernameInput) {
  try {
    _assertAdminBWS_(usernameInput);
    var master = _getOutletBWSSourceMaster_();
    return { success:true, sales:master.sales, outletMap:master.map };
  } catch(e) {
    return { success:false, message:e && e.message ? e.message : String(e), sales:[], outletMap:{} };
  }
}

function _monthCountInclusiveBWS_(startDate, endDate) {
  if (!startDate || !endDate || endDate.getTime() < startDate.getTime()) return 0;
  return ((endDate.getFullYear() - startDate.getFullYear()) * 12) +
         (endDate.getMonth() - startDate.getMonth()) + 1;
}

function _parseBWSNumber_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  var s=String(value == null ? '' : value).trim();
  if (!s) return 0;
  s=s.replace(/\s/g,'');
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    // Format Indonesia: 13.824,5 -> 13824.5
    s=s.replace(/\./g,'').replace(',', '.');
  } else if (s.indexOf(',') !== -1) {
    s=s.replace(',', '.');
  }
  var n=parseFloat(s.replace(/[^0-9.-]/g,''));
  return isNaN(n) ? 0 : n;
}

function _getBWSCalculation_(startRaw, endRaw, totalVolumeRaw, channelRaw) {
  var start=_parseBWSDateInput_(startRaw);
  var end=_parseBWSDateInput_(endRaw);
  var volume=_parseBWSNumber_(totalVolumeRaw);
  var channel=String(channelRaw == null ? '' : channelRaw).trim().toUpperCase();
  var months=_monthCountInclusiveBWS_(start,end);

  if (!months || volume === 0) {
    return { months:months, targetLiterPerMonth:months ? volume/months : 0, totalTargetKarton:0, targetKartonPerMonth:0 };
  }

  var divisor = channel === 'BP' ? 9.6 : (channel === 'CAS' ? 12 : 0);
  var totalKarton = divisor ? volume/divisor : 0;
  return {
    months:months,
    targetLiterPerMonth:volume/months,
    totalTargetKarton:totalKarton,
    targetKartonPerMonth:totalKarton/months
  };
}

function _recalculateBWSRowValues_(headers, output) {
  var idx={};
  headers.forEach(function(h,i){ idx[_normalizeBWSHeader_(h)] = i; });
  function find(names){
    for (var i=0;i<names.length;i++) if (idx[names[i]] !== undefined) return idx[names[i]];
    return -1;
  }
  var cSales=find(['SALES']);
  var cOutlet=find(['OUTLET']);
  var cCust=find(['CUST CODE','CUSTCODE']);
  var cChannel=find(['CHANNEL OUTLET','CHANNEL']);
  var cStart=find(['START','TANGGAL START']);
  var cEnd=find(['END','TANGGAL END']);
  var cVolume=find(['TOTAL TARGET VOLUME','TARGET VOLUME']);
  var cLiterMonth=find(['TARGET LITER PER BULAN']);
  var cKarton=find(['TOTAL TARGET KARTON']);
  var cKartonMonth=find(['TARGET KARTON PER BULAN']);

  // CUST.CODE tetap mengikuti pasangan Sales + Outlet pada sheet OUTLET.
  // CHANNEL OUTLET dipilih manual oleh admin (BP / CAS), sehingga tidak ditimpa dari sheet OUTLET.
  if (cSales !== -1 && cOutlet !== -1) {
    var master=_getOutletBWSSourceMaster_();
    var salesKey=String(output[cSales]||'').trim().toUpperCase();
    var outletKey=String(output[cOutlet]||'').trim().toUpperCase();
    var candidates=master.map[salesKey] || [];
    var found=null;
    for (var i=0;i<candidates.length;i++) {
      if (String(candidates[i].outlet||'').trim().toUpperCase()===outletKey) { found=candidates[i]; break; }
    }
    if (!found) throw new Error('Outlet "'+String(output[cOutlet]||'')+'" tidak terdaftar pada Sales "'+String(output[cSales]||'')+'" di sheet OUTLET.');
    if (cCust !== -1) output[cCust]=found.custCode || '';
    // Jangan menimpa CHANNEL OUTLET. Nilainya berasal dari pilihan admin BP/CAS.
  }

  var calc=_getBWSCalculation_(cStart!==-1?output[cStart]:'', cEnd!==-1?output[cEnd]:'', cVolume!==-1?output[cVolume]:'', cChannel!==-1?output[cChannel]:'');
  if (cLiterMonth!==-1) output[cLiterMonth]=calc.months ? calc.targetLiterPerMonth : '';
  if (cKarton!==-1) output[cKarton]=calc.months && String(output[cChannel]||'').trim() ? calc.totalTargetKarton : '';
  if (cKartonMonth!==-1) output[cKartonMonth]=calc.months && String(output[cChannel]||'').trim() ? calc.targetKartonPerMonth : '';
  return output;
}

function saveDataOutletBWSAdmin(usernameInput, rowNumber, values) {
  try {
    _assertAdminBWS_(usernameInput);
    var sheet = _getDataOutletBWSSheet_();
    var headers = _getBWSHeaders_(sheet);
    var data = Array.isArray(values) ? values : [];
    var output = [];

    for (var i = 0; i < headers.length; i++) {
      var raw = i < data.length ? data[i] : '';
      var header = headers[i];
      if (_isBWSDateHeader_(header)) {
        var dateValue = _parseBWSDateInput_(raw);
        output.push(dateValue || '');
      } else if (_isBWSNumberHeader_(header)) {
        if (raw === null || raw === undefined || String(raw).trim() === '') output.push('');
        else {
          var numberValue = _parseBWSNumber_(raw);
          output.push(numberValue);
        }
      } else {
        output.push(raw === null || raw === undefined ? '' : String(raw).trim());
      }
    }

    // Server-side: Sales + Outlet menjadi sumber CUST.CODE; CHANNEL OUTLET berasal dari pilihan admin dan target dihitung ulang.
    output = _recalculateBWSRowValues_(headers, output);

    var parsedRow = Number(rowNumber);
    var isEdit = parsedRow >= 2 && parsedRow <= sheet.getMaxRows();
    if (isEdit) {
      sheet.getRange(parsedRow, 1, 1, headers.length).setValues([output]);
    } else {
      parsedRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(parsedRow, 1, 1, headers.length).setValues([output]);
    }

    for (var c = 0; c < headers.length; c++) {
      if (_isBWSDateHeader_(headers[c])) sheet.getRange(parsedRow, c + 1).setNumberFormat('d mmmm yyyy');
    }
    clearAppDataCaches();
    SpreadsheetApp.flush();
    return { success:true, rowNumber:parsedRow, message:isEdit ? 'Data Outlet BWS berhasil diperbarui.' : 'Outlet BWS berhasil ditambahkan.' };
  } catch (e) {
    return { success:false, message:e && e.message ? e.message : String(e) };
  }
}

function deleteDataOutletBWSAdmin(usernameInput, rowNumber) {
  try {
    _assertAdminBWS_(usernameInput);
    var sheet = _getDataOutletBWSSheet_();
    var parsedRow = Number(rowNumber);
    if (!parsedRow || parsedRow < 2 || parsedRow > sheet.getLastRow()) {
      return { success: false, message: 'Baris data tidak valid.' };
    }
    sheet.deleteRow(parsedRow);
    clearAppDataCaches();
    SpreadsheetApp.flush();
    return { success: true, message: 'Outlet BWS berhasil dihapus.' };
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}

function getOutletKontrakData_BWS(salesName, usernameInput) { var admin=String(usernameInput||'').trim().toLowerCase()==='admin'; var k=admin?'KONTRAK_ADMIN':'KONTRAK_SALES_'+String(salesName||'').trim().toUpperCase(); var c=_cacheGetJson(k); if(c) return c; var r=getOutletKontrakData_BWS_UNCACHED(salesName,usernameInput); _cachePutJson(k,r,APP_CACHE_TTL); return r; }
function getPencapaianSales(salesName, usernameInput, selectedPeriod) { var admin=String(usernameInput||'').trim().toLowerCase()==='admin'; var k=(admin?'PENCAPAIAN_ADMIN':'PENCAPAIAN_SALES_'+String(salesName||'').trim().toUpperCase())+'_'+String(selectedPeriod||''); var c=_cacheGetJson(k); if(c) return c; var r=getPencapaianSales_UNCACHED(salesName,usernameInput,selectedPeriod); _cachePutJson(k,r,APP_CACHE_TTL); return r; }

function getListSalesOptions_UNCACHED() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("OUTLET");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colDsr = headers.indexOf("DSR") !== -1 ? headers.indexOf("DSR") : 3;
    
    var salesSet = {};
    for (var j = 1; j < data.length; j++) {
      var val = String(data[j][colDsr]).trim();
      if (val) salesSet[val.toUpperCase()] = true;
    }
    return Object.keys(salesSet).sort();
  } catch(e) {
    return [];
  }
}

function getOutletsBySales_UNCACHED(salesName) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("OUTLET");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    
    var colOutlet = headers.indexOf("OUTLET");
    var colDsr = headers.indexOf("DSR") !== -1 ? headers.indexOf("DSR") : 3;
    var colCust = headers.indexOf("CUST.CODE");
    var colChannel = headers.indexOf("CHANNEL") !== -1 ? headers.indexOf("CHANNEL") : -1;
    
    var outlets = [];
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][colDsr]).trim().toUpperCase() === String(salesName).trim().toUpperCase()) {
        outlets.push({
          outlet: data[j][colOutlet] || "",
          custCode: colCust !== -1 ? data[j][colCust] : "",
          channel: colChannel !== -1 ? data[j][colChannel] : "-"
        });
      }
    }
    return outlets;
  } catch(e) {
    return [];
  }
}

function getAllOutlets_UNCACHED() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("OUTLET");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colOutlet = headers.indexOf("OUTLET");
    var colCust = headers.indexOf("CUST.CODE");
    var colChannel = headers.indexOf("CHANNEL") !== -1 ? headers.indexOf("CHANNEL") : -1;
    
    var outlets = [];
    for (var j = 1; j < data.length; j++) {
      if (data[j][colOutlet]) {
        outlets.push({
          outlet: data[j][colOutlet] || "",
          custCode: colCust !== -1 ? data[j][colCust] : "",
          channel: colChannel !== -1 ? data[j][colChannel] : "-"
        });
      }
    }
    return outlets;
  } catch(e) {
    return [];
  }
}

function getPesananOutletData(forceRefresh, salesName, usernameInput) {
  try {
    var cache = CacheService.getScriptCache();
    var isAdmin = String(usernameInput || '').trim().toLowerCase() === 'admin';
    var salesKey = String(salesName || '').trim().toUpperCase();
    var cacheKey = isAdmin ? 'PO_OUTLET_DATA_V1_ADMIN' : 'PO_OUTLET_DATA_V1_SALES_' + salesKey.replace(/[^A-Z0-9]+/g,'_');
    if (!forceRefresh) {
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    }
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("PO OUTLET");
    if (!sheet) return { success:false, error:'Sheet "PO OUTLET" tidak ditemukan.', headers:[], rows:[] };
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return { success:true, headers:[], rows:[] };
    var data = sheet.getRange(1,1,lastRow,lastCol).getDisplayValues();
    var headers = (data[0] || []).map(function(h,i){ var v=String(h||'').trim(); return v || ('Kolom '+(i+1)); });
    // User non-admin hanya melihat PO OUTLET milik SALES/NAMA SALES/DSR yang sedang login.
    // Admin tetap dapat melihat seluruh data.
    var salesIdx = -1;
    var salesAliases = ['SALES','NAMA SALES','NAMA SALES/DSR','DSR','SALES NAME','NAMA DSR'];
    var headerNorm = headers.map(function(h){ return String(h||'').trim().toUpperCase().replace(/[.\/_-]+/g,' ').replace(/\s+/g,' '); });
    for (var sa=0; sa<salesAliases.length; sa++) {
      var targetHeader = String(salesAliases[sa]).toUpperCase().replace(/[.\/_-]+/g,' ').replace(/\s+/g,' ');
      salesIdx = headerNorm.indexOf(targetHeader);
      if (salesIdx >= 0) break;
    }
    var rows = [];
    for (var i=1;i<data.length;i++) {
      if (!data[i].some(function(v){return String(v||'').trim()!=='';})) continue;
      if (!isAdmin) {
        if (!salesKey || salesIdx < 0) continue;
        var rowSales = String(data[i][salesIdx] == null ? '' : data[i][salesIdx]).trim().toUpperCase();
        if (rowSales !== salesKey) continue;
      }
      rows.push(data[i]);
    }
    var response={success:true,headers:headers,rows:rows,salesFilter:isAdmin?'ADMIN':salesName};
    try{cache.put(cacheKey,JSON.stringify(response),120);}catch(e){}
    return response;
  } catch(e) {
    return {success:false,error:e&&e.message?e.message:String(e),headers:[],rows:[]};
  }
}


function getPesananOutletDetail(poHeaders, poRow) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("ORDERS");
    if (!sheet) return {success:false, message:'Sheet "ORDERS" tidak ditemukan.', headers:[], rows:[]};
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return {success:true, headers:[], rows:[], message:'Sheet ORDERS belum memiliki data.'};

    var ordersData = sheet.getDataRange().getDisplayValues();
    var orderHeaders = (ordersData[0] || []).map(function(h,i){
      var v=String(h||'').trim(); return v || ('Kolom '+(i+1));
    });
    var sourceHeaders = Array.isArray(poHeaders) ? poHeaders : [];
    var sourceRow = Array.isArray(poRow) ? poRow : [];

    function norm(v){
      return String(v==null?'':v).trim().toUpperCase().replace(/[.\\/_-]+/g,' ').replace(/\s+/g,' ');
    }
    function compact(v){ return norm(v).replace(/[^A-Z0-9]/g,''); }
    function keyFamily(h){
      var n=compact(h);
      if(/IDINPUT|INPUTID|NOINPUT/.test(n)) return 'INPUT';
      if(/ORDER|PESANAN|TRANSAKSI|NOPO|IDPO|KODEPO|NOORDER|IDORDER/.test(n)) return 'ORDER';
      if(/CUSTCODE|CUSTOMERCODE|CUSTOMERID/.test(n)) return 'CUST';
      return '';
    }

    // Cari pasangan kolom PO OUTLET -> ORDERS. ID/nomor order menjadi prioritas.
    var pairs=[];
    for(var i=0;i<sourceHeaders.length;i++){
      var sh=sourceHeaders[i], sf=keyFamily(sh), sn=norm(sh), sc=compact(sh);
      if(!sn) continue;
      var oi=-1;
      for(var j=0;j<orderHeaders.length;j++){
        var oh=orderHeaders[j], of=keyFamily(oh);
        if((sf && of && sf===of) || norm(oh)===sn || compact(oh)===sc){oi=j;break;}
      }
      if(oi>=0){
        var val=String(sourceRow[i]==null?'':sourceRow[i]).trim();
        if(val) pairs.push({si:i,oi:oi,value:val,key:!!sf,family:sf});
      }
    }

    var strongPairs=pairs.filter(function(x){return x.key && (x.family==='INPUT'||x.family==='ORDER'||x.family==='CUST');});
    var candidates=[];
    for(var r=1;r<ordersData.length;r++){
      var row=ordersData[r], score=0, matched=0;
      strongPairs.forEach(function(p){
        var a=compact(p.value), b=compact(row[p.oi]);
        if(a && b && a===b){score+=100;matched++;}
      });
      pairs.forEach(function(p){
        if(strongPairs.indexOf(p)!==-1) return;
        var a=compact(p.value), b=compact(row[p.oi]);
        if(a && b && a===b){score+=1;matched++;}
      });
      if(matched>0) candidates.push({row:row,score:score,matched:matched,index:r});
    }

    if(strongPairs.length) candidates=candidates.filter(function(x){return x.score>=100;});
    else candidates=candidates.filter(function(x){return x.matched>=2 || (pairs.length===1 && x.matched===1);});

    candidates.sort(function(a,b){return b.score-a.score || b.matched-a.matched || a.index-b.index;});
    var rows=[];
    if(candidates.length){
      // Untuk satu PO, semua baris ORDERS dengan key utama yang sama adalah detail produk order.
      var top=candidates[0];
      if(strongPairs.length){
        candidates.forEach(function(c){ if(c.score===top.score && c.matched===top.matched) rows.push(c.row); });
      } else {
        rows=candidates.map(function(c){return c.row;});
      }
    }

    // Jika hanya ada satu pasangan key yang kuat, ambil semua baris ORDERS yang benar-benar
    // memiliki nilai key yang sama, supaya seluruh daftar produk order tampil.
    if(rows.length && strongPairs.length){
      var primary=strongPairs[0];
      var primaryVal=compact(primary.value);
      var all=[];
      for(var rr=1;rr<ordersData.length;rr++){
        if(primaryVal && compact(ordersData[rr][primary.oi])===primaryVal) all.push(ordersData[rr]);
      }
      if(all.length) rows=all;
    }

    // Total Volume untuk popup Cashback diambil LANGSUNG dari PO OUTLET.
    // Jangan menghitung ulang dari ORDERS karena nilai Volume PO OUTLET adalah
    // nilai yang harus dijadikan acuan saat user memilih Cashback.
    var totalVolume=0;
    try {
      var volumeIdx=-1;
      var volumeAliases=['TOTAL VOLUME','VOLUME','TOTAL VOL','VOL','VOLUME ORDER'];
      var sourceUp=sourceHeaders.map(function(h){return compact(h);});
      for(var va=0;va<volumeAliases.length;va++){
        var vk=compact(volumeAliases[va]);
        volumeIdx=sourceUp.indexOf(vk);
        if(volumeIdx>=0) break;
      }
      if(volumeIdx>=0){
        var rawVol=String(sourceRow[volumeIdx]==null?'':sourceRow[volumeIdx]).trim();
        if(rawVol){
          var cleaned=rawVol.replace(/[^0-9,.-]/g,'');
          if(cleaned.indexOf(',')>=0 && cleaned.indexOf('.')>=0){
            cleaned=cleaned.replace(/\./g,'').replace(',','.');
          }else if(cleaned.indexOf(',')>=0){
            cleaned=cleaned.replace(',','.');
          }
          totalVolume=parseFloat(cleaned)||0;
        }
      }
    } catch(eVol) { totalVolume=0; }

    // Data yang diperlukan tombol Respon. Mapping sebenarnya dilakukan ulang di server saat aksi.
    return {
      success:true,
      headers:orderHeaders,
      rows:rows,
      matched:rows.length>0,
      totalVolume:totalVolume,
      sourceHeaders:sourceHeaders,
      sourceRow:sourceRow,
      message:rows.length ? '' : 'Order yang sesuai tidak ditemukan di sheet ORDERS.'
    };
  } catch(e) {
    return {success:false,message:e&&e.message?e.message:String(e),headers:[],rows:[]};
  }
}

/**
 * Merespon Pesanan Outlet:
 * 1) PO OUTLET -> INPUT
 * 2) ORDERS -> DAMEN
 *
 * Mapping memakai nama header/alias, sehingga posisi kolom tujuan tetap mengikuti
 * struktur INPUT dan DAMEN yang sudah ada. Tidak memindahkan atau mengubah kolom existing.
 */
function responPesananOutlet(poHeaders, poRow, orderHeaders, orderRows, usernameInput, cashbackOverride, discountOverride) {
  try {
    var ph=Array.isArray(poHeaders)?poHeaders:[], pr=Array.isArray(poRow)?poRow:[];
    var oh=Array.isArray(orderHeaders)?orderHeaders:[], ors=Array.isArray(orderRows)?orderRows:[];
    if(!pr.length) return {success:false,message:'Data PO OUTLET tidak ditemukan.'};
    if(!ors.length) return {success:false,message:'Detail order pada sheet ORDERS tidak ditemukan.'};

    function norm(v){return String(v==null?'':v).trim().toUpperCase().replace(/[.\/_-]+/g,' ').replace(/\s+/g,' ');}
    function compact(v){return norm(v).replace(/[^A-Z0-9]/g,'');}
    function findHeader(headers, aliases){
      var up=headers.map(function(h){return compact(h);});
      for(var i=0;i<aliases.length;i++){
        var a=compact(aliases[i]), idx=up.indexOf(a);
        if(idx>=0) return idx;
      }
      return -1;
    }
    function valBy(headers,row,aliases){
      var i=findHeader(headers,aliases); return i>=0 ? row[i] : '';
    }
    function num(v){
      if(typeof v==='number') return isNaN(v)?0:v;
      var s=String(v==null?'':v).trim().replace(/[^0-9,.-]/g,'');
      if(!s) return 0;
      if(s.indexOf(',')>=0 && s.indexOf('.')>=0) s=s.replace(/\./g,'').replace(',','.');
      else if(s.indexOf(',')>=0) s=s.replace(',','.');
      var n=parseFloat(s); return isNaN(n)?0:n;
    }
    function parseOrderText(v){
      var s=String(v==null?'':v).trim();
      // Kolom N ORDERS adalah sumber detail order.
      // Contoh: Activ Matic 10w-30 0.8L - 6 Dus
      var m=s.match(/^(.*?)(?:\s+-\s*)([0-9]+(?:[.,][0-9]+)?)\s*(PCS|Pcs|DUS|Dus|UNIT|Unit)$/i);
      if(!m) return null;
      return {produk:String(m[1]||'').trim(),qty:num(m[2]),satuan:String(m[3]||'Pcs').trim()};
    }

    // ============================================================
    // PO OUTLET -> payload INPUT
    // Jangan menulis langsung ke INPUT/DAMEN di sini.
    // Gunakan mesin simpanMultiOrder() yang sama dengan tab Input,
    // supaya posisi kolom, rumus volume/value, DISC, cashback,
    // metadata produk dan kolom turunan benar-benar identik.
    // ============================================================
    var outlet=String(valBy(ph,pr,['OUTLET','NAMA OUTLET'])||'').trim();
    var custCode=String(valBy(ph,pr,['CUST.CODE','CUST CODE','CUSTOMER CODE'])||'').trim();
    var sales=String(valBy(ph,pr,['DSR','SALES','SALES NAME','NAMA SALES'])||'Admin').trim();
    var channel=String(valBy(ph,pr,['CHANNEL','CHANNEL OUTLET'])||'').trim();
    var tanggal=String(valBy(ph,pr,['DATE','TANGGAL','TGL'])||'').trim();
    var payment=String(valBy(ph,pr,['PAYMENT','PEMBAYARAN'])||'').trim();
    var note=String(valBy(ph,pr,['NOTE ORDER','NOTE','CATATAN'])||'').trim();
    var discRaw=valBy(ph,pr,['DISC','DISCOUNT','DISKON']);
    // Jika Discount diisi dari popup Aksi Respon, gunakan nilai tersebut.
    if(String(discountOverride==null?'':discountOverride).trim()!=='') discRaw=String(discountOverride).trim()+'%';
    var cashback=String(valBy(ph,pr,['CASHBACK','CASHBACK INTERN','PROGRAM CASHBACK'])||'').trim();
    if(String(cashbackOverride||'').trim()) cashback=String(cashbackOverride).trim();

    // Jika tanggal PO OUTLET ada, pertahankan. Jika kosong, mesin Input memakai tanggal hari ini.
    // simpanMultiOrder saat ini memakai tanggal hari ini sebagai tanggal transaksi.
    // Untuk konsistensi perhitungan, tanggal PO hanya dipakai jika tersedia lewat payload.
    // Ambil master produk yang sama dengan yang dipakai tab Input, terutama HARGA.
    // simpanMultiOrder menerima harga dari frontend, jadi Aksi Respon harus mengisinya
    // dari sheet PRODUK agar VALUE/NET VALUE tidak menjadi 0.
    var produkMaster=[];
    try { produkMaster=getListProduk()||[]; } catch(ignoreProduk) { produkMaster=[]; }
    function cleanProductName(v){
      return String(v==null?'':v).trim().replace(/\s+-\s+\d+(?:[.,]\d+)?\s*(?:PCS|Pcs|DUS|Dus)$/i,'').trim();
    }
    function productKey(v){ return compact(cleanProductName(v)); }
    var produkByCode={}, produkByName={};
    for(var pm=0;pm<produkMaster.length;pm++){
      var pi=produkMaster[pm]||{};
      if(String(pi.kode||'').trim()) produkByCode[compact(pi.kode)]=pi;
      if(String(pi.nama||'').trim()) produkByName[productKey(pi.nama)]=pi;
    }

    var items=[];
    for(var r=0;r<ors.length;r++){
      var row=ors[r]||[];
      // WAJIB ambil daftar order dari ORDERS kolom N (index 13).
      var orderText=row.length>13?String(row[13]||'').trim():'';
      if(!orderText) continue;
      var parsed=parseOrderText(orderText);
      if(!parsed || !parsed.produk || parsed.qty<=0) continue;

      // Bila ORDERS mempunyai KODE PRODUK/SKU, kirim juga ke mesin Input.
      // Namun nama+qty+satuan dari kolom N tetap menjadi sumber order utama.
      var kode=String(valBy(oh,row,['KODE PRODUK','SKU','PRODUCT CODE','PRODUCT SKU','KODE SKU','CODE'])||'').trim();
      var master=(kode && produkByCode[compact(kode)]) || produkByName[productKey(parsed.produk)] || null;
      if(master){
        if(!kode) kode=String(master.kode||'').trim();
      }
      items.push({
        produk: parsed.produk,
        kode: kode,
        qty: parsed.qty,
        satuan: parsed.satuan,
        // HARGA wajib dikirim karena simpanMultiOrder menghitung VALUE
        // dari harga item yang diterima, bukan mengambil HARGA lagi di sana.
        harga: master ? (Number(master.harga)||0) : 0,
        liter: master ? (Number(master.liter)||1) : 1,
        pack: master ? (Number(master.pack)||1) : 1
      });
    }

    if(!items.length){
      return {success:false,message:'Tidak ada order valid di kolom N sheet ORDERS. Format harus seperti "Nama Produk - Qty Dus/Pcs".'};
    }

    var payload={
      outlet: outlet,
      tanggal: tanggal,
      payment: payment,
      disc: discRaw,
      custCode: custCode,
      channel: channel,
      note: note,
      cashback: cashback,
      sales: sales,
      items: items
    };

    // INI KUNCI PERBAIKANNYA:
    // gunakan fungsi yang sama persis dengan tombol Kirim Order pada tab Input.
    // Dengan demikian INPUT dan DAMEN tidak mempunyai logika perhitungan kedua.
    var result=simpanMultiOrder(payload);
    if(result && result.success){
      // Setelah INPUT dan DAMEN berhasil dibuat, hapus sumber PO OUTLET
      // dan detail ORDERS yang benar-benar diproses pada aksi ini.
      // Penghapusan dilakukan berdasarkan baris sumber yang dikirim ke aksi Respon,
      // sehingga tidak mengganggu PO/ORDERS milik order lain.
      var deleteResult = hapusPesananOutletSumber_(ph, pr, oh, ors, result.idInput||'');
      if(!deleteResult.success){
        return {
          success:true,
          idInput:result.idInput||'',
          message:'Berhasil di input. Namun data sumber belum terhapus: '+deleteResult.message
        };
      }
      return {
        success:true,
        idInput:result.idInput||'',
        message:'Berhasil di input.'
      };
    }
    return result || {success:false,message:'Gagal menyimpan respon order.'};
  } catch(e) {
    return {success:false,message:'Gagal memproses Aksi Respon: '+(e&&e.message?e.message:String(e))};
  }
}


/**
 * Menghapus data sumber setelah Aksi Respon berhasil:
 * - 1 baris PO OUTLET yang dipilih
 * - baris ORDERS yang tampil sebagai detail pada PO tersebut
 *
 * Pencocokan menggunakan isi baris sumber yang dikirim dari UI, lalu menghapus
 * indeks baris yang benar-benar cocok. Jika ada baris kembar, hanya sebanyak baris
 * yang diproses yang akan dihapus.
 */
function hapusPesananOutletSumber_(poHeaders, poRow, orderHeaders, orderRows, idInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var poSheet = ss.getSheetByName('PO OUTLET');
    var orderSheet = ss.getSheetByName('ORDERS');
    if (!poSheet) return {success:false,message:'Sheet "PO OUTLET" tidak ditemukan.'};
    if (!orderSheet) return {success:false,message:'Sheet "ORDERS" tidak ditemukan.'};

    // Pastikan ID INPUT hasil respon benar-benar sudah tercatat di DAMEN.
    // Penghapusan sumber hanya dilakukan setelah ID tersebut ditemukan.
    var idKey=String(idInput||'').trim();
    if(!idKey) return {success:false,message:'ID INPUT hasil respon tidak ditemukan, sehingga data sumber tidak dihapus.'};
    var damenSheet=ss.getSheetByName('DAMEN');
    if(!damenSheet || damenSheet.getLastRow()<2) return {success:false,message:'ID INPUT '+idKey+' belum ditemukan di DAMEN, sehingga data sumber tidak dihapus.'};
    var damenLastCol=damenSheet.getLastColumn();
    var damenValues=damenSheet.getRange(2,1,damenSheet.getLastRow()-1,Math.max(1,damenLastCol)).getDisplayValues();
    var idFound=damenValues.some(function(r){return String(r[0]==null?'':r[0]).trim()===idKey;});
    if(!idFound) return {success:false,message:'ID INPUT '+idKey+' belum ditemukan di DAMEN, sehingga data sumber tidak dihapus.'};

    function normCell(v){
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      return String(v == null ? '' : v).trim();
    }
    function rowSignature(row, width){
      var a=[];
      for(var i=0;i<width;i++) a.push(normCell(row && i<row.length ? row[i] : ''));
      return JSON.stringify(a);
    }
    function findMatchingRows(sheet, targetRows, startRow) {
      var lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
      if(lastRow < startRow || lastCol < 1 || !targetRows.length) return [];
      var data=sheet.getRange(startRow,1,lastRow-startRow+1,lastCol).getDisplayValues();
      var used={};
      var found=[];
      targetRows.forEach(function(target){
        var sig=rowSignature(target,lastCol), foundAt=-1;
        for(var i=0;i<data.length;i++){
          if(used[i]) continue;
          if(rowSignature(data[i],lastCol)===sig){ foundAt=i; break; }
        }
        if(foundAt>=0){
          used[foundAt]=true;
          found.push(startRow+foundAt);
        }
      });
      return found;
    }

    // PO OUTLET: cari tepat satu baris yang sama dengan baris yang dipilih.
    var poMatches=findMatchingRows(poSheet,[poRow],2);
    if(!poMatches.length) return {success:false,message:'Baris PO OUTLET yang diproses tidak ditemukan, sehingga data sumber tidak dihapus.'};

    // ORDERS: hapus hanya detail yang benar-benar diproses pada aksi ini.
    var orderTargets=Array.isArray(orderRows)?orderRows:[];
    var orderMatches=findMatchingRows(orderSheet,orderTargets,2);
    if(orderTargets.length && orderMatches.length!==orderTargets.length){
      return {success:false,message:'Sebagian detail ORDERS yang diproses tidak ditemukan, sehingga data sumber tidak dihapus untuk mencegah salah hapus.'};
    }

    // Hapus dari bawah ke atas agar nomor baris tidak bergeser.
    poMatches.sort(function(a,b){return b-a;});
    poMatches.forEach(function(rowNum){poSheet.deleteRow(rowNum);});
    orderMatches.sort(function(a,b){return b-a;});
    orderMatches.forEach(function(rowNum){orderSheet.deleteRow(rowNum);});

    clearAppDataCaches();
    return {success:true,deletedPO:poMatches.length,deletedOrders:orderMatches.length};
  } catch(e) {
    return {success:false,message:e && e.message ? e.message : String(e)};
  }
}

function getPricelistData(forceRefresh) {
  try {
    // Cache server 5 menit agar perpindahan tab tidak membaca Sheet berulang-ulang.
    var cache = CacheService.getScriptCache();
    var cacheKey = 'PRICELIST_DATA_V5';
    if (!forceRefresh) {
      var cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("PRODUK");
    if (!sheet) {
      return { success: false, error: 'Sheet "PRODUK" tidak ditemukan.' };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return { success: true, data: [] };

    // Ambil hanya area yang benar-benar terisi, bukan seluruh kolom Sheet.
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toUpperCase();
    });

    var colNama = headers.indexOf("PRODUK");
    var colLiter = headers.indexOf("LITER");
    var colHarga = headers.indexOf("HARGA");
    var colUpdate = headers.indexOf("UPDATE PRICELIST");

    // SEGMENT PRODUK untuk Pricelist wajib mengikuti KOLOM H pada sheet PRODUK.
    // Kolom H = index 7 (A=0, B=1, ... H=7).
    // Jangan menggunakan pencarian nama header agar hasil filter selalu mengikuti
    // mapping produk yang ada di kolom H.
    var colSegment = 7;

    if (colNama === -1) colNama = 0;
    if (colLiter === -1) colLiter = 1;
    if (colHarga === -1) colHarga = 2;

    var tz = Session.getScriptTimeZone();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var nama = data[i][colNama];
      if (!nama) continue;

      var updateValue = colUpdate !== -1 ? data[i][colUpdate] : "";
      if (updateValue instanceof Date && !isNaN(updateValue.getTime())) {
        updateValue = Utilities.formatDate(updateValue, tz, "yyyy-MM-dd");
      } else {
        updateValue = String(updateValue || '').trim();
      }

      // Hanya kirim field yang benar-benar diperlukan browser.
      result.push({
        nama: String(nama),
        liter: Number(data[i][colLiter]) || 0,
        harga: Number(data[i][colHarga]) || 0,
        updatePricelist: updateValue,
        segment: colSegment !== -1 ? String(data[i][colSegment] == null ? "" : data[i][colSegment]).replace(/\u00A0/g,' ').trim() : ""
      });
    }

    var response = { success: true, data: result };
    // CacheService memiliki batas ukuran; jika terlalu besar, abaikan cache tanpa menggagalkan request.
    try {
      cache.put(cacheKey, JSON.stringify(response), 300);
    } catch (cacheErr) {}

    return response;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getListProduk_UNCACHED() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("PRODUK");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colKode = headers.indexOf("KODE PRODUK") !== -1 ? headers.indexOf("KODE PRODUK") : -1;
    var colNama = headers.indexOf("PRODUK") !== -1 ? headers.indexOf("PRODUK") : 0;
    var colLiter = headers.indexOf("LITER") !== -1 ? headers.indexOf("LITER") : 1;
    var colHarga = headers.indexOf("HARGA") !== -1 ? headers.indexOf("HARGA") : 2;
    var colAktif = headers.indexOf("PRODUK AKTIF");
    
    var produk = [];
    
    for (var j = 1; j < data.length; j++) {
      var namaProduk = data[j][colNama];
      if (!namaProduk) continue;
      
      if (colAktif !== -1) {
        var statusAktif = String(data[j][colAktif]).trim().toUpperCase();
        if (statusAktif !== "ON") {
          continue;
        }
      }
      
      var kodeVal = (colKode !== -1) ? data[j][colKode] : data[j][colNama];
      
      produk.push({ 
        kode: kodeVal || "",
        nama: namaProduk,
        liter: Number(data[j][colLiter]) || 1,
        harga: Number(data[j][colHarga]) || 0
      });
    }
    return produk;
  } catch(e) {
    return [];
  }
}


function getCanonicalCashback_(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('CASHBACK');
    if (!sh || sh.getLastRow() < 2) return raw;
    // CASHBACK diambil MURNI dari kolom A sheet CASHBACK.
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var key = raw.toUpperCase();
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] || '').trim();
      if (v && v.toUpperCase() === key) return v;
    }
  } catch (e) {}
  return raw;
}

function getListCashback_UNCACHED() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetCashback = ss.getSheetByName("CASHBACK");
    if (!sheetCashback) return [];
    
    var lastRowCashback = sheetCashback.getLastRow();
    if (lastRowCashback < 2) return [];
    
    var data = sheetCashback.getRange(2, 1, lastRowCashback - 1, 1).getValues();
    
    var uniqueValues = [];
    var set = new Set();
    
    for (var i = 0; i < data.length; i++) {
      var val = data[i][0];
      if (val !== "" && val !== null && val !== undefined) {
        var cleanVal = String(val).trim();
        if (cleanVal !== "" && !set.has(cleanVal)) {
          set.add(cleanVal);
          uniqueValues.push(cleanVal);
        }
      }
    }
    
    return uniqueValues;
  } catch(e) {
    return [];
  }
}

function updateAccount(oldUsername, newSales, newPassword, newUsername) {
  try {
    var sheet = getOrCreateAccountSheet();
    var data = sheet.getDataRange().getValues();
    var cleanOldUser = String(oldUsername == null ? "" : oldUsername).trim().toLowerCase();
    var cleanNewUser = String(newUsername == null ? "" : newUsername).trim().toLowerCase();
    var passVal = String(newPassword == null ? "" : newPassword).trim();

    if (!cleanOldUser) return { success: false, message: "Username lama tidak ditemukan." };
    if (!cleanNewUser) return { success: false, message: "Username baru tidak boleh kosong." };

    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowUser = String(data[i][1] == null ? "" : data[i][1]).trim().toLowerCase();
      if (rowUser === cleanOldUser) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) return { success: false, message: "Account tidak ditemukan." };

    // Username baru harus unik kecuali sama dengan username lama.
    for (var j = 1; j < data.length; j++) {
      var existingUser = String(data[j][1] == null ? "" : data[j][1]).trim().toLowerCase();
      if (existingUser === cleanNewUser && (j + 1) !== targetRow) {
        return { success: false, message: "Username sudah digunakan." };
      }
    }

    // Nama Sales / DSR TIDAK diubah dari menu Edit Account.
    sheet.getRange(targetRow, 2).setValue(cleanNewUser);
    if (passVal) sheet.getRange(targetRow, 3).setValue(passVal);

    return { success: true, username: cleanNewUser, message: "Account berhasil diperbarui." };
  } catch (e) {
    return { success: false, message: "Gagal memperbarui account: " + e.message };
  }
}


/**
 * Memastikan struktur sheet DAMEN memiliki minimal 40 kolom (AN)
 * dan header posisi AH (34) serta AK (37) tidak kosong.
 * Fungsi ini TIDAK mengubah header lain yang sudah ada.
 */
function ensureDamenStructure_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName("DAMEN");
  if (!sh) return null;

  var requiredCols = 40; // AN
  if (sh.getMaxColumns() < requiredCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), requiredCols - sh.getMaxColumns());
  }

  var headers = sh.getRange(1, 1, 1, requiredCols).getValues()[0];

  // Jangan menghapus header existing. Hanya isi bila kosong.
  if (String(headers[33] || "").trim() === "") headers[33] = "AH";
  if (String(headers[36] || "").trim() === "") headers[36] = "AK";

  sh.getRange(1, 1, 1, requiredCols).setValues([headers]);
  return sh;
}

function simpanMultiOrder(orderPayload) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var items = orderPayload.items;

    // Normalisasi Disc ke nilai desimal persen.
    // 5 atau "5%" menjadi 0.05, sedangkan 0.05 tetap 0.05.
    var discRaw = orderPayload.disc == null ? "" : String(orderPayload.disc).trim().replace(",", ".");
    var discValue = 0;
    if (discRaw !== "") {
      var discHasPercent = discRaw.indexOf("%") !== -1;
      var discNumber = parseFloat(discRaw.replace("%", "").trim());
      if (!isNaN(discNumber)) {
        discValue = discHasPercent || Math.abs(discNumber) > 1 ? discNumber / 100 : discNumber;
      }
    }
    
    if (!items || items.length === 0) {
      return { success: false, message: "Keranjang order masih kosong!" };
    }

    var sheetProduk = ss.getSheetByName("PRODUK");
    var produkInfoMap = {}; 
    
    if (sheetProduk) {
      var prodData = sheetProduk.getDataRange().getValues();
      if (prodData.length > 1) {
        var pHeaders = prodData[0].map(function(h) { return String(h).trim().toUpperCase(); });
        var pKodeIdx = pHeaders.indexOf("KODE PRODUK");
        var pNamaIdx = pHeaders.indexOf("PRODUK") !== -1 ? pHeaders.indexOf("PRODUK") : 0;
        var pLiterIdx = pHeaders.indexOf("LITER") !== -1 ? pHeaders.indexOf("LITER") : 1;
        var pPackIdx = pHeaders.indexOf("PACK") !== -1 ? pHeaders.indexOf("PACK") : (pHeaders.indexOf("ISI") !== -1 ? pHeaders.indexOf("ISI") : -1);
        var pXIdx = pHeaders.indexOf("TYPE PRODUK");
        if (pXIdx === -1) pXIdx = 23; 
        
        for (var p = 1; p < prodData.length; p++) {
          var pKode = pKodeIdx !== -1 ? String(prodData[p][pKodeIdx]).trim().toUpperCase() : "";
          var pNama = String(prodData[p][pNamaIdx]).trim().toUpperCase();
          var pLiterVal = Number(prodData[p][pLiterIdx]) || 1;
          var pPackVal = pPackIdx !== -1 ? (Number(prodData[p][pPackIdx]) || 1) : 1;
          
          var info = { 
            liter: pLiterVal, 
            pack: pPackVal, 
            kode: pKode,
            colD: prodData[p][3] || "",
            colE: prodData[p][4] || "",
            colF: prodData[p][5] || "",
            colG: prodData[p][6] || "",
            colL: prodData[p][11] || "",
            colR: prodData[p][17] || "",
            colS: prodData[p][18] || "",
            colX: prodData[p][pXIdx] || ""
          };

          if (pKode) produkInfoMap[pKode] = info;
          if (pNama) produkInfoMap[pNama] = info;
        }
      }
    }

    var totalVolume = 0;
    var totalValue = 0;

    for (var i = 0; i < items.length; i++) {
      var qtyInput = Number(items[i].qty) || 0;
      var satuan = String(items[i].satuan || "Pcs").trim().toUpperCase();
      var hargaSatuan = Number(items[i].harga) || 0;
      
      var itemKodeInput = String(items[i].kode || "").trim().toUpperCase();
      var itemNama = String(items[i].produk || "").trim().toUpperCase();
      
      var packPerDus = 1;
      var literProduk = Number(items[i].liter) || 1;
      var resolvedKode = itemKodeInput;
      
      items[i].valD = ""; items[i].valE = ""; items[i].valF = "";
      items[i].valG = ""; items[i].valL = ""; items[i].valR = ""; items[i].valS = ""; items[i].valX = "";

      if (itemKodeInput && produkInfoMap[itemKodeInput]) {
        var pInf = produkInfoMap[itemKodeInput];
        literProduk = pInf.liter;
        packPerDus = pInf.pack;
        resolvedKode = pInf.kode || itemKodeInput;
        items[i].valD = pInf.colD; items[i].valE = pInf.colE; items[i].valF = pInf.colF;
        items[i].valG = pInf.colG; items[i].valL = pInf.colL; items[i].valR = pInf.colR; items[i].valS = pInf.colS; items[i].valX = pInf.colX;
      } else if (itemNama && produkInfoMap[itemNama]) {
        var pInf = produkInfoMap[itemNama];
        literProduk = pInf.liter;
        packPerDus = pInf.pack;
        resolvedKode = pInf.kode || "";
        items[i].valD = pInf.colD; items[i].valE = pInf.colE; items[i].valF = pInf.colF;
        items[i].valG = pInf.colG; items[i].valL = pInf.colL; items[i].valR = pInf.colR; items[i].valS = pInf.colS; items[i].valX = pInf.colX;
      } else {
        packPerDus = Number(items[i].pack) || 1;
      }
      
      if (qtyInput <= 0) {
        return { success: false, message: "Terdapat Qty produk yang tidak valid (harus > 0)." };
      }

      var finalPcs = 0;
      var finalDus = "";

      if (satuan === "DUS") {
        finalDus = qtyInput;
        finalPcs = qtyInput * packPerDus; 
      } else {
        finalPcs = qtyInput; 
        finalDus = "";
      }

      var calculatedVolume = finalPcs * literProduk;
      var grossValue = finalPcs * hargaSatuan;
      // Kolom Q (VALUE DAMEN) harus sudah dipotong Disc (kolom O).
      // Disc disimpan sebagai desimal: 5% = 0.05.
      var netValue = grossValue * (1 - discValue);
      if (netValue < 0) netValue = 0;

      totalVolume += calculatedVolume;
      totalValue += netValue;
      
      items[i].finalPcs = finalPcs;
      items[i].finalDus = finalDus;
      items[i].satuanOpt = satuan;
      items[i].calculatedVolume = calculatedVolume;
      items[i].resolvedKode = resolvedKode;
      items[i].netValue = netValue;
    }

    var now = new Date();
    var datePrefix = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMMdd");
    var idInput = "INP-" + datePrefix + "-" + Utilities.getUuid().replace(/-/g, "").substring(0, 12);
    
    var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "d MMMM yyyy");
    var monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
    var yearNum = now.getFullYear();
    
    var monthIndex = now.getMonth(); 
    var quartalNum = Math.floor(monthIndex / 3) + 1;
    var quartalText = "QUARTAL " + quartalNum + " " + yearNum;

    var sheetInput = ss.getSheetByName("INPUT");
    if (!sheetInput) {
      sheetInput = ss.insertSheet("INPUT");
      sheetInput.appendRow(["ID INPUT", "OUTLET", "DATE", "MONTH", "DSR", "PAYMENT", "DISC", "CUST.CODE", "CHANNEL", "NOTE ORDER", "VOLUME", "VALUE", "STATUS APPROVE", "CASHBACK"]);
    }
    
    var selectedCashback = getCanonicalCashback_(orderPayload.cashback);
    var salesName = orderPayload.sales ? String(orderPayload.sales).trim() : "Admin";

    sheetInput.appendRow([
      idInput,
      orderPayload.outlet,
      formattedDate,
      monthName + "-" + String(yearNum).slice(-2),
      salesName,
      orderPayload.payment,
      discValue,
      orderPayload.custCode,
      orderPayload.channel,
      orderPayload.note,
      totalVolume,
      totalValue,
      "PENDING",
      selectedCashback
    ]);

    var sheetDamen = ensureDamenStructure_();
    if (!sheetDamen) return { success: false, message: "Sheet DAMEN tidak ditemukan." };
    var damenRowsToWrite = [];
    
    for (var j = 0; j < items.length; j++) {
      var namaProdukItem = String(items[j].produk || "").trim();
      var trimSkuText = namaProdukItem + " - " + items[j].qty + " " + items[j].satuan;
      
      var valAh = _getProdukKategoriCashbackByKodeQ_(items[j].resolvedKode);
      // AK = CASHBACK + AH + AD. AH diambil dari PRODUK!S (KATEGORI CASHBACK)
      // berdasarkan Kode Produk DAMEN!E yang dicocokkan ke PRODUK!Q.
      var formatProgramAk = String(selectedCashback || "").trim();
      if (valAh) formatProgramAk += " - " + String(valAh).trim();
      if (quartalText) formatProgramAk += " - " + String(quartalText).trim();

      var fullRow = new Array(40).fill("");
      var idDamen = "DAM-" + datePrefix + "-" + Utilities.getUuid().replace(/-/g, "").substring(0, 12);
      
      fullRow[0]  = idInput;
      fullRow[1]  = idDamen;
      fullRow[2]  = orderPayload.outlet;
      fullRow[3]  = orderPayload.custCode;
      fullRow[4]  = items[j].resolvedKode;
      fullRow[5]  = namaProdukItem;
      fullRow[6]  = items[j].qty;
      fullRow[7]  = items[j].finalPcs;
      fullRow[8]  = items[j].finalDus;
      fullRow[9]  = formattedDate;
      fullRow[10] = monthName + "-" + String(yearNum).slice(-2);
      fullRow[11] = salesName;
      fullRow[12] = trimSkuText;
      fullRow[13] = orderPayload.payment;
      fullRow[14] = discValue;
      fullRow[15] = items[j].calculatedVolume;
      fullRow[16] = items[j].netValue;
      fullRow[17] = items[j].satuanOpt;
      
      fullRow[21] = items[j].valD;            
      fullRow[22] = items[j].valE;            
      fullRow[23] = items[j].valF;            
      fullRow[24] = items[j].valG;            

      // Z = PRODUK!H dan AA = PRODUK!I, berdasarkan Kode Produk DAMEN!E -> PRODUK!Q.
      fullRow[25] = _getProdukModelByKodeQ_(items[j].resolvedKode);
      fullRow[26] = _getProdukAAByKodeQ_(items[j].resolvedKode);

      fullRow[28] = monthName;
      fullRow[29] = quartalText;
      fullRow[30] = items[j].valL;
      fullRow[33] = valAh;
      fullRow[36] = formatProgramAk;
      fullRow[39] = items[j].valR;

      damenRowsToWrite.push(fullRow);
    }
    
    if (damenRowsToWrite.length) {
      var firstDamenRow = sheetDamen.getLastRow() + 1;
      sheetDamen.getRange(firstDamenRow, 1, damenRowsToWrite.length, 40).setValues(damenRowsToWrite);
      sheetDamen.getRange(firstDamenRow, 15, damenRowsToWrite.length, 1).setNumberFormat("0%");
    }

    // Update kolom turunan HANYA untuk baris order baru (lebih cepat daripada menghitung seluruh DAMEN/INPUT).
    refreshDerivedForNewOrder_(sheetDamen, sheetInput, idInput, items, selectedCashback, formattedDate);
    clearAppDataCaches();
  
    return { success: true, idInput: idInput, message: "Order berhasil disimpan!" };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan: " + e.message };
  }
}


function _normalizeEditableProductKey_(v) {
  return String(v == null ? '' : v)
    .trim()
    .toUpperCase()
    .replace(/\s+-\s+\d+(?:[.,]\d+)?\s*(?:PCS|DUS)$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}
function _getEditableProductMap_() {
  var ss=SpreadsheetApp.openById(SS_ID), sh=ss.getSheetByName('PRODUK'), map={};
  if(!sh||sh.getLastRow()<2)return map;
  var d=sh.getDataRange().getValues(), hd=d[0].map(function(x){return String(x).trim().toUpperCase();});
  var k=hd.indexOf('KODE PRODUK'),n=hd.indexOf('PRODUK'),l=hd.indexOf('LITER'),p=hd.indexOf('HARGA'),pk=hd.indexOf('PACK'); if(pk<0)pk=hd.indexOf('ISI');
  for(var i=1;i<d.length;i++){
    var name=n>=0?String(d[i][n]||'').trim():'';
    if(!name)continue;
    var kode=k>=0?String(d[i][k]||'').trim():name;
    var inf={nama:name,kode:kode,liter:l>=0?(Number(d[i][l])||1):1,harga:p>=0?(Number(d[i][p])||0):0,pack:pk>=0?(Number(d[i][pk])||1):1,colD:d[i][3]||'',colE:d[i][4]||'',colF:d[i][5]||'',colG:d[i][6]||'',colL:d[i][11]||'',colR:d[i][17]||'',colS:d[i][18]||'',colX:d[i][23]||''};
    var nameKey=_normalizeEditableProductKey_(name);
    var codeKey=String(kode||'').trim().toUpperCase().replace(/\s+/g,' ');
    if(nameKey)map[nameKey]=inf;
    if(codeKey)map[codeKey]=inf;
    // Index tambahan yang mengabaikan tanda baca/spasi/apostrophe.
    var compactCode=_normalizeEditableCodeKey_(kode);
    var compactName=_normalizeEditableNameKey_(name);
    if(compactCode)map['#CODE#'+compactCode]=inf;
    if(compactName)map['#NAME#'+compactName]=inf;
    // Simpan juga bentuk mentah untuk kompatibilitas.
    if(name)map[name.toUpperCase()]=inf;
  }
  return map;
}
function _getProdukModelByKodeQ_(kode) {
  var key = _normalizeEditableCodeKey_(kode);
  if (!key) return '';
  var ss = SpreadsheetApp.openById(SS_ID), sh = ss.getSheetByName('PRODUK');
  if (!sh || sh.getLastRow() < 2) return '';
  var d = sh.getDataRange().getValues();
  // PRODUK!Q = KODE PRODUK untuk pencocokan, PRODUK!H = nilai yang ditulis ke DAMEN!Z.
  for (var i = 1; i < d.length; i++) {
    if (_normalizeEditableCodeKey_(d[i][16]) === key) return d[i][7] == null ? '' : d[i][7];
  }
  return '';
}
function _getProdukAAByKodeQ_(kode) {
  var key = _normalizeEditableCodeKey_(kode);
  if (!key) return '';
  var ss = SpreadsheetApp.openById(SS_ID), sh = ss.getSheetByName('PRODUK');
  if (!sh || sh.getLastRow() < 2) return '';
  var d = sh.getDataRange().getValues();
  // PRODUK!Q = KODE PRODUK untuk pencocokan, PRODUK!I = nilai yang ditulis ke DAMEN!AA.
  for (var i = 1; i < d.length; i++) {
    if (_normalizeEditableCodeKey_(d[i][16]) === key) return d[i][8] == null ? '' : d[i][8];
  }
  return '';
}
function _getProdukKategoriCashbackByKodeQ_(kode) {
  var key = _normalizeEditableCodeKey_(kode);
  if (!key) return '';
  var ss = SpreadsheetApp.openById(SS_ID), sh = ss.getSheetByName('PRODUK');
  if (!sh || sh.getLastRow() < 2) return '';
  var d = sh.getDataRange().getValues();
  var h = d[0].map(function(x){ return String(x == null ? '' : x).trim().toUpperCase(); });
  var idxKodeQ = 16; // Q
  var idxKategoriCashback = h.indexOf('KATEGORI CASHBACK');
  if (idxKategoriCashback < 0) idxKategoriCashback = 18; // S
  // Cocokkan Kode Produk DAMEN!E dengan PRODUK!Q, lalu ambil PRODUK!S (KATEGORI CASHBACK).
  for (var i = 1; i < d.length; i++) {
    if (_normalizeEditableCodeKey_(d[i][idxKodeQ]) === key) {
      return d[i][idxKategoriCashback] == null ? '' : d[i][idxKategoriCashback];
    }
  }
  return '';
}
function _normalizeEditableCodeKey_(v) {
  return String(v == null ? '' : v).trim().toUpperCase().replace(/[^A-Z0-9]/g,'');
}
function _normalizeEditableNameKey_(v) {
  return String(v == null ? '' : v).trim().toUpperCase()
    .replace(/\s+-\s+\d+(?:[.,]\d+)?\s*(?:PCS|Pcs|DUS|Dus)$/i,'')
    .replace(/\([^)]*\)$/,'')
    .replace(/[^A-Z0-9]+/g,'');
}
function _resolveEditableProduct_(map, name, kode) {
  var raw = String(name == null ? '' : name).trim();
  var k = String(kode == null ? '' : kode).trim().toUpperCase().replace(/\s+/g,' ');
  if (k && map[k]) return map[k];
  var kc = _normalizeEditableCodeKey_(k);
  if (kc && map['#CODE#'+kc]) return map['#CODE#'+kc];
  var key = _normalizeEditableProductKey_(raw);
  if (key && map[key]) return map[key];
  var nk = _normalizeEditableNameKey_(raw);
  if (nk && map['#NAME#'+nk]) return map['#NAME#'+nk];
  var rawUpper = raw.toUpperCase();
  if (rawUpper && map[rawUpper]) return map[rawUpper];
  // Fallback terakhir: scan nilai master, berguna jika ada apostrophe/spasi/format
  // berbeda pada KODE PRODUK atau nama produk di sheet PRODUK.
  var keys = Object.keys(map);
  for (var i=0;i<keys.length;i++) {
    var inf = map[keys[i]];
    if (!inf) continue;
    if (kc && _normalizeEditableCodeKey_(inf.kode) === kc) return inf;
    if (nk && _normalizeEditableNameKey_(inf.nama) === nk) return inf;
  }
  return null;
}
function _resolveEditableProductDirect_(name, kode) {
  try {
    var ss=SpreadsheetApp.openById(SS_ID), sh=ss.getSheetByName('PRODUK');
    if(!sh || sh.getLastRow()<2) return null;
    var data=sh.getDataRange().getValues();
    var hd=data[0].map(function(x){return String(x==null?'':x).trim().toUpperCase();});
    var k=hd.indexOf('KODE PRODUK'), n=hd.indexOf('PRODUK'), l=hd.indexOf('LITER'), h=hd.indexOf('HARGA'), pk=hd.indexOf('PACK');
    if(pk<0) pk=hd.indexOf('ISI');
    var needCode=_normalizeEditableCodeKey_(kode), needName=_normalizeEditableNameKey_(name);
    for(var r=1;r<data.length;r++){
      var row=data[r], rc=k>=0?row[k]:'', rn=n>=0?row[n]:'';
      var codeMatch=needCode && _normalizeEditableCodeKey_(rc)===needCode;
      var nameMatch=needName && _normalizeEditableNameKey_(rn)===needName;
      // Jika kode tidak ditemukan pada header yang diharapkan, cari kode di seluruh baris.
      if(!codeMatch && needCode){
        for(var c=0;c<row.length;c++){ if(_normalizeEditableCodeKey_(row[c])===needCode){codeMatch=true;break;} }
      }
      if(codeMatch || nameMatch){
        return {nama:String(rn||name||'').trim(),kode:String(rc||kode||'').trim(),liter:l>=0?(Number(row[l])||1):1,harga:h>=0?(Number(row[h])||0):0,pack:pk>=0?(Number(row[pk])||1):1,colD:row[3]||'',colE:row[4]||'',colF:row[5]||'',colG:row[6]||'',colL:row[11]||'',colR:row[17]||'',colS:row[18]||'',colX:row[23]||''};
      }
    }
  } catch(e) {}
  return null;
}
function _getExistingDamenProductInfo_(dm, targetId, name, kode) {
  try {
    if (!dm || dm.getLastRow() < 2) return null;
    var dd=dm.getDataRange().getValues();
    var dh=dd[0].map(function(x){return String(x==null?'':x).trim().toUpperCase();});
    var ix=function(n,f){var q=dh.indexOf(n);return q>=0?q:f;};
    var iId=ix('ID INPUT',0), iKode=ix('KODE PRODUK',4), iProd=ix('PRODUK',5), iQty=ix('QTY',6), iPcs=ix('PCS',7), iDus=ix('DUS',8), iVol=ix('VOLUME',15), iVal=ix('NET VALUE',16), iUnit=ix('SATUAN',17);
    var wantCode=_normalizeEditableCodeKey_(kode), wantName=_normalizeEditableNameKey_(name);
    for(var r=1;r<dd.length;r++){
      if(String(dd[r][iId]||'').trim()!==String(targetId||'').trim()) continue;
      var rc=String(dd[r][iKode]||'').trim(), rn=String(dd[r][iProd]||'').trim();
      var codeMatch=wantCode && _normalizeEditableCodeKey_(rc)===wantCode;
      var nameMatch=wantName && _normalizeEditableNameKey_(rn)===wantName;
      if(!codeMatch && !nameMatch) continue;
      var qty=Number(dd[r][iQty])||0, pcs=Number(dd[r][iPcs])||0, vol=Number(dd[r][iVol])||0, val=Number(dd[r][iVal])||0;
      var unit=String(dd[r][iUnit]||'PCS').trim().toUpperCase();
      var pack=(unit==='DUS' && qty>0 && pcs>0)?(pcs/qty):1;
      var liter=(pcs>0 && vol>0)?(vol/pcs):1;
      var harga=(pcs>0 && val>0)?(val/pcs):0;
      var rowD=ix('VAL D',-1),rowE=ix('VAL E',-1),rowF=ix('VAL F',-1),rowG=ix('VAL G',-1),rowL=ix('VAL L',-1),rowR=ix('VAL R',-1),rowS=ix('VAL S',-1),rowX=ix('VAL X',-1);return {nama:rn||String(name||''),kode:rc||String(kode||''),liter:liter,harga:harga,pack:pack,colD:rowD>=0?dd[r][rowD]:'',colE:rowE>=0?dd[r][rowE]:'',colF:rowF>=0?dd[r][rowF]:'',colG:rowG>=0?dd[r][rowG]:'',colL:rowL>=0?dd[r][rowL]:'',colR:rowR>=0?dd[r][rowR]:'',colS:rowS>=0?dd[r][rowS]:'',colX:rowX>=0?dd[r][rowX]:''};
    }
  } catch(e) {}
  return null;
}
function getEditableOrderById(idInput,salesName,usernameInput){try{var ss=SpreadsheetApp.openById(SS_ID),sh=ss.getSheetByName('INPUT');if(!sh||sh.getLastRow()<2)return{success:false,message:'Order tidak ditemukan.'};var d=sh.getDataRange().getValues(),hd=d[0].map(function(x){return String(x).trim().toUpperCase();}),idx=function(n,f){var x=hd.indexOf(n);return x>=0?x:f;},cId=idx('ID INPUT',0),cOutlet=idx('OUTLET',1),cDate=idx('DATE',2),cDsr=idx('DSR',4),cPay=idx('PAYMENT',5),cDisc=idx('DISC',6),cCust=idx('CUST.CODE',7),cChan=idx('CHANNEL',8),cNote=idx('NOTE ORDER',9),cStatus=idx('STATUS APPROVE',12),cCb=idx('CASHBACK',13),target=String(idInput||'').trim(),admin=String(usernameInput||'').trim().toLowerCase()==='admin',salesKey=String(salesName||'').trim().toUpperCase(),found=null;
for(var r=1;r<d.length;r++)if(String(d[r][cId]||'').trim()===target){found=d[r];break;}if(!found)return{success:false,message:'Order dengan ID INPUT '+target+' tidak ditemukan.'};var owner=String(found[cDsr]||'').trim();if(!admin&&owner.toUpperCase()!==salesKey)return{success:false,message:'Anda hanya dapat mengedit order milik Anda sendiri.'};var items=[],dm=ss.getSheetByName('DAMEN'),pm=_getEditableProductMap_();if(dm&&dm.getLastRow()>=2){var dd=dm.getDataRange().getValues(),dh=dd[0].map(function(x){return String(x).trim().toUpperCase();}),di=function(n,f){var x=dh.indexOf(n);return x>=0?x:f;},iId=di('ID INPUT',0),iKode=di('KODE PRODUK',4),iProd=di('PRODUK',5),iQty=di('QTY',6),iPcs=di('PCS',7),iUnit=di('SATUAN',17),iHarga=di('HARGA',-1);for(var j=1;j<dd.length;j++)if(String(dd[j][iId]||'').trim()===target){var nm=String(dd[j][iProd]||'').trim(),inf=_resolveEditableProduct_(pm,nm,dd[j][iKode])||{},qty=Number(dd[j][iQty])||0,unit=String(dd[j][iUnit]||'PCS').trim(),pack=Number(inf.pack)||1,pcs=Number(dd[j][iPcs])||((unit.toUpperCase()==='DUS')?qty*pack:qty),harga=iHarga>=0?(Number(dd[j][iHarga])||0):(Number(inf.harga)||0);var cleanNm=String(nm||'').replace(/\s+-\s+\d+(?:[.,]\d+)?\s*(?:PCS|Pcs|DUS|Dus)$/i,'').trim();items.push({kode:String(dd[j][iKode]||inf.kode||''),produk:cleanNm,qty:qty,satuan:unit,pack:pack,harga:harga,liter:Number(inf.liter)||1});}}
var dateStr='';if(found[cDate] instanceof Date)dateStr=Utilities.formatDate(found[cDate],Session.getScriptTimeZone(),'yyyy-MM-dd');else{try{var qd=new Date(String(found[cDate]||''));if(!isNaN(qd.getTime()))dateStr=Utilities.formatDate(qd,Session.getScriptTimeZone(),'yyyy-MM-dd');}catch(e){}}
return{success:true,order:{idInput:target,outlet:found[cOutlet]||'',tanggal:dateStr,payment:found[cPay]||'',disc:Number(found[cDisc])||0,custCode:found[cCust]||'',channel:found[cChan]||'',note:found[cNote]||'',cashback:found[cCb]||'',status:found[cStatus]||'PENDING',sales:owner,products:items}};}catch(e){return{success:false,message:'Gagal memuat order: '+e.message};}}
function updateOrderById(idInput,payload,salesName,usernameInput){try{payload=payload||{};var ss=SpreadsheetApp.openById(SS_ID),sh=ss.getSheetByName('INPUT'),dm=ensureDamenStructure_();if(!sh||!dm)return{success:false,message:'Sheet INPUT atau DAMEN tidak ditemukan.'};var d=sh.getDataRange().getValues(),hd=d[0].map(function(x){return String(x).trim().toUpperCase();}),idx=function(n,f){var x=hd.indexOf(n);return x>=0?x:f;},cId=idx('ID INPUT',0),cDsr=idx('DSR',4),target=String(idInput||'').trim(),admin=String(usernameInput||'').trim().toLowerCase()==='admin',salesKey=String(salesName||'').trim().toUpperCase(),rn=-1;for(var r=1;r<d.length;r++)if(String(d[r][cId]||'').trim()===target){rn=r+1;break;}if(rn<0)return{success:false,message:'Order tidak ditemukan.'};var old=d[rn-1],owner=String(old[cDsr]||'').trim();if(!admin&&owner.toUpperCase()!==salesKey)return{success:false,message:'Anda tidak berhak mengedit order ini.'};var items=Array.isArray(payload.items)?payload.items:[];if(!items.length)return{success:false,message:'Minimal harus ada 1 produk.'};var pm=_getEditableProductMap_(),prep=[],tv=0,tn=0;for(var i=0;i<items.length;i++){var nm=String(items[i].produk||'').trim(),code=String(items[i].kode||'').trim(),originalCode=String(items[i].originalKode||'').trim(),originalName=String(items[i].originalProduk||'').trim();/* Jika produk tidak berubah, pertahankan data lama untuk harga/Q. Jika Kode Produk berubah, WAJIB ambil master berdasarkan Kode Produk baru. */var sameOriginalCode=!!originalCode&&_normalizeEditableCodeKey_(originalCode)===_normalizeEditableCodeKey_(code);var inf=sameOriginalCode?_getExistingDamenProductInfo_(dm,target,nm,code):null;if(!inf && !originalCode && originalName && _normalizeEditableNameKey_(originalName)===_normalizeEditableNameKey_(nm)) inf=_getExistingDamenProductInfo_(dm,target,nm,code);if(!inf) inf=_resolveEditableProduct_(pm,nm,code);if(!inf) inf=_resolveEditableProductDirect_(nm,code);/* Produk baru yang dipilih dari frontend boleh membawa metadata master sendiri. Ini mencegah error 'Produk tidak valid' bila key nama/kode berbeda format. */if(!inf && nm){inf={nama:nm,kode:code||String(items[i].kode||''),liter:Number(items[i].liter)||1,harga:Number(items[i].harga)||0,pack:Number(items[i].pack)||1,colD:'',colE:'',colF:'',colG:'',colL:'',colR:'',colS:'',colX:''};}if(!inf)return{success:false,message:'Produk tidak valid: '+nm+(code?' (Kode: '+code+')':'')};var qty=Number(items[i].qty)||0,unit=String(items[i].satuan||'PCS').trim().toUpperCase();if(qty<=0)return{success:false,message:'Qty produk harus lebih dari 0.'};var pcs=unit==='DUS'?qty*(Number(inf.pack)||1):qty;var unitHarga=Number(items[i].harga)||Number(inf.harga)||0;/* Jangan biarkan Q menjadi 0 saat produk lama diedit. Fallback memakai Q lama / PCS dari DAMEN. */if(unitHarga<=0){var oldInfo=_getExistingDamenProductInfo_(dm,target,nm,code);if(oldInfo&&Number(oldInfo.harga)>0)unitHarga=Number(oldInfo.harga)||0;}var vol=pcs*(Number(inf.liter)||1),grossVal=pcs*unitHarga;
var editDiscRaw=payload.disc==null?'':String(payload.disc).trim().replace(',','.');
var editDiscValue=0;if(editDiscRaw!==''){var editDiscHasPercent=editDiscRaw.indexOf('%')!==-1;var editDiscNumber=parseFloat(editDiscRaw.replace('%','').trim());if(!isNaN(editDiscNumber))editDiscValue=editDiscHasPercent||Math.abs(editDiscNumber)>1?editDiscNumber/100:editDiscNumber;}
var netVal=grossVal*(1-editDiscValue);if(netVal<0)netVal=0;tv+=vol;tn+=netVal;prep.push({produk:nm,qty:qty,satuan:unit,finalPcs:pcs,finalDus:unit==='DUS'?qty:'',calculatedVolume:vol,grossValue:grossVal,netValue:netVal,hargaSatuan:unitHarga,resolvedKode:inf.kode||'',originalKode:String(items[i].originalKode||'').trim(),originalProduk:String(items[i].originalProduk||'').trim(),valD:inf.colD,valE:inf.colE,valF:inf.colF,valG:inf.colG,valL:inf.colL,valR:inf.colR,valS:inf.colS,valX:inf.colX});}
var dt=payload.tanggal?new Date(String(payload.tanggal)+'T00:00:00'):new Date();if(isNaN(dt.getTime()))dt=new Date();var tz=Session.getScriptTimeZone(),fd=Utilities.formatDate(dt,tz,'d MMMM yyyy'),mn=Utilities.formatDate(dt,tz,'MMMM'),yr=dt.getFullYear(),qt='QUARTAL '+(Math.floor(dt.getMonth()/3)+1)+' '+yr,dr=String(payload.disc==null?'':payload.disc).trim().replace(',','.'),disc=0,dn=parseFloat(dr.replace('%',''));if(dr&&!isNaN(dn))disc=(dr.indexOf('%')>=0||Math.abs(dn)>1)?dn/100:dn;for(var pi=0;pi<prep.length;pi++){prep[pi].netValue=Number(prep[pi].grossValue||0)*(1-disc);}tn=prep.reduce(function(a,it){return a+Number(it.netValue||0);},0);var status=old[idx('STATUS APPROVE',12)]||'PENDING',set=function(n,f,v){var x=idx(n,f);if(x>=0)old[x]=v;};set('OUTLET',1,payload.outlet||'');set('DATE',2,fd);set('MONTH',3,mn+'-'+String(yr).slice(-2));set('PAYMENT',5,payload.payment||'');set('DISC',6,disc);set('CUST.CODE',7,payload.custCode||'');set('CHANNEL',8,payload.channel||'');set('NOTE ORDER',9,payload.note||'');set('VOLUME',10,tv);set('VALUE',11,tn);set('STATUS APPROVE',12,status);set('CASHBACK',13,getCanonicalCashback_(payload.cashback));sh.getRange(rn,1,1,old.length).setValues([old]);
var dd=dm.getDataRange().getValues(),dh=dd[0].map(function(x){return String(x).trim().toUpperCase();}),ii=dh.indexOf('ID INPUT');if(ii<0)ii=0;var iKodeOld=dh.indexOf('KODE PRODUK'),iProdOld=dh.indexOf('PRODUK');var oldRows=[];for(var oz=1;oz<dd.length;oz++)if(String(dd[oz][ii]||'').trim()===target)oldRows.push(dd[oz].slice());var usedOld={};for(var z=dd.length-1;z>=1;z--)if(String(dd[z][ii]||'').trim()===target)dm.deleteRow(z+1);var start=dm.getLastRow()+1,rows=[];for(var j=0;j<prep.length;j++){var x=prep[j],originalCode=String(x.originalKode||'').trim(),originalName=String(x.originalProduk||'').trim(),oldRow=null,oldIndex=-1;/* Cari baris lama berdasarkan produk asal, bukan posisi array. Ini aman saat produk ditambah/dihapus/diurutkan ulang. */for(var oi=0;oi<oldRows.length;oi++){if(usedOld[oi])continue;var oc=iKodeOld>=0?String(oldRows[oi][iKodeOld]||'').trim():'';var on=iProdOld>=0?String(oldRows[oi][iProdOld]||'').trim():'';var sameCode=originalCode&&_normalizeEditableCodeKey_(oc)===_normalizeEditableCodeKey_(originalCode);var sameName=originalName&&_normalizeEditableNameKey_(on)===_normalizeEditableNameKey_(originalName);if(sameCode||(!originalCode&&sameName)){oldRow=oldRows[oi];oldIndex=oi;usedOld[oi]=true;break;}}var row=oldRow ? oldRow.slice() : new Array(Math.max(40,dh.length)).fill('');while(row.length<Math.max(40,dh.length))row.push('');var unchanged=!!oldRow && _normalizeEditableCodeKey_(String(x.resolvedKode||''))===_normalizeEditableCodeKey_(originalCode||String(oldRow[iKodeOld>=0?iKodeOld:4]||''));row[0]=target;row[1]="DAM-EDIT-"+Utilities.getUuid().replace(/-/g,"").substring(0,12);row[2]=payload.outlet||'';row[3]=payload.custCode||'';row[4]=x.resolvedKode;row[5]=x.produk;row[6]=x.qty;row[7]=x.finalPcs;row[8]=x.finalDus;row[9]=fd;row[10]=mn+'-'+String(yr).slice(-2);row[11]=owner;row[12]=x.produk+' - '+x.qty+' '+x.satuan;row[13]=payload.payment||'';row[14]=disc;row[15]=x.calculatedVolume;row[16]=Number(x.netValue)||0;row[17]=x.satuan;/* V:AA = master PRODUK D,E,F,G,L,R. Khusus Z, cocokkan DAMEN!E (Kode Produk) dengan PRODUK!Q (Kode Produk), lalu ambil PRODUK!H. Produk yang sama tetap mempertahankan kolom lain; Z selalu mengikuti lookup Kode Produk. */if(!unchanged){row[21]=x.valD||'';row[22]=x.valE||'';row[23]=x.valF||'';row[24]=x.valG||'';}var modelProdukByKodeQ=_getProdukModelByKodeQ_(x.resolvedKode);if(modelProdukByKodeQ!=='')row[25]=modelProdukByKodeQ;else if(unchanged)row[25]=oldRow[25]||'';var aaProdukByKodeQ=_getProdukAAByKodeQ_(x.resolvedKode);if(aaProdukByKodeQ!=='')row[26]=aaProdukByKodeQ;else if(unchanged)row[26]=oldRow[26]||'';/* Kolom turunan/metadata lain tetap mengikuti perilaku sebelumnya. */if(!oldRow){row[30]=x.valL||'';row[33]=x.valS||'';row[39]=x.valR||'';}else{if(row[30]==='')row[30]=oldRow[30]||'';if(row[33]==='')row[33]=oldRow[33]||'';if(row[39]==='')row[39]=oldRow[39]||'';}row[28]=mn;row[29]=qt;row[33]=_getProdukKategoriCashbackByKodeQ_(x.resolvedKode);var akCashback = getCanonicalCashback_(payload.cashback);
      var akAh = String(row[33] || '').trim();
      var akAd = String(row[29] || '').trim();
      // AK harus persis seperti saat order baru: CASHBACK - AH - AD.
      // AL mencocokkan AK secara exact-match dengan key CASHBACK kolom L.
      var akValue = '';
      if (akCashback) akValue = akCashback;
      if (akAh) akValue += (akValue ? ' - ' : '') + akAh;
      if (akAd) akValue += (akValue ? ' - ' : '') + akAd;
      row[36] = akValue;rows.push(row);}dm.getRange(start,1,rows.length,rows[0].length).setValues(rows);dm.getRange(start,15,rows.length,1).setNumberFormat('0%');refreshDerivedForNewOrder_(dm,sh,target,prep,String(payload.cashback||''),fd);clearAppDataCaches();return{success:true,message:'Order '+target+' berhasil diperbarui.'};}catch(e){return{success:false,message:'Gagal memperbarui order: '+e.message};}}

function getRiwayatOrder_UNCACHED(salesName, usernameInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("INPUT");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colDsr = headers.indexOf("DSR") !== -1 ? headers.indexOf("DSR") : 4;
    var colId = headers.indexOf("ID INPUT") !== -1 ? headers.indexOf("ID INPUT") : 0;
    var colOutlet = headers.indexOf("OUTLET") !== -1 ? headers.indexOf("OUTLET") : 1;
    var colDate = headers.indexOf("DATE") !== -1 ? headers.indexOf("DATE") : 2;
    var colVol = headers.indexOf("VOLUME") !== -1 ? headers.indexOf("VOLUME") : 10;
    var colVal = headers.indexOf("VALUE") !== -1 ? headers.indexOf("VALUE") : 11;
    var colStatus = headers.indexOf("STATUS APPROVE") !== -1 ? headers.indexOf("STATUS APPROVE") : 12;
    var colNote = headers.indexOf("NOTE ORDER") !== -1 ? headers.indexOf("NOTE ORDER") : (headers.indexOf("NOTE") !== -1 ? headers.indexOf("NOTE") : 9);

    var isAdmin = (usernameInput && String(usernameInput).trim().toLowerCase() === 'admin');
    var salesKey = String(salesName || '').trim().toUpperCase();

    // Siapkan rincian produk DAMEN berdasarkan ID INPUT.
    // DAMEN A = ID INPUT, DAMEN M = Rincian Produk.
    var productsById = {};
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (sheetDamen && sheetDamen.getLastRow() >= 2) {
      var lastDamen = sheetDamen.getLastRow();
      var damenData = sheetDamen.getRange(2, 1, lastDamen - 1, 13).getValues();
      for (var d = 0; d < damenData.length; d++) {
        var damenId = String(damenData[d][0] == null ? "" : damenData[d][0]).trim();
        if (!damenId) continue;
        var produk = String(damenData[d][12] == null ? "" : damenData[d][12]).trim();
        if (!productsById[damenId]) productsById[damenId] = [];
        if (produk) productsById[damenId].push(produk);
      }
    }

    var riwayat = [];
    for (var j = 1; j < data.length; j++) {
      var rowSales = String(data[j][colDsr] == null ? "" : data[j][colDsr]).trim();
      if (!(isAdmin || rowSales.toUpperCase() === salesKey)) continue;

      var idValue = String(data[j][colId] == null ? "" : data[j][colId]).trim();
      var products = productsById[idValue] ? productsById[idValue].slice() : [];
      var outletForWA = String(data[j][colOutlet] == null ? "" : data[j][colOutlet]).trim();
      var noteForWA = String(data[j][colNote] == null ? "" : data[j][colNote]).trim();
      var pesanWA = "*" + outletForWA + "*" + "\n" + (products.length ? products.map(function(p){ return "- " + String(p).trim(); }).join("\n") : "-");
      if (noteForWA) pesanWA += "\n\n*Note :* " + noteForWA;

      riwayat.push({
        id: idValue || "-",
        outlet: data[j][colOutlet] || "-",
        tanggal: data[j][colDate] ? Utilities.formatDate(new Date(data[j][colDate]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "-",
        volume: Number(data[j][colVol] || 0).toFixed(1),
        value: Number(data[j][colVal] || 0),
        status: data[j][colStatus] || "PENDING",
        dsr: rowSales || "-",
        products: products,
        whatsappUrl: "https://api.whatsapp.com/send?text=" + encodeURIComponent(pesanWA)
      });
    }

    return riwayat.reverse();
  } catch(e) {
    Logger.log("Error getRiwayatOrder: " + e.message);
    return [];
  }
}


/**
 * Pencapaian Sales bulanan.
 * TARGET PER SALES: A=SALES, B=KATEGORI, C=MONTH, D=TARGET, F=PERIODE.
 * Actual dihitung dari DAMEN yang sudah APPROVE pada bulan berjalan.
 */
function getPencapaianSales_UNCACHED(salesName, usernameInput, selectedPeriod) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var targetSheet = ss.getSheetByName("TARGET PER SALES");
    if (!targetSheet) return {success:false, message:'Sheet "TARGET PER SALES" tidak ditemukan.', data:[], periods:[]};

    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var salesKey = String(salesName || "").trim().toUpperCase();
    var isAdmin = String(usernameInput || "").trim().toLowerCase() === "admin";

    var targetRange = targetSheet.getDataRange();
    var targetData = targetRange.getValues();
    var targetDisplay = targetRange.getDisplayValues();
    if (targetData.length <= 1) return {success:true, periode:'-', selectedPeriod:'', data:[], periods:[]};

    var monthNames = {
      JANUARI:0, FEBRUARI:1, MARET:2, APRIL:3, MEI:4, JUNI:5, JULI:6, AGUSTUS:7,
      SEPTEMBER:8, OKTOBER:9, NOVEMBER:10, DESEMBER:11,
      JANUARY:0, FEBRUARY:1, MARCH:2, APRIL:3, MAY:4, JUNE:5, JULY:6,
      AUGUST:7, SEPTEMBER:8, OCTOBER:9, NOVEMBER:10, DECEMBER:11
    };

    function norm(v) {
      return String(v == null ? "" : v).trim().toUpperCase().replace(/\s+/g, " ");
    }
    function parseNum(v) {
      if (typeof v === 'number') return v || 0;
      var s = String(v == null ? "" : v).trim();
      if (!s) return 0;
      if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) s = s.replace(/\./g,'').replace(',','.');
      else if (s.indexOf(',') !== -1) s = s.replace(',', '.');
      else s = s.replace(/\.(?=\d{3}(?:\D|$))/g, '');
      var n = parseFloat(s.replace(/[^0-9.-]/g,''));
      return isNaN(n) ? 0 : n;
    }
    function category(v) {
      var c = norm(v);
      if (c.indexOf("VALUE") !== -1) return "VALUE";
      if (c.indexOf("VOLUME") !== -1) return "VOLUME";
      if (c.indexOf("ACTIVE OUTLET") !== -1 || c.indexOf("ACTIV OUTLET") !== -1 || c === "OUTLET AKTIF" || c === "ACTIVE") return "ACTIV OUTLET";
      if (c.indexOf("PRODUK FOKUS") !== -1 || c.indexOf("PRODUCT FOCUS") !== -1 || c.indexOf("FOCUS PRODUCT") !== -1) return "PRODUK FOKUS";
      if (c.indexOf("POWER BRAND") !== -1 || c === "PB") return "POWER BRAND";
      return c;
    }
    function periodKeyFromDate(d) {
      return Utilities.formatDate(d, tz, "yyyy-MM");
    }
    function periodLabelFromKey(key) {
      if (!key) return '-';
      var parts = key.split('-');
      var y = Number(parts[0]), m = Number(parts[1]) - 1;
      return Utilities.formatDate(new Date(y, m, 1), tz, "MMMM-yy");
    }
    function parsePeriodValue(v, display) {
      if (v instanceof Date && !isNaN(v.getTime())) return periodKeyFromDate(v);
      var raw = String(display || v || '').trim();
      if (!raw) return '';
      var n = norm(raw).replace(/\//g,'-').replace(/\./g,'-');
      var m = n.match(/^(\d{4})-(\d{1,2})$/);
      if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
      var m2 = n.match(/^([A-ZÀ-Ÿ]+)[ -](\d{2,4})$/);
      if (m2 && monthNames[m2[1]] !== undefined) {
        var yy = Number(m2[2]); if (yy < 100) yy += 2000;
        return yy + '-' + ('0' + (monthNames[m2[1]] + 1)).slice(-2);
      }
      return '';
    }
    function monthIndexFromMonthLabel(v) {
      var m = norm(v).match(/^MONTH\s*(\d+)$/);
      return m ? Number(m[1]) : 0;
    }
    function pct(a,t) { return t > 0 ? (a/t)*100 : 0; }

    // Periode pilihan ditampilkan sebagai KUARTAL: Q1 2026, Q2 2026, dst.
    // Sumbernya tetap kolom F (PERIODE) pada TARGET PER SALES.
    var availableQuarters = {};
    for (var r0=1; r0<targetData.length; r0++) {
      var pk = parsePeriodValue(targetData[r0][5], targetDisplay[r0][5]);
      if (!pk) continue;
      var pparts = pk.split('-');
      var py = Number(pparts[0]);
      var pm = Number(pparts[1]) - 1;
      var pq = Math.floor(pm / 3) + 1;
      availableQuarters[py + '-Q' + pq] = true;
    }

    var currentMonthKey = periodKeyFromDate(now);
    var currentParts = currentMonthKey.split('-');
    var currentYear = Number(currentParts[0]);
    var currentMonth = Number(currentParts[1]) - 1;
    var currentQuarterKey = currentYear + '-Q' + (Math.floor(currentMonth / 3) + 1);

    if (!Object.keys(availableQuarters).length) availableQuarters[currentQuarterKey] = true;
    var periods = Object.keys(availableQuarters).sort().reverse().map(function(k){
      return {value:k, label:k.replace('-Q', ' Q').replace(/^(\d+) /, '$1 ')};
    });

    var chosenKey = String(selectedPeriod || '').trim();
    if (!availableQuarters[chosenKey]) chosenKey = availableQuarters[currentQuarterKey] ? currentQuarterKey : periods[0].value;

    var qmatch = chosenKey.match(/^(\d{4})-Q([1-4])$/);
    var chosenYear = qmatch ? Number(qmatch[1]) : currentYear;
    var quarterNo = qmatch ? Number(qmatch[2]) : (Math.floor(currentMonth / 3) + 1);
    var quarterStartMonth = (quarterNo - 1) * 3;
    var quarterStartKey = chosenYear + '-' + ('0' + (quarterStartMonth + 1)).slice(-2);
    var quarterEndKey = chosenYear + '-' + ('0' + (quarterStartMonth + 3)).slice(-2);
    var quarterLabel = 'Q' + quarterNo + ' ' + chosenYear;

    // Tiga bulan dalam kuartal: MONTH 1, MONTH 2, MONTH 3.
    var monthKeys = [];
    for (var mi=0; mi<3; mi++) {
      var mk = chosenYear + '-' + ('0' + (quarterStartMonth + mi + 1)).slice(-2);
      monthKeys.push(mk);
    }

    // Sales yang memiliki target pada periode yang dipilih.
    var selectedPeriodSales = {};
    for (var rp=1; rp<targetData.length; rp++) {
      var salesP = norm(targetData[rp][0]);
      if (!salesP) continue;
      if (salesKey && salesP !== salesKey) continue;
      var pkP = parsePeriodValue(targetData[rp][5], targetDisplay[rp][5]);
      if (pkP) {
        var pp = pkP.split('-');
        var py2 = Number(pp[0]), pm2 = Number(pp[1]) - 1;
        var pq2 = Math.floor(pm2 / 3) + 1;
        if ((py2 + '-Q' + pq2) === chosenKey) selectedPeriodSales[salesP] = true;
      }
    }

    // Target per Sales, per kategori, per bulan.
    var targets = {};
    var salesOrder = [];
    for (var r=1; r<targetData.length; r++) {
      var row=targetData[r], disp=targetDisplay[r];
      var sales=norm(row[0]), cat=category(row[1]);
      if (!sales || !cat || !selectedPeriodSales[sales]) continue;

      var rowKey=parsePeriodValue(row[5], disp[5]);
      if (!rowKey) continue;
      var mi2=monthKeys.indexOf(rowKey);
      if (mi2 < 0) continue;

      if (!targets[sales]) {
        targets[sales] = {
          sales:sales,
          value:[0,0,0], volume:[0,0,0], active:[0,0,0], product:[0,0,0], powerBrand:[0,0,0]
        };
        salesOrder.push(sales);
      }
      var target=parseNum(row[3]);
      if (cat === 'VALUE') targets[sales].value[mi2] += target;
      else if (cat === 'VOLUME') targets[sales].volume[mi2] += target;
      else if (cat === 'ACTIV OUTLET') targets[sales].active[mi2] += target;
      else if (cat === 'PRODUK FOKUS') targets[sales].product[mi2] += target;
      else if (cat === 'POWER BRAND') targets[sales].powerBrand[mi2] += target;
    }

    if (!salesOrder.length) {
      return {
        success:true, periode:quarterLabel, selectedPeriod:chosenKey,
        volumeQuarter:quarterLabel, data:[], periods:periods
      };
    }

    // Actual per Sales per bulan.
    var actual={};
    salesOrder.forEach(function(s){
      actual[s] = {
        value:[0,0,0],
        volume:[0,0,0],
        powerBrand:[0,0,0],
        outlets:[{},{},{}],
        pbOutlets:[{},{},{}]
      };
    });

    var sheetDamen=ss.getSheetByName('DAMEN');
    if (sheetDamen && sheetDamen.getLastRow() >= 2) {
      var damenData=sheetDamen.getDataRange().getValues();
      var dh=damenData[0].map(norm);
      var cDsr=dh.indexOf('DSR') !== -1 ? dh.indexOf('DSR') : 11;
      var cOutlet=dh.indexOf('OUTLET') !== -1 ? dh.indexOf('OUTLET') : 2;
      var cDate=dh.indexOf('DATE') !== -1 ? dh.indexOf('DATE') : 9;
      var cVolume=dh.indexOf('VOLUME') !== -1 ? dh.indexOf('VOLUME') : 15;
      var cValue=dh.indexOf('VALUE') !== -1 ? dh.indexOf('VALUE') : 16;
      var cStatus=-1;
      for (var hh=0; hh<dh.length; hh++) {
        if (dh[hh]==='STATUS CETAK'||dh[hh]==='STATUS_CETAK'||dh[hh]==='CETAK'||dh[hh]==='STATUS APPROVE') {
          cStatus=hh; break;
        }
      }
      // PRODUK FOKUS menggunakan kolom AE / header SEGMENT PB di sheet DAMEN.
      // Jangan menggunakan kolom KATEGORI karena PB ditentukan dari SEGMENT PB.
      var cSegmentPB=dh.indexOf('SEGMENT PB') !== -1 ? dh.indexOf('SEGMENT PB') : 30; // AE

      for (var d=1; d<damenData.length; d++) {
        var dr=damenData[d];
        if (cStatus !== -1 && norm(dr[cStatus]) !== 'APPROVE') continue;
        var dsr=norm(dr[cDsr]);
        if (!actual[dsr]) continue;
        var dt=new Date(dr[cDate]);
        if (isNaN(dt.getTime())) continue;
        var dk=periodKeyFromDate(dt);
        var idx=monthKeys.indexOf(dk);
        if (idx < 0) continue;

        var outlet=String(dr[cOutlet] == null ? '' : dr[cOutlet]).trim();
        var vol=parseNum(dr[cVolume]);
        var val=parseNum(dr[cValue]);

        actual[dsr].value[idx] += val;
        actual[dsr].volume[idx] += vol;
        // POWER BRAND = total VOLUME hanya untuk baris dengan SEGMENT PB = PB pada kolom AE.
        if (norm(dr[cSegmentPB]) === 'PB') actual[dsr].powerBrand[idx] += vol;
        if (outlet) {
          var ok=outlet.toUpperCase();
          actual[dsr].outlets[idx][ok]=true;
          // PRODUK FOKUS = SEGMENT PB pada kolom AE, dengan akumulasi
          // volume per outlet per bulan. Outlet dihitung 1 jika total >= 9,6.
          if (norm(dr[cSegmentPB]) === 'PB') {
            actual[dsr].pbOutlets[idx][ok]=(actual[dsr].pbOutlets[idx][ok] || 0) + vol;
          }
        }
      }
    }

    var resultData=salesOrder.map(function(s){
      var t=targets[s], a=actual[s];
      var cats={};
      [
        ['value','Value','value'],
        ['volume','Volume','volume'],
        ['active','Active Outlet','active'],
        ['product','Produk Fokus','product'],
        ['powerBrand','Power Brand','powerBrand']
      ].forEach(function(def){
        var key=def[0], type=def[2], months=[];
        for (var j=0;j<3;j++) {
          var av;
          if (type === 'active') av=Object.keys(a.outlets[j]).length;
          else if (type === 'product') av=Object.keys(a.pbOutlets[j]).filter(function(k){return Number(a.pbOutlets[j][k]||0)>=9.6;}).length;
          else av=a[type][j];

          var tv = key === 'powerBrand' ? a.volume[j] : t[type][j];
          months.push({
            monthNo:j+1,
            period:periodLabelFromKey(monthKeys[j]),
            actual:av,
            target:tv,
            gap:Number(av || 0) - Number(tv || 0),
            percentage:pct(av,tv)
          });
        }
        var qActual=months.reduce(function(sum,x){return sum+Number(x.actual||0);},0);
        var qTarget=months.reduce(function(sum,x){return sum+Number(x.target||0);},0);
        cats[key]={
          months:months,
          quarter:{
            actual:qActual,
            target:qTarget,
            gap:qActual - qTarget,
            percentage:pct(qActual,qTarget)
          }
        };
      });

      return {
        sales:s,
        periode:quarterLabel,
        volumeQuarter:quarterLabel,
        categories:cats,
        // Field lama dipertahankan agar kompatibel dengan kode lain.
        targetValue:t.value[0] || 0,
        targetVolume:t.volume.reduce(function(x,y){return x+y;},0),
        targetActiveOutlet:t.active[0] || 0,
        targetProductFocus:t.product[0] || 0,
        targetPowerBrand:t.powerBrand[0] || 0,
        actualValue:a.value[0] || 0,
        actualVolume:a.volume.reduce(function(x,y){return x+y;},0),
        actualActiveOutlet:Object.keys(a.outlets[0] || {}).length,
        actualProductFocus:Object.keys(a.pbOutlets[0] || {}).filter(function(k){return Number(a.pbOutlets[0][k]||0)>=9.6;}).length,
        actualPowerBrand:a.powerBrand[0] || 0
      };
    });

    return {
      success:true,
      periode:quarterLabel,
      selectedPeriod:chosenKey,
      volumeQuarter:quarterLabel,
      quarterMonths:monthKeys.map(function(k,i){return {monthNo:i+1, key:k, label:periodLabelFromKey(k)};}),
      data:resultData,
      periods:periods
    };
  } catch(e) {
    Logger.log('Error getPencapaianSales: ' + e.message);
    return {success:false, message:'Gagal memuat pencapaian: ' + e.message, data:[], periods:[]};
  }
}

/**
 * Menentukan STATUS APPROVE pada sheet INPUT berdasarkan ID INPUT.
 * Hanya username ADMIN yang diizinkan mengubah status.
 * status yang valid: APPROVE / DECLINE.
 */
function updateStatusApprove(idInput, status, usernameInput) {
  try {
    var cleanUser = String(usernameInput || "").trim().toLowerCase();
    if (cleanUser !== "admin") {
      return { success: false, message: "Akses ditolak. Hanya Admin yang dapat menentukan status approve." };
    }

    var cleanId = String(idInput || "").trim();
    var cleanStatus = String(status || "").trim().toUpperCase();

    if (!cleanId) {
      return { success: false, message: "ID INPUT tidak valid." };
    }

    if (cleanStatus !== "APPROVE" && cleanStatus !== "DECLINE") {
      return { success: false, message: "Status hanya boleh APPROVE atau DECLINE." };
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("INPUT");
    if (!sheet) {
      return { success: false, message: "Sheet INPUT tidak ditemukan." };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: "Belum ada data pada sheet INPUT." };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colId = headers.indexOf("ID INPUT");
    var colStatus = headers.indexOf("STATUS APPROVE");
    if (colId < 0) colId = 0;
    if (colStatus < 0) colStatus = 12; // M

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colId] == null ? "" : data[i][colId]).trim() === cleanId) {
        // Simpan status pada INPUT (kolom M / STATUS APPROVE)
        sheet.getRange(i + 1, colStatus + 1).setValue(cleanStatus);

        // Sinkronkan status ke DAMEN kolom AB berdasarkan ID INPUT pada kolom A.
        // Semua baris DAMEN dengan ID INPUT yang sama akan di-update.
        var sheetDamen = ss.getSheetByName("DAMEN");
        var damenUpdated = 0;
        if (sheetDamen && sheetDamen.getLastRow() >= 2) {
          var lastDamenRow = sheetDamen.getLastRow();
          var damenIds = sheetDamen.getRange(2, 1, lastDamenRow - 1, 1).getValues();
          var statusValues = sheetDamen.getRange(2, 28, lastDamenRow - 1, 1).getValues(); // AB

          for (var d = 0; d < damenIds.length; d++) {
            if (String(damenIds[d][0] == null ? "" : damenIds[d][0]).trim() === cleanId) {
              statusValues[d][0] = cleanStatus;
              damenUpdated++;
            }
          }

          if (damenUpdated > 0) {
            sheetDamen.getRange(2, 28, statusValues.length, 1).setValues(statusValues);
          }
        }

        clearAppDataCaches();
        return {
          success: true,
          id: cleanId,
          status: cleanStatus,
          damenUpdated: damenUpdated,
          message: "Status " + cleanStatus + " berhasil disimpan untuk " + cleanId + ". INPUT dan " + damenUpdated + " baris DAMEN telah diperbarui."
        };
      }
    }

    return { success: false, message: "ID INPUT " + cleanId + " tidak ditemukan." };
  } catch (e) {
    return { success: false, message: "Gagal mengubah status: " + e.message };
  }
}

function getDetailOrderById(idInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var targetId = String(idInput == null ? "" : idInput).trim();
    if (!targetId) return { found: false, order: null, products: [] };

    // ============================================================
    // SUMBER UTAMA DETAIL ORDER: SHEET INPUT berdasarkan ID INPUT
    // A = ID INPUT, B = OUTLET, K = VOLUME, L = VALUE
    // ============================================================
    var sheetInput = ss.getSheetByName("INPUT");
    if (!sheetInput || sheetInput.getLastRow() < 2) {
      return { found: false, order: null, products: [] };
    }

    var inputData = sheetInput.getDataRange().getValues();
    var headers = inputData[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var colId = headers.indexOf("ID INPUT");
    var colOutlet = headers.indexOf("OUTLET");
    var colVolume = headers.indexOf("VOLUME");
    var colValue = headers.indexOf("VALUE");

    if (colId < 0) colId = 0;
    if (colOutlet < 0) colOutlet = 1;
    if (colVolume < 0) colVolume = 10;
    if (colValue < 0) colValue = 11;

    var order = null;
    for (var r = 1; r < inputData.length; r++) {
      var rowId = String(inputData[r][colId] == null ? "" : inputData[r][colId]).trim();
      if (rowId === targetId) {
        order = {
          idInput: rowId,
          outlet: String(inputData[r][colOutlet] == null ? "" : inputData[r][colOutlet]).trim() || "-",
          volume: Number(inputData[r][colVolume] || 0),
          value: Number(inputData[r][colValue] || 0)
        };
        break;
      }
    }

    // Jangan pernah mengambil order lain bila ID INPUT tidak cocok.
    if (!order) {
      return { found: false, order: null, products: [] };
    }

    // ============================================================
    // RINCIAN PRODUK: SHEET DAMEN berdasarkan ID INPUT kolom A
    // M = rincian produk yang ditampilkan
    // ============================================================
    var sheetDamen = ss.getSheetByName("DAMEN");
    var products = [];

    if (sheetDamen && sheetDamen.getLastRow() >= 2) {
      var damenData = sheetDamen.getRange(2, 1, sheetDamen.getLastRow() - 1, 13).getValues();
      for (var i = 0; i < damenData.length; i++) {
        var damenId = String(damenData[i][0] == null ? "" : damenData[i][0]).trim();
        if (damenId !== targetId) continue;

        var produk = String(damenData[i][12] == null ? "" : damenData[i][12]).trim();
        products.push({
          no: products.length + 1,
          produk: produk || "-"
        });
      }
    }

    return {
      found: true,
      order: order,
      products: products
    };
  } catch (e) {
    Logger.log("Error getDetailOrderById: " + e.message);
    throw new Error("Gagal mengambil detail berdasarkan ID INPUT: " + e.message);
  }
}

function _deleteRowsByIdsFast_(sheet, idSet) {
  if (!sheet || !idSet || !Object.keys(idSet).length) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // Baca hanya kolom A (ID INPUT), bukan seluruh sheet.
  // Ini jauh lebih ringan untuk sheet INPUT/DAMEN yang besar.
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rows = [];

  for (var i = 0; i < values.length; i++) {
    var id = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (id && idSet[id]) rows.push(i + 2);
  }

  if (!rows.length) return 0;

  // Kelompokkan baris berurutan agar deleteRows() dipanggil sesedikit mungkin.
  var groups = [];
  var start = rows[0];
  var prev = rows[0];

  for (var r = 1; r < rows.length; r++) {
    if (rows[r] === prev + 1) {
      prev = rows[r];
    } else {
      groups.push({ start: start, count: prev - start + 1 });
      start = prev = rows[r];
    }
  }
  groups.push({ start: start, count: prev - start + 1 });

  // Hapus dari bawah ke atas supaya nomor baris tidak bergeser.
  for (var g = groups.length - 1; g >= 0; g--) {
    sheet.deleteRows(groups[g].start, groups[g].count);
  }

  return rows.length;
}

function hapusOrderByIds(ids) {
  try {
    var list = Array.isArray(ids) ? ids : [ids];
    var idSet = {};

    list.forEach(function(id) {
      var key = String(id == null ? '' : id).trim();
      if (key) idSet[key] = true;
    });

    var uniqueIds = Object.keys(idSet);
    if (!uniqueIds.length) {
      return { success: false, message: 'Tidak ada transaksi yang dipilih.' };
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetInput = ss.getSheetByName('INPUT');
    var sheetDamen = ss.getSheetByName('DAMEN');

    var deletedInput = _deleteRowsByIdsFast_(sheetInput, idSet);
    var deletedDamen = _deleteRowsByIdsFast_(sheetDamen, idSet);

    clearAppDataCaches();

    return {
      success: true,
      message: uniqueIds.length + ' transaksi berhasil dihapus.',
      deletedIds: uniqueIds,
      deletedInputRows: deletedInput,
      deletedDamenRows: deletedDamen
    };
  } catch (e) {
    return { success: false, message: 'Gagal menghapus transaksi terpilih: ' + e.message };
  }
}

function hapusOrderById(idInput) {
  return hapusOrderByIds([idInput]);
}

function getWhatsAppUrlById(idInput) {
  try {
    let ss = SpreadsheetApp.openById(SS_ID);
    let sheetDamen = ss.getSheetByName("DAMEN");
    let sheetInput = ss.getSheetByName("INPUT");

    if (!sheetDamen) {
      return { success: false, whatsappUrl: "" };
    }

    let dataDamen = sheetDamen.getDataRange().getValues();
    let listItems = [];
    let outletName = "";
    let noteOrder = "";

    for (let i = 1; i < dataDamen.length; i++) {
      if (String(dataDamen[i][0]).trim() === String(idInput).trim()) {
        if (!outletName) outletName = dataDamen[i][2];

        // Produk untuk pesan WhatsApp diambil langsung dari DAMEN kolom M.
        // Kolom M sudah berisi format: Nama Produk - Qty Satuan.
        let produkVal = dataDamen[i][12];

        if (produkVal) {
          listItems.push("- " + String(produkVal).trim());
        }
      }
    }

    // NOTE ORDER diambil dari INPUT kolom J.
    if (sheetInput) {
      let dataInput = sheetInput.getDataRange().getValues();
      for (let i = 1; i < dataInput.length; i++) {
        if (String(dataInput[i][0]).trim() === String(idInput).trim()) {
          noteOrder = String(dataInput[i][9] == null ? "" : dataInput[i][9]).trim();
          break;
        }
      }
    }

    let pesan = "*" + (outletName || "") + "*" + "\n" +
                (listItems.length > 0 ? listItems.join("\n") : "-");

    if (noteOrder) {
      pesan += "\n\n*Note :* " + noteOrder;
    }

    let encodedPesan = encodeURIComponent(pesan);
    let url = "https://api.whatsapp.com/send?text=" + encodedPesan;

    return { success: true, whatsappUrl: url };
  } catch (error) {
    return { success: false, whatsappUrl: "" };
  }
}

/**
 * Fungsi untuk menghitung dan mengisi Kolom AL pada sheet DAMEN:
 * Cocokkan Kolom AK sheet DAMEN dengan Kolom H sheet CASHBACK,
 * lalu kalikan dengan Kolom F sheet CASHBACK dan Kolom P sheet DAMEN.
 */
function rebuildDamenDerivedColumns() {
  try {
    hitungKolomALDamen();
    hitungKolomAMDamen();
    hitungKolomLInputDariAMDamen();
    hitungKolomAODamen();
    return { success:true, message:'Kolom AL, AM, AO DAMEN dan INPUT kolom L berhasil dihitung ulang.' };
  } catch(e) {
    Logger.log('rebuildDamenDerivedColumns: '+e.message);
    return { success:false, message:e.message };
  }
}

function hitungKolomALDamen() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetDamen = ss.getSheetByName("DAMEN");
    var sheetCashback = ss.getSheetByName("CASHBACK");
    if (!sheetDamen || !sheetCashback) return;
    
    var lastRowDamen = sheetDamen.getLastRow();
    if (lastRowDamen < 2) return;
    
    var dataCashback = sheetCashback.getDataRange().getValues();
    var cashbackMap = {};
    
    for (var c = 1; c < dataCashback.length; c++) {
      // Kolom H di CASHBACK (indeks array ke-7) sebagai kunci pencocokan dengan Kolom AK DAMEN
      var keyLookup = String(dataCashback[c][11]).trim(); 
      // Nilai cashback diambil dari Kolom F (indeks array ke-5)
      var valCashback = dataCashback[c][5]; 
      
      if (keyLookup !== "") {
        cashbackMap[keyLookup] = parseFloat(valCashback) || 0;
      }
    }
    
    // Perbaikan Indeks Kolom: 
    // Kolom AK = Kolom ke-37 (getRange parameter kolom ke-37)
    // Kolom P  = Kolom ke-16 (getRange parameter kolom ke-16)
    var rangeAK = sheetDamen.getRange(2, 37, lastRowDamen - 1, 1).getValues();
    var rangeP  = sheetDamen.getRange(2, 16, lastRowDamen - 1, 1).getValues();
    var hasilAL = [];
    
    for (var i = 0; i < rangeAK.length; i++) {
      var valAK = String(rangeAK[i][0]).trim();
      var valP  = parseFloat(rangeP[i][0]) || 0;
      
      var matchedVal = cashbackMap[valAK] !== undefined ? cashbackMap[valAK] : 0;
      
      // Hasil = Nilai Cashback (Kolom F Cashback) * Kolom P DAMEN
      hasilAL.push([matchedVal * valP]);
    }
    
    // Masukkan ke Kolom AL (Kolom ke-38)
    sheetDamen.getRange(2, 38, hasilAL.length, 1).setValues(hasilAL);
  } catch(e) {
    // Tangani error
  }
}

/**
 * Mengisi Kolom AM pada sheet DAMEN dengan:
 * AM = Q - AL
 *
 * Kolom Q  = kolom ke-17
 * Kolom AL = kolom ke-38
 * Kolom AM = kolom ke-39
 */
function hitungKolomAMDamen() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (!sheetDamen) return;

    var lastRowDamen = sheetDamen.getLastRow();
    if (lastRowDamen < 2) return;

    // Ambil Kolom Q dan AL sekaligus.
    var rangeQ  = sheetDamen.getRange(2, 17, lastRowDamen - 1, 1).getValues();
    var rangeAL = sheetDamen.getRange(2, 38, lastRowDamen - 1, 1).getValues();

    var hasilAM = [];

    for (var i = 0; i < rangeQ.length; i++) {
      var valQ = parseFloat(rangeQ[i][0]) || 0;
      var valAL = parseFloat(rangeAL[i][0]) || 0;

      // AM = Q - AL
      hasilAM.push([valQ - valAL]);
    }

    // Simpan hasil ke Kolom AM (kolom ke-39)
    sheetDamen.getRange(2, 39, hasilAM.length, 1).setValues(hasilAM);

  } catch(e) {
    // Tangani error tanpa mengganggu proses utama
  }
}

/**
 * Mengisi Kolom L pada sheet INPUT dengan total Kolom AM pada sheet DAMEN
 * yang dikelompokkan berdasarkan ID INPUT (DAMEN kolom A).
 *
 * INPUT: A = ID INPUT, L = hasil total AM
 * DAMEN: A = ID INPUT, AM = nilai yang dijumlahkan
 */
function hitungKolomLInputDariAMDamen() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetInput = ss.getSheetByName("INPUT");
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (!sheetInput || !sheetDamen) return;

    var lastInputRow = sheetInput.getLastRow();
    var lastDamenRow = sheetDamen.getLastRow();
    if (lastInputRow < 2) return;

    // Bentuk map: ID INPUT -> total AM DAMEN
    var damenMap = {};
    if (lastDamenRow >= 2) {
      var damenData = sheetDamen.getRange(2, 1, lastDamenRow - 1, 39).getValues();
      for (var i = 0; i < damenData.length; i++) {
        var id = String(damenData[i][0] == null ? "" : damenData[i][0]).trim();
        if (!id) continue;

        var am = parseFloat(damenData[i][38]); // AM = kolom ke-39
        if (isNaN(am)) am = 0;
        damenMap[id] = (damenMap[id] || 0) + am;
      }
    }

    // Ambil ID INPUT dari kolom A, lalu tulis hasil ke kolom L.
    var inputIds = sheetInput.getRange(2, 1, lastInputRow - 1, 1).getValues();
    var hasilL = [];
    for (var j = 0; j < inputIds.length; j++) {
      var inputId = String(inputIds[j][0] == null ? "" : inputIds[j][0]).trim();
      hasilL.push([inputId && damenMap[inputId] !== undefined ? damenMap[inputId] : 0]);
    }

    sheetInput.getRange(2, 12, hasilL.length, 1).setValues(hasilL); // L
  } catch (e) {
    Logger.log("Error hitungKolomLInputDariAMDamen: " + e.message);
  }
}

/**
 * Mengisi Kolom AO pada sheet DAMEN dengan Tahun berformat YYYY
 * berdasarkan tanggal pada Kolom J.
 *
 * Kolom J  = kolom ke-10
 * Kolom AO = kolom ke-41
 */
function hitungKolomAODamen() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (!sheetDamen) return;

    var lastRowDamen = sheetDamen.getLastRow();
    if (lastRowDamen < 2) return;

    // Ambil nilai Date dari Kolom J
    var rangeJ = sheetDamen.getRange(2, 10, lastRowDamen - 1, 1).getValues();
    var hasilAO = [];

    for (var i = 0; i < rangeJ.length; i++) {
      var tglVal = rangeJ[i][0];

      if (tglVal === "" || tglVal === null || tglVal === undefined) {
        hasilAO.push([""]);
        continue;
      }

      var dObj = new Date(tglVal);

      if (!isNaN(dObj.getTime())) {
        // Tahun 4 digit, contoh: 2026
        hasilAO.push([
          Utilities.formatDate(
            dObj,
            Session.getScriptTimeZone(),
            "yyyy"
          )
        ]);
      } else {
        // Jika Kolom J bukan tanggal yang valid, kosongkan AO
        hasilAO.push([""]);
      }
    }

    // Simpan hasil ke Kolom AO (kolom ke-41)
    sheetDamen.getRange(2, 41, hasilAO.length, 1).setValues(hasilAO);

  } catch(e) {
    // Tangani error tanpa mengganggu proses utama
  }
}


/**
 * ============================================
 * OUTLET KONTRAK
 * ============================================
 * Sumber kontrak:
 * - Sheet DATA OUTLET BWS
 * - START = Kolom E
 * - END   = Kolom F
 * - TARGET = header "TARGET" bila ada, fallback Kolom G
 *
 * Actual volume:
 * - Sheet DAMEN
 * - DATE = Kolom J
 * - VOLUME = Kolom P
 * - STATUS CETAK harus APPROVE bila kolom tersebut tersedia
 *
 * Point Trip tahun berjalan:
 * - Kategori = DAMEN Kolom AN
 * - Volume = DAMEN Kolom P
 * - Point rate = Sheet DATA TRIP BWS
 * - Channel BP -> Volume / 9.6 x Point
 * - Channel CAS -> Volume / 12 x Point
 */
function parseTanggalBWS(value) {
  if (value === null || value === undefined || value === "") return null;
  var d = value instanceof Date ? new Date(value) : new Date(value);
  if (isNaN(d.getTime())) {
    var text = String(value).trim();
    var m = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (m) {
      d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    }
  }
  return isNaN(d.getTime()) ? null : d;
}

function getOutletKontrakData_BWS_UNCACHED(salesName, usernameInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var currentYear = new Date().getFullYear();
    var isAdmin = (usernameInput && String(usernameInput).trim().toLowerCase() === "admin");

    // =========================
    // 1. MASTER OUTLET / KONTRAK
    // =========================
    // Cari sheet DATA OUTLET BWS secara fleksibel:
    // - abaikan spasi awal/akhir
    // - tidak sensitif huruf besar/kecil
    // Ini menghindari error jika nama tab memiliki spasi tersembunyi.
    var sheetOutlet = null;
    var targetSheetName = "DATA OUTLET BWS";
    var targetNormalized = targetSheetName.replace(/\s+/g, " ").trim().toUpperCase();
    var allSheets = ss.getSheets();

    for (var si = 0; si < allSheets.length; si++) {
      var actualSheetName = String(allSheets[si].getName() || "");
      var normalizedSheetName = actualSheetName.replace(/\s+/g, " ").trim().toUpperCase();
      if (normalizedSheetName === targetNormalized) {
        sheetOutlet = allSheets[si];
        break;
      }
    }

    if (!sheetOutlet) {
      var daftarSheet = allSheets.map(function(sh) { return sh.getName(); });
      return {
        success: false,
        error: 'Sheet "DATA OUTLET BWS" tidak ditemukan. Sheet yang tersedia: ' + daftarSheet.join(', ')
      };
    }

    var outletData = sheetOutlet.getDataRange().getValues();
    if (outletData.length <= 1) {
      return { success: true, data: [] };
    }

    var oh = outletData[0].map(function(h) {
      return String(h).trim().toUpperCase();
    });

    var idxOutlet = oh.indexOf("OUTLET");
    var idxCust = oh.indexOf("CUST.CODE");
    if (idxCust === -1) idxCust = oh.indexOf("CUST CODE");
    var idxDsr = oh.indexOf("DSR");
    if (idxDsr === -1) idxDsr = oh.indexOf("SALES");
    var idxChannel = oh.indexOf("CHANNEL");

    if (idxOutlet === -1) idxOutlet = 2;
    if (idxCust === -1) idxCust = 1;
    if (idxDsr === -1) idxDsr = 3;
    if (idxChannel === -1) idxChannel = 3;

    // Sesuai permintaan: Start dan End berada pada kolom E dan F.
    var idxStart = 4;  // E
    var idxEnd = 5;    // F

    // Target dicari berdasarkan header TARGET; fallback ke kolom G.
    var idxTarget = oh.indexOf("TARGET");
    if (idxTarget === -1) idxTarget = oh.indexOf("TARGET VOLUME");
    if (idxTarget === -1) idxTarget = 6; // G

    var contracts = [];
    var allowedOutletMap = {};

    // Ambil SEMUA outlet dari DATA OUTLET BWS.
    // Baris tidak dibuang hanya karena Start/End kosong; outlet tetap tampil.
    // Start/End/Target digunakan bila nilainya valid untuk perhitungan kontrak.
    for (var i = 1; i < outletData.length; i++) {
      var row = outletData[i];
      var outletName = String(row[idxOutlet] == null ? "" : row[idxOutlet]).trim();
      if (!outletName) continue;

      var rowDsr = String(row[idxDsr] == null ? "" : row[idxDsr]).trim();
      if (!isAdmin && rowDsr && rowDsr.toUpperCase() !== String(salesName || "").trim().toUpperCase()) {
        continue;
      }

      var startDate = parseTanggalBWS(row[idxStart]);
      var endDate = parseTanggalBWS(row[idxEnd]);
      var hasPeriod = !!(startDate && endDate && endDate.getTime() >= startDate.getTime());

      if (startDate) startDate.setHours(0, 0, 0, 0);
      if (endDate) endDate.setHours(23, 59, 59, 999);

      var targetRaw = row[idxTarget];
      var target = parseFloat(String(targetRaw == null ? "" : targetRaw).replace(/[^0-9,.-]/g, "").replace(',', '.')) || 0;
      var channel = String(row[idxChannel] == null ? "" : row[idxChannel]).trim().toUpperCase();
      var custCode = String(row[idxCust] == null ? "" : row[idxCust]).trim();

      var contract = {
        outlet: outletName,
        custCode: custCode,
        channel: channel,
        dsr: rowDsr,
        start: startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        end: endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        startMs: hasPeriod ? startDate.getTime() : null,
        endMs: hasPeriod ? endDate.getTime() : null,
        target: target,
        actual: 0,
        gap: 0,
        percentage: 0,
        totalPoint: 0,
        pointByCategory: []
      };

      contracts.push(contract);
      var outletKey = outletName.toLowerCase();
      if (!allowedOutletMap[outletKey]) allowedOutletMap[outletKey] = [];
      allowedOutletMap[outletKey].push(contract);
    }

    // =========================
    // 2. DATA TRIP BWS
    // =========================
    var sheetTrip = ss.getSheetByName("DATA TRIP BWS");
    var tripData = sheetTrip ? sheetTrip.getDataRange().getValues() : [];
    var tripPointMap = {};

    if (tripData.length > 1) {
      var th = tripData[0].map(function(h) {
        return String(h).trim().toUpperCase();
      });

      var idxKatTrip = 0;
      var idxPointTrip = 1;

      for (var h = 0; h < th.length; h++) {
        if (
          th[h].indexOf("KAT") !== -1 ||
          th[h].indexOf("PRODUK") !== -1 ||
          th[h].indexOf("SKU") !== -1
        ) {
          idxKatTrip = h;
        }

        if (
          th[h].indexOf("POINT") !== -1 ||
          th[h].indexOf("POIN") !== -1 ||
          th[h].indexOf("VAL") !== -1
        ) {
          idxPointTrip = h;
        }
      }

      for (var t = 1; t < tripData.length; t++) {
        var katTrip = String(tripData[t][idxKatTrip] || "").trim().toUpperCase();
        var pointTrip = parseFloat(tripData[t][idxPointTrip]) || 0;

        if (katTrip) {
          tripPointMap[katTrip] = pointTrip;
        }
      }
    }

    // =========================
    // 3. DATA DAMEN
    // =========================
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (!sheetDamen) {
      return { success: false, error: 'Sheet "DAMEN" tidak ditemukan.' };
    }

    var damenData = sheetDamen.getDataRange().getValues();
    if (damenData.length <= 1) {
      return {
        success: true,
        year: currentYear,
        data: contracts.map(function(c) {
          return {
            outlet: c.outlet,
            custCode: c.custCode,
            channel: c.channel,
            dsr: c.dsr,
            start: c.start,
            end: c.end,
            target: c.target,
            actual: 0,
            gap: 0,
            percentage: 0,
            totalPoint: 0,
            pointByCategory: []
          };
        })
      };
    }

    var dh = damenData[0].map(function(h) {
      return String(h).trim().toUpperCase();
    });

    var dIdxOutlet = dh.indexOf("OUTLET");
    if (dIdxOutlet === -1) dIdxOutlet = 2;

    var dIdxDate = dh.indexOf("DATE");
    if (dIdxDate === -1) dIdxDate = 9; // J

    // Volume utama pada aplikasi saat ini berada di Kolom P.
    var dIdxVolume = 15; // P

    // Kategori Trip pada Kolom AN.
    var dIdxCategory = 39; // AN

    var dIdxStatus = -1;
    for (var ds = 0; ds < dh.length; ds++) {
      if (
        dh[ds] === "STATUS CETAK" ||
        dh[ds] === "STATUS_CETAK" ||
        dh[ds] === "CETAK"
      ) {
        dIdxStatus = ds;
        break;
      }
    }

    var currentYearStart = new Date(currentYear, 0, 1);
    currentYearStart.setHours(0, 0, 0, 0);
    var currentYearEnd = new Date(currentYear, 11, 31);
    currentYearEnd.setHours(23, 59, 59, 999);

    // =========================
    // 4. AGREGASI DATA DAMEN SEKALI SAJA
    //    Hanya simpan row untuk outlet yang memang ada di DATA OUTLET BWS.
    // =========================
    var damenByOutlet = {};

    for (var d = 1; d < damenData.length; d++) {
      var dr = damenData[d];

      if (dIdxStatus !== -1 && String(dr[dIdxStatus] == null ? '' : dr[dIdxStatus]).trim().toUpperCase() !== 'APPROVE') {
        continue;
      }

      var damenOutletKey = String(dr[dIdxOutlet] == null ? '' : dr[dIdxOutlet]).trim().toLowerCase();
      if (!damenOutletKey || !allowedOutletMap[damenOutletKey]) continue;

      if (!damenByOutlet[damenOutletKey]) damenByOutlet[damenOutletKey] = [];
      damenByOutlet[damenOutletKey].push(dr);
    }

    // =========================
    // 5. HITUNG ACTUAL + POINT PER KONTRAK
    // =========================
    for (var cIdx = 0; cIdx < contracts.length; cIdx++) {
      var c = contracts[cIdx];
      var categories = {};
      var outletRows = damenByOutlet[String(c.outlet).trim().toLowerCase()] || [];

      for (var or = 0; or < outletRows.length; or++) {
        var dr = outletRows[or];
        var rowTanggal = dr[dIdxDate];
        var dDate = new Date(rowTanggal);
        if (isNaN(dDate.getTime())) continue;

        var volume = parseFloat(dr[dIdxVolume]) || 0;

        // Actual hanya selama periode kontrak.
        if (c.startMs !== null && c.endMs !== null && dDate.getTime() >= c.startMs && dDate.getTime() <= c.endMs) {
          c.actual += volume;
        }

        // Point Trip hanya tahun berjalan.
        if (dDate.getTime() >= currentYearStart.getTime() && dDate.getTime() <= currentYearEnd.getTime()) {
          var kategori = String(dr[dIdxCategory] || '').trim();
          if (!kategori) kategori = 'Lainnya';

          var katUpper = kategori.toUpperCase();
          var pointRate = tripPointMap[katUpper] || 0;

          if (pointRate === 0 && Object.keys(tripPointMap).length > 0) {
            for (var key in tripPointMap) {
              if (katUpper.indexOf(key) !== -1 || key.indexOf(katUpper) !== -1) {
                pointRate = tripPointMap[key];
                break;
              }
            }
          }

          var divisor = 0;
          var channelUpper = String(c.channel || '').toUpperCase();
          if (channelUpper.indexOf('BP') !== -1) {
            divisor = 9.6;
          } else if (channelUpper.indexOf('CAS') !== -1) {
            divisor = 12;
          }

          var kartonTrip = divisor > 0 ? volume / divisor : 0;
          var pointValue = kartonTrip * pointRate;

          if (!categories[kategori]) {
            categories[kategori] = {
              kategori: kategori,
              volume: 0,
              karton: 0,
              pointRate: pointRate,
              point: 0,
              divisor: divisor
            };
          }

          categories[kategori].volume += volume;
          categories[kategori].karton += kartonTrip;
          categories[kategori].point += pointValue;
          c.totalPoint += pointValue;
        }
      }

      c.gap = c.actual - c.target;
      c.percentage = c.target > 0 ? (c.actual / c.target) * 100 : 0;

      c.pointByCategory = Object.keys(categories).map(function(k) {
        var x = categories[k];
        return {
          kategori: x.kategori,
          volume: Number(x.volume.toFixed(2)),
          karton: Number(x.karton.toFixed(2)),
          pointRate: x.pointRate,
          point: Number(x.point.toFixed(2)),
          divisor: x.divisor
        };
      }).sort(function(a, b) {
        return b.point - a.point;
      });

      c.totalPoint = Number(c.totalPoint.toFixed(2));
      c.gap = Number(c.actual || 0) - Number(c.target || 0);

      delete c.startMs;
      delete c.endMs;
    }

    return {
      success: true,
      year: currentYear,
      data: contracts
    };

  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Detail Outlet Kontrak.
 * Data sudah dihitung oleh getOutletKontrakData_BWS(), lalu difilter berdasarkan outlet.
 */
function getOutletKontrakDetail(outletName, custCode, start, end) {
  try {
    var result = getOutletKontrakData_BWS("Admin", "admin");
    if (!result || !result.success) return result;

    var found = result.data.filter(function(x) {
      var sameOutlet = String(x.outlet).trim().toLowerCase() === String(outletName).trim().toLowerCase();
      var sameCust = !custCode || String(x.custCode).trim().toLowerCase() === String(custCode).trim().toLowerCase();
      var sameStart = !start || String(x.start) === String(start);
      var sameEnd = !end || String(x.end) === String(end);
      return sameOutlet && sameCust && sameStart && sameEnd;
    });

    return {
      success: true,
      year: result.year,
      data: found.length ? found[0] : null
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getMonitoringOutletData_UNCACHED(salesName, usernameInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    
    // 1. Ambil master outlet dan channel-nya dari sheet OUTLET
    var sheetOutletMaster = ss.getSheetByName("OUTLET");
    var assignedOutlets = [];
    var outletChannelMap = {};
    var outletDayVisitMap = {};
    var isAdmin = (usernameInput && String(usernameInput).trim().toLowerCase() === 'admin');
    
    if (sheetOutletMaster) {
      var dataOM = sheetOutletMaster.getDataRange().getValues();
      if (dataOM.length > 1) {
        var hOM = dataOM[0].map(function(h) { return String(h).trim().toUpperCase(); });
        var idxOutletName = hOM.indexOf("OUTLET") !== -1 ? hOM.indexOf("OUTLET") : 2;
        var idxChannel = hOM.indexOf("CHANNEL") !== -1 ? hOM.indexOf("CHANNEL") : 3;
        var idxDsrOM = hOM.indexOf("DSR") !== -1 ? hOM.indexOf("DSR") : (hOM.indexOf("SALES") !== -1 ? hOM.indexOf("SALES") : 11);
        
        for (var o = 1; o < dataOM.length; o++) {
          var outName = String(dataOM[o][idxOutletName]).trim();
          var outChannel = String(dataOM[o][idxChannel]).trim().toUpperCase();
          var outDsr = String(dataOM[o][idxDsrOM]).trim();
          var idxDayVisit = hOM.indexOf("DAY VISIT");
          if (idxDayVisit === -1) idxDayVisit = 0;
          var outDayVisit = String(dataOM[o][idxDayVisit] == null ? "" : dataOM[o][idxDayVisit]).trim();
          
          if (outName) {
            outletChannelMap[outName] = outChannel;
            outletDayVisitMap[outName.toUpperCase().replace(/\s+/g, " ").trim()] = outDayVisit;
            if (isAdmin || outDsr.toUpperCase() === String(salesName).trim().toUpperCase()) {
              if (assignedOutlets.indexOf(outName) === -1) {
                assignedOutlets.push(outName);
              }
            }
          }
        }
      }
    }

    // 2. Ambil TARGET OUTLET berdasarkan OUTLET + QUARTAL
    // Struktur sheet TARGET OUTLET:
    // A DAY VISIT | B OUTLET | C CUST.CODE | D DSR | E CHANNEL |
    // F TARGET QUARTER | G TARGET MONTHLY | H TARGET QUARTER LITER |
    // I TARGET MONTHLY LITER | J QUARTAL
    var outletTargetMap = {};
    var targetCustMap = {};
    var targetChannelMap = {};
    var sheetTargetOutlet = ss.getSheetByName("TARGET OUTLET");

    function normalizeQuarterLabel_(v) {
      var s = String(v == null ? "" : v).trim().toUpperCase();
      if (!s) return "";
      var m = s.match(/QUARTAL\s*([1-4])\s*(\d{4})/i);
      if (m) return "Q" + m[1] + " " + m[2];
      m = s.match(/Q\s*([1-4])\s*(?:-|\/|\s)*?(\d{4})/i);
      if (m) return "Q" + m[1] + " " + m[2];
      return s;
    }

    if (sheetTargetOutlet) {
      var targetOutletData = sheetTargetOutlet.getDataRange().getValues();
      if (targetOutletData.length > 1) {
        var th = targetOutletData[0].map(function(h) {
          return String(h == null ? "" : h).trim().toUpperCase();
        });
        var tOutlet = th.indexOf("OUTLET"); if (tOutlet === -1) tOutlet = 1;
        var tCust = th.indexOf("CUST.CODE"); if (tCust === -1) tCust = th.indexOf("CUST CODE"); if (tCust === -1) tCust = 2;
        var tDsr = th.indexOf("DSR"); if (tDsr === -1) tDsr = 3;
        var tQuarterTarget = th.indexOf("TARGET QUARTER"); if (tQuarterTarget === -1) tQuarterTarget = 5;
        var tQuarterLabel = th.indexOf("QUARTAL"); if (tQuarterLabel === -1) tQuarterLabel = 9;

        for (var ti = 1; ti < targetOutletData.length; ti++) {
          var tr = targetOutletData[ti];
          var to = String(tr[tOutlet] == null ? "" : tr[tOutlet]).trim();
          if (!to) continue;

          var tc = String(tr[tCust] == null ? "" : tr[tCust]).trim();
          var td = String(tr[tDsr] == null ? "" : tr[tDsr]).trim();
          if (!isAdmin && td && td.toUpperCase() !== String(salesName || "").trim().toUpperCase()) continue;

          var tq = normalizeQuarterLabel_(tr[tQuarterLabel]);
          if (!tq) {
            // Jika kolom QUARTAL kosong, coba tentukan dari tanggal/day visit.
            var dv = new Date(tr[0]);
            if (!isNaN(dv.getTime())) {
              tq = "Q" + (Math.floor(dv.getMonth() / 3) + 1) + " " + dv.getFullYear();
            }
          }
          if (!tq) continue;

          var targetQuarter = Number(String(tr[tQuarterTarget] == null ? "" : tr[tQuarterTarget]).replace(/[^0-9,.-]/g, "").replace(",", ".")) || 0;
          var outletNorm = to.toUpperCase().replace(/\s+/g, " ").trim();
          var custNorm = tc.toUpperCase().replace(/\s+/g, " ").trim();
          var key = outletNorm.toLowerCase() + "||" + tq;
          var custKey = key + "||" + custNorm.toLowerCase();
          if (!targetCustMap[custKey]) targetCustMap[custKey] = true;
          var tcChannel = String(tr[4] == null ? "" : tr[4]).trim().toUpperCase();
          if (tcChannel) targetChannelMap[custKey] = tcChannel;

          // Jika terdapat lebih dari satu baris outlet pada kuartal yang sama,
          // target dijumlahkan agar tidak kehilangan data.
          if (!outletTargetMap[key]) {
            outletTargetMap[key] = {
              outlet: to,
              quarter: tq,
              targetQuarter: 0,
              targetQuarterLiter: 0,
              targetMonthly: 0,
              targetMonthlyLiter: 0
            };
          }
          outletTargetMap[key].targetQuarter += targetQuarter;

          var targetQuarterLiter = Number(String(tr[7] == null ? "" : tr[7]).replace(/[^0-9,.-]/g, "").replace(",", ".")) || 0;
          var targetMonthly = Number(String(tr[6] == null ? "" : tr[6]).replace(/[^0-9,.-]/g, "").replace(",", ".")) || 0;
          var targetMonthlyLiter = Number(String(tr[8] == null ? "" : tr[8]).replace(/[^0-9,.-]/g, "").replace(",", ".")) || 0;
          outletTargetMap[key].targetQuarterLiter += targetQuarterLiter;
          outletTargetMap[key].targetMonthly += targetMonthly;
          outletTargetMap[key].targetMonthlyLiter += targetMonthlyLiter;
        }
      }
    }

    // 3. Ambil data dari sheet DAMEN
    var sheetDamen = ss.getSheetByName("DAMEN");
    if (!sheetDamen) return {};
    var dataDamen = sheetDamen.getDataRange().getValues();
    if (dataDamen.length <= 1) return {};
    
    var hDamen = dataDamen[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var cDsr = hDamen.indexOf("DSR") !== -1 ? hDamen.indexOf("DSR") : 11;
    var cOutlet = hDamen.indexOf("OUTLET") !== -1 ? hDamen.indexOf("OUTLET") : 2;
    var cDate = hDamen.indexOf("DATE") !== -1 ? hDamen.indexOf("DATE") : 9;
    var cVolume = hDamen.indexOf("VOLUME") !== -1 ? hDamen.indexOf("VOLUME") : -1;
    var cCust = hDamen.indexOf("CUST.CODE") !== -1 ? hDamen.indexOf("CUST.CODE") : hDamen.indexOf("CUST CODE");
    var cChannelDamen = hDamen.indexOf("CHANNEL") !== -1 ? hDamen.indexOf("CHANNEL") : -1;
    
    var cStatusCetak = -1;
    for (var c = 0; c < hDamen.length; c++) {
      if (hDamen[c] === "STATUS CETAK" || hDamen[c] === "STATUS_CETAK" || hDamen[c] === "CETAK") {
        cStatusCetak = c;
        break;
      }
    }

    var aggregated = {};
    var activeQuarters = {};
    
    for (var i = 1; i < dataDamen.length; i++) {
      var tglVal = dataDamen[i][cDate];
      var dObj = new Date(tglVal);
      if (!isNaN(dObj.getTime())) {
        var year = dObj.getFullYear();
        var monthIndex = dObj.getMonth();
        var quarter = Math.floor(monthIndex / 3) + 1;
        var qKey = "Q" + quarter + " " + year;
        activeQuarters[qKey] = true;
      }
    }
    
    // Outlet yang mempunyai target tetap ditampilkan walaupun belum ada transaksi.
    Object.keys(outletTargetMap).forEach(function(tk) {
      var targetOutletName = tk.split("||")[0];
      var ti2 = outletTargetMap[tk];
      var targetDisplayOutlet = String(ti2.outlet || "").trim();
      if (targetDisplayOutlet && assignedOutlets.indexOf(targetDisplayOutlet) === -1) {
        assignedOutlets.push(targetDisplayOutlet);
      }
    });

    if (Object.keys(activeQuarters).length === 0) {
      var now = new Date();
      var defQ = "Q" + (Math.floor(now.getMonth() / 3) + 1) + " " + now.getFullYear();
      activeQuarters[defQ] = true;
    }

    // Tambahkan kuartal yang hanya memiliki TARGET OUTLET walaupun belum ada order.
    Object.keys(outletTargetMap).forEach(function(tk) {
      var qOnly = tk.split("||")[1];
      if (qOnly) activeQuarters[qOnly] = true;
    });

    Object.keys(activeQuarters).forEach(function(qKey) {
      aggregated[qKey] = {};
      assignedOutlets.forEach(function(outletName) {
        var targetInfoInit = outletTargetMap[outletName.toUpperCase().replace(/\s+/g, " ").trim().toLowerCase() + "||" + qKey] || {};
        aggregated[qKey][outletName] = {
          outlet: outletName,
          dayVisit: outletDayVisitMap[String(outletName || "").toUpperCase().replace(/\s+/g, " ").trim()] || "",
          quarter: qKey,
          totalQuarterVol: 0,
          targetQuarter: Number(targetInfoInit.targetQuarter || 0),
          targetQuarterLiter: Number(targetInfoInit.targetQuarterLiter || 0),
          targetMonthly: Number(targetInfoInit.targetMonthly || 0),
          targetMonthlyLiter: Number(targetInfoInit.targetMonthlyLiter || 0),
          m1: 0,
          m2: 0,
          m3: 0
        };
      });
    });

    // 3. Proses perhitungan dan konversi ke karton
    for (var i = 1; i < dataDamen.length; i++) {
      var statusCetakVal = cStatusCetak !== -1 ? String(dataDamen[i][cStatusCetak]).trim().toUpperCase() : "";
      if (cStatusCetak !== -1 && statusCetakVal !== "APPROVE") {
        continue;
      }

      var rowSales = String(dataDamen[i][cDsr]).trim();
      if (isAdmin || rowSales.toUpperCase() === String(salesName).trim().toUpperCase()) {
        var outletName = String(dataDamen[i][cOutlet] || "").trim();
        if (!outletName) continue;

        if (assignedOutlets.indexOf(outletName) === -1) {
          assignedOutlets.push(outletName);
        }

        var tglVal = dataDamen[i][cDate];
        var rawVolume = cVolume !== -1 ? (Number(dataDamen[i][cVolume]) || 0) : 0;
        
        var dObj = new Date(tglVal);
        if (isNaN(dObj.getTime())) continue;
        
        var year = dObj.getFullYear();
        var monthIndex = dObj.getMonth();
        var quarter = Math.floor(monthIndex / 3) + 1;
        var quarterKey = "Q" + quarter + " " + year;
        var monthInQuarter = (monthIndex % 3) + 1;

        // Monitoring outlet dihitung berdasarkan pasangan OUTLET + CUST.CODE.
        // Hanya CUST.CODE yang mempunyai target pada TARGET OUTLET untuk kuartal
        // tersebut yang masuk ke actual. Ini mencegah transaksi customer lain
        // pada outlet yang sama ikut menghitung pencapaian.
        // ================================================================
        // AKTUAL MONITORING = OUTLET + CUST.CODE + QUARTAL
        // ================================================================
        // TARGET OUTLET adalah master untuk menentukan CUST.CODE mana yang
        // boleh dihitung. Jangan menjumlahkan transaksi customer lain pada
        // outlet yang sama.
        var custCodeVal = cCust !== -1 ? String(dataDamen[i][cCust] == null ? "" : dataDamen[i][cCust]).trim() : "";
        var outletKeyNorm = outletName.toUpperCase().replace(/\s+/g, " ").trim();
        var custKeyNorm = custCodeVal.toUpperCase().replace(/\s+/g, " ").trim();
        var targetCustKey = outletKeyNorm.toLowerCase() + "||" + quarterKey + "||" + custKeyNorm.toLowerCase();

        // Jika TARGET OUTLET tersedia, transaksi hanya boleh masuk jika
        // CUST.CODE-nya benar-benar ada pada TARGET OUTLET untuk outlet + quarter.
        // CUST.CODE kosong juga TIDAK dihitung.
        if (Object.keys(targetCustMap).length > 0) {
          if (!custKeyNorm || !targetCustMap[targetCustKey]) continue;
        }

        // Channel TARGET OUTLET menjadi acuan utama karena target dan divisor
        // ditetapkan berdasarkan CUST.CODE tersebut. DAMEN/OUTLET hanya fallback.
        var channelVal = targetChannelMap[targetCustKey] || "";
        if (!channelVal && cChannelDamen !== -1) {
          channelVal = String(dataDamen[i][cChannelDamen] || "").trim().toUpperCase();
        }
        if (!channelVal) channelVal = outletChannelMap[outletName] || "";

        var volumeInCtn = rawVolume;
        if (channelVal === "MCO-4T") {
          volumeInCtn = rawVolume / 9.6;
        } else if (channelVal === "PCO") {
          volumeInCtn = rawVolume / 12;
        }

        if (!aggregated[quarterKey]) {
          aggregated[quarterKey] = {};
        }
        if (!aggregated[quarterKey][outletName]) {
          var targetInfo = outletTargetMap[outletName.toUpperCase().replace(/\s+/g, " ").trim().toLowerCase() + "||" + quarterKey] || {};
          aggregated[quarterKey][outletName] = {
            outlet: outletName,
            quarter: quarterKey,
            totalQuarterVol: 0,
            targetQuarter: Number(targetInfo.targetQuarter || 0),
            targetQuarterLiter: Number(targetInfo.targetQuarterLiter || 0),
            targetMonthly: Number(targetInfo.targetMonthly || 0),
            targetMonthlyLiter: Number(targetInfo.targetMonthlyLiter || 0),
            m1: 0,
            m2: 0,
            m3: 0
          };
        }

        var item = aggregated[quarterKey][outletName];
        item.totalQuarterVol += volumeInCtn;
        if (monthInQuarter === 1) item.m1 += volumeInCtn;
        if (monthInQuarter === 2) item.m2 += volumeInCtn;
        if (monthInQuarter === 3) item.m3 += volumeInCtn;
      }
    }

    return aggregated;
  } catch(e) {
    return {};
  }
}

function getProdukDibeliOutletQuarter(outletName, quarterKey) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DAMEN");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    // Detail Monitoring hanya menggunakan:
    // Kolom C = OUTLET (untuk filter outlet)
    // Kolom J = DATE   -> ditampilkan sebagai "Tanggal"
    // Kolom M =       -> ditampilkan sebagai "Produk"
    var idxOutlet = 2;
    var idxTanggal = 9;
    var idxProduk = 12;

    // Status cetak tetap menjadi filter bila kolom tersebut ada.
    var headers = data[0].map(function(h) {
      return String(h).trim().toUpperCase();
    });
    var idxStatusCetak = -1;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === "STATUS CETAK" ||
          headers[h] === "STATUS_CETAK" ||
          headers[h] === "CETAK") {
        idxStatusCetak = h;
        break;
      }
    }

    var parts = String(quarterKey || "").split(" ");
    var q = parts[0] || "";
    var year = parseInt(parts[1], 10) || new Date().getFullYear();

    var startMonth = 0, endMonth = 11;
    if (q === "Q1") { startMonth = 0; endMonth = 2; }
    else if (q === "Q2") { startMonth = 3; endMonth = 5; }
    else if (q === "Q3") { startMonth = 6; endMonth = 8; }
    else if (q === "Q4") { startMonth = 9; endMonth = 11; }

    var targetOutletClean = String(outletName || "").trim().toLowerCase();
    var resultList = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[idxOutlet] || "").trim().toLowerCase() !== targetOutletClean) continue;

      if (idxStatusCetak !== -1 &&
          String(row[idxStatusCetak] || "").trim().toUpperCase() !== "APPROVE") {
        continue;
      }

      var d = new Date(row[idxTanggal]);
      if (isNaN(d.getTime())) continue;
      if (d.getFullYear() !== year) continue;
      if (d.getMonth() < startMonth || d.getMonth() > endMonth) continue;

      resultList.push({
        produk: row[idxProduk] || "-",
        tanggal: Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")
      });
    }

    resultList.sort(function(a, b) {
      return new Date(b.tanggal).getTime() - new Date(a.tanggal).getTime();
    });

    return resultList;
  } catch (err) {
    Logger.log("Error getProdukDibeliOutletQuarter: " + err.message);
    return [];
  }
}

/* =========================================================
   TAMBAH CUSTOMER - MODULE
   ========================================================= */

/**
 * Mengambil master untuk Form Tambah Customer.
 *
 * DAY VISIT : TARGET OUTLET kolom A
 * OUTLET    : OUTLET
 * DSR       : OUTLET
 * CHANNEL   : OUTLET
 */
function getTambahCustomerMasterData() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dayVisitSet = {};
    var outletSet = {};
    var dsrSet = {};
    var channelSet = {};

    // DAY VISIT dari TARGET OUTLET kolom A
    var sheetTarget = ss.getSheetByName('TARGET OUTLET');
    if (sheetTarget && sheetTarget.getLastRow() >= 2) {
      var targetData = sheetTarget.getDataRange().getValues();
      for (var i = 1; i < targetData.length; i++) {
        var dayVisit = String(targetData[i][0] == null ? '' : targetData[i][0]).trim();
        if (dayVisit) {
          // Pertahankan nilai asli agar jika berupa teks tidak berubah.
          dayVisitSet[dayVisit] = true;
        }
      }
    }

    // OUTLET, DSR, CHANNEL dari OUTLET
    var sheetOutlet = ss.getSheetByName('OUTLET');
    if (sheetOutlet && sheetOutlet.getLastRow() >= 2) {
      var outletData = sheetOutlet.getDataRange().getValues();
      var headers = outletData[0].map(function(h) {
        return String(h == null ? '' : h).trim().toUpperCase();
      });

      var idxOutlet = headers.indexOf('OUTLET');
      var idxDsr = headers.indexOf('DSR');
      var idxChannel = headers.indexOf('CHANNEL');

      if (idxOutlet === -1) idxOutlet = 2;
      if (idxDsr === -1) idxDsr = 3;
      if (idxChannel === -1) idxChannel = 4;

      for (var j = 1; j < outletData.length; j++) {
        var outlet = String(outletData[j][idxOutlet] == null ? '' : outletData[j][idxOutlet]).trim();
        var dsr = String(outletData[j][idxDsr] == null ? '' : outletData[j][idxDsr]).trim();
        var channel = String(outletData[j][idxChannel] == null ? '' : outletData[j][idxChannel]).trim();

        if (outlet) outletSet[outlet.toUpperCase()] = true;
        if (dsr) dsrSet[dsr.toUpperCase()] = true;
        if (channel) channelSet[channel.toUpperCase()] = true;
      }
    }

    var dayVisitList = Object.keys(dayVisitSet).sort();

    // Fallback jika TARGET OUTLET belum mempunyai DAY VISIT.
    if (!dayVisitList.length) {
      dayVisitList = ['SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'];
    }

    return {
      success: true,
      dayVisit: dayVisitList,
      outlet: Object.keys(outletSet).sort(),
      dsr: Object.keys(dsrSet).sort(),
      channel: Object.keys(channelSet).sort()
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal mengambil master Tambah Customer: ' + e.message,
      dayVisit: [],
      outlet: [],
      dsr: [],
      channel: []
    };
  }
}


/**
 * Tambah Customer sekarang langsung menggunakan sheet OUTLET.
 *
 * Struktur data yang ditulis tetap sama seperti modul sebelumnya:
 *
 * A DAY VISIT
 * B OUTLET
 * C CUST.CODE
 * D DSR
 * E CHANNEL
 * F B - C          (disembunyikan)
 * G B - D          (disembunyikan)
 * H B - E          (disembunyikan)
 * I DATE
 * J kosong
 * K kosong
 * L kosong
 * M MAPS / LOCATION
 * N PHONE
 *
 * Perubahan hanya pada tujuan penyimpanan:
 * sebelumnya -> TAMBAH CUSTOMER
 * sekarang   -> OUTLET
 */
function ensureTambahCustomerSheet_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('OUTLET');

  if (!sheet) {
    throw new Error('Sheet "OUTLET" tidak ditemukan.');
  }

  // Format kolom tetap mengikuti struktur data customer.
  // Tidak membuat, menghapus, atau memindahkan kolom pada sheet OUTLET.
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('I:I').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('N:N').setNumberFormat('@');

  // F:H tetap disembunyikan seperti perilaku sebelumnya.
  sheet.hideColumns(6, 3);

  return sheet;
}


/**
 * Membuat CUST.CODE dengan prefix CNOO + Unique ID.
 * Contoh: CNOO-260918-A1B2C3D4E5F6
 */
function generateCustCode_() {
  var now = new Date();
  var datePrefix = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyMMdd'
  );

  var uniqueId = Utilities.getUuid()
    .replace(/-/g, '')
    .substring(0, 12)
    .toUpperCase();

  return 'CNOO-' + uniqueId;
}


/**
 * Simpan customer baru langsung ke sheet OUTLET.
 *
 * Selain tujuan sheet, logika existing tidak diubah.
 */
function simpanTambahCustomer(payload) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!payload) {
      return { success: false, message: 'Data customer tidak ditemukan.' };
    }

    var dayVisit = String(payload.dayVisit == null ? '' : payload.dayVisit).trim();
    var outlet = String(payload.outlet == null ? '' : payload.outlet).trim().toUpperCase();
    var dsr = String(payload.dsr == null ? '' : payload.dsr).trim().toUpperCase();
    var channel = String(payload.channel == null ? '' : payload.channel).trim().toUpperCase();
    var maps = String(payload.maps == null ? '' : payload.maps).trim();
    var phone = String(payload.phone == null ? '' : payload.phone).trim();

    if (!dayVisit) {
      return { success: false, message: 'DAY VISIT wajib dipilih.' };
    }
    if (!outlet) {
      return { success: false, message: 'OUTLET wajib diisi.' };
    }
    if (!dsr) {
      return { success: false, message: 'DSR wajib dipilih.' };
    }
    if (!channel) {
      return { success: false, message: 'CHANNEL wajib dipilih.' };
    }

    // Sekarang mengambil sheet OUTLET, bukan TAMBAH CUSTOMER.
    var sheet = ensureTambahCustomerSheet_();
    var now = new Date();
    var custCode = generateCustCode_();

    // Kolom F-H otomatis, tetapi disembunyikan.
    var valueBC = outlet + ' - ' + custCode;
    var valueBD = outlet + ' - ' + dsr;
    var valueBE = outlet + ' - ' + channel;

    var row = sheet.getLastRow() + 1;

    sheet.getRange(row, 1, 1, 14).setValues([[
      dayVisit,
      outlet,
      custCode,
      dsr,
      channel,
      valueBC,
      valueBD,
      valueBE,
      now,
      '',
      '',
      '',
      maps,
      phone
    ]]);

    sheet.getRange(row, 3).setNumberFormat('@');
    sheet.getRange(row, 9).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(row, 14).setNumberFormat('@');
    sheet.hideColumns(6, 3);

    clearAppDataCaches();

    return {
      success: true,
      custCode: custCode,
      message: 'Customer berhasil ditambahkan.'
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal menyimpan customer: ' + e.message
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

/* =========================================================
   TAMBAH CUSTOMER - TABLE / CRUD
   ========================================================= */
function getTambahCustomerTableData(salesName, usernameInput) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('OUTLET');
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h == null ? '' : h).trim().toUpperCase(); });
    var idxDay = headers.indexOf('DAY VISIT'); if (idxDay === -1) idxDay = 0;
    var idxOutlet = headers.indexOf('OUTLET'); if (idxOutlet === -1) idxOutlet = 1;
    var idxCust = headers.indexOf('CUST.CODE'); if (idxCust === -1) idxCust = 2;
    var idxDsr = headers.indexOf('DSR'); if (idxDsr === -1) idxDsr = 3;
    var idxChannel = headers.indexOf('CHANNEL'); if (idxChannel === -1) idxChannel = 4;
    var idxDate = headers.indexOf('DATE'); if (idxDate === -1) idxDate = 8;
    var idxMaps = headers.indexOf('MAPS / LOCATION'); if (idxMaps === -1) idxMaps = headers.indexOf('MAPS'); if (idxMaps === -1) idxMaps = 12;
    var idxPhone = headers.indexOf('PHONE'); if (idxPhone === -1) idxPhone = 13;

    var isAdmin = String(usernameInput || '').trim().toLowerCase() === 'admin';
    var targetSales = String(salesName || '').trim().toUpperCase();
    var tz = Session.getScriptTimeZone();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var dsr = String(data[i][idxDsr] == null ? '' : data[i][idxDsr]).trim();
      if (!isAdmin && dsr.toUpperCase() !== targetSales) continue;

      var dateValue = data[i][idxDate];
      if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
        dateValue = Utilities.formatDate(dateValue, tz, 'dd/MM/yyyy');
      } else {
        dateValue = String(dateValue == null ? '' : dateValue).trim();
      }

      result.push({
        rowNumber: i + 1,
        dayVisit: data[i][idxDay] == null ? '' : data[i][idxDay],
        outlet: data[i][idxOutlet] == null ? '' : data[i][idxOutlet],
        custCode: data[i][idxCust] == null ? '' : data[i][idxCust],
        dsr: dsr,
        channel: data[i][idxChannel] == null ? '' : data[i][idxChannel],
        date: dateValue,
        maps: data[i][idxMaps] == null ? '' : data[i][idxMaps],
        phone: data[i][idxPhone] == null ? '' : data[i][idxPhone]
      });
    }

    return result;
  } catch (e) {
    throw new Error('Gagal mengambil data customer: ' + e.message);
  }
}

function getTambahCustomerDetail(rowNumber, salesName, usernameInput) {
  var rows = getTambahCustomerTableData(salesName, usernameInput);
  var row = Number(rowNumber);
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].rowNumber) === row) return rows[i];
  }
  return null;
}

function updateTambahCustomer(rowNumber, payload, salesName, usernameInput) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ensureTambahCustomerSheet_();
    var row = Number(rowNumber);
    if (!row || row < 2 || row > sheet.getLastRow()) return { success:false, message:'Data customer tidak ditemukan.' };

    var current = getTambahCustomerDetail(row, salesName, usernameInput);
    if (!current) return { success:false, message:'Anda tidak memiliki akses ke customer ini.' };

    var dayVisit = String(payload && payload.dayVisit || '').trim();
    var outlet = String(payload && payload.outlet || '').trim().toUpperCase();
    var dsr = String(payload && payload.dsr || '').trim().toUpperCase();
    var channel = String(payload && payload.channel || '').trim().toUpperCase();
    var maps = String(payload && payload.maps || '').trim();
    var phone = String(payload && payload.phone || '').trim();
    var dateValue = String(payload && payload.date || '').trim();
    var isAdmin = String(usernameInput || '').trim().toLowerCase() === 'admin';

    if (!dayVisit || !outlet || !dsr || !channel) return { success:false, message:'DAY VISIT, OUTLET, DSR, dan CHANNEL wajib diisi.' };

    var custCode = String(current.custCode || '').trim();
    if (!custCode) custCode = generateCustCode_();
    var valueBC = outlet + ' - ' + custCode;
    var valueBD = outlet + ' - ' + dsr;
    var valueBE = outlet + ' - ' + channel;

    sheet.getRange(row, 1, 1, 8).setValues([[dayVisit, outlet, custCode, dsr, channel, valueBC, valueBD, valueBE]]);
    if (dateValue) {
      var parts = dateValue.split('-');
      if (parts.length === 3) {
        var dateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        if (!isNaN(dateObj.getTime())) sheet.getRange(row, 9).setValue(dateObj);
      }
    }
    sheet.getRange(row, 13).setValue(maps);
    sheet.getRange(row, 14).setValue(phone);
    sheet.getRange(row, 3).setNumberFormat('@');
    sheet.getRange(row, 14).setNumberFormat('@');
    sheet.hideColumns(6, 3);
    clearAppDataCaches();

    return { success:true, message:'Customer berhasil diperbarui.' };
  } catch (e) {
    return { success:false, message:'Gagal memperbarui customer: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function deleteTambahCustomer(rowNumber, salesName, usernameInput) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = ensureTambahCustomerSheet_();
    var row = Number(rowNumber);
    if (!row || row < 2 || row > sheet.getLastRow()) return { success:false, message:'Data customer tidak ditemukan.' };

    var current = getTambahCustomerDetail(row, salesName, usernameInput);
    if (!current) return { success:false, message:'Anda tidak memiliki akses ke customer ini.' };

    sheet.deleteRow(row);
    clearAppDataCaches();
    return { success:true, message:'Customer berhasil dihapus.' };
  } catch (e) {
    return { success:false, message:'Gagal menghapus customer: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

