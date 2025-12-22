/* Code.gs - Hoàn chỉnh và Đồng bộ */
const SS = SpreadsheetApp.getActive();

/* ===== WEB ===== */
function doGet(){
  init();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản lý Thu Chi V3')
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(f){ 
  return HtmlService.createHtmlOutputFromFile(f).getContent(); 
}

/* ===== INIT & HELPERS ===== */
function init(){
  let sh = SS.getSheetByName('Data');
  if(!sh){
    sh = SS.insertSheet('Data');
    sh.appendRow(['ID','Ngày','Loại','Nội dung','Số tiền','Tạo lúc']);
  }
  let us = SS.getSheetByName('Users');
  if(!us){
    us = SS.insertSheet('Users');
    us.appendRow(['Tài khoản','Mật khẩu','Tên']);
    us.appendRow(['admin','admin','Quản trị']);
    us.hideSheet();
  }
}

function getSheet(){ 
  init(); 
  return SS.getSheetByName('Data'); 
}

/* ===== LOGIN ===== */
function login(u,p){
  const rows = SS.getSheetByName('Users').getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==u && rows[i][1]==p) return {ok:true, name:rows[i][2]};
  }
  return {ok:false};
}

/* ===== BASIC LIST ===== */
function listData(){
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  const out = [];
  for(let i=1;i<data.length;i++){
    if(!data[i] || data[i].length===0) continue;
    let cellDate = data[i][1];
    let dateStr='';
    try{ 
      dateStr = Utilities.formatDate(new Date(cellDate), Session.getScriptTimeZone(), 'dd/MM/yyyy'); 
    } catch(e){ 
      dateStr = cellDate ? String(cellDate) : ''; 
    }
    out.push({
      id: String(data[i][0]),
      date: dateStr,
      type: data[i][2],
      content: data[i][3],
      amount: Number(data[i][4]||0)
    });
  }
  return out;
}

/* ===== CẦU NỐI JAVASCRIPT & SERVER (FIX LỖI) ===== */
function getDataFromServer(filterObj) {
  try {
    const params = {
      type: filterObj.type,
      mode: filterObj.mode,
      from: filterObj.dateFrom,
      to: filterObj.dateTo,
      year: filterObj.year,
      fromMonth: filterObj.monthFrom,
      toMonth: filterObj.monthTo,
      quarter: filterObj.quarter
    };
    return processFilterLogic(params);
  } catch (e) {
    throw new Error("Lỗi Server: " + e.message);
  }
}

/* ===== LOGIC LỌC DỮ LIỆU TRUNG TÂM ===== */
function processFilterLogic(params){
  const all = listData().reverse(); // Đảo lại để mới nhất lên đầu
  if(!params || Object.keys(params).length===0) return all;
  
  const res = [];
  let from = null, to = null;
  
  try {
    if(params.mode === 'date' || params.mode === 'range'){
      if(params.from) from = parseToDate(params.from);
      if(params.to) to = parseToDate(params.to);
    } else if(params.mode === 'month'){
      const y = Number(params.year || new Date().getFullYear());
      const fm = Number(params.fromMonth || 1);
      const tm = Number(params.toMonth || fm);
      from = new Date(y, fm - 1, 1);
      to = new Date(y, tm, 0); 
    } else if(params.mode === 'quarter'){
      const y = Number(params.year || new Date().getFullYear());
      const q = Number(params.quarter || 1);
      const fm = (q - 1) * 3 + 1;
      from = new Date(y, fm - 1, 1);
      to = new Date(y, fm + 2, 0);
    } else if(params.mode === 'year'){
      const y = Number(params.year || new Date().getFullYear());
      from = new Date(y, 0, 1);
      to = new Date(y, 11, 31);
    }
  } catch(e) { }

  all.forEach(r => {
    let rDate = null;
    try { 
      const p = r.date.split('/'); 
      rDate = new Date(Number(p[2]), Number(p[1])-1, Number(p[0])); 
    } catch(e) { }
    
    let ok = true;
    if(params.type && params.type !== '') ok = (r.type === params.type);
    if(ok && from && rDate) ok = (rDate >= from);
    if(ok && to && rDate) ok = (rDate <= to);
    if(ok) res.push(r);
  });
  return res;
}

/* ================= EXPORT EXCEL ================= */
function exportFilteredExcel(filterObj) {
  try {
    const params = {
      type: filterObj.type, mode: filterObj.mode, from: filterObj.dateFrom, to: filterObj.dateTo,
      year: filterObj.year, fromMonth: filterObj.monthFrom, toMonth: filterObj.monthTo, quarter: filterObj.quarter
    };
    
    const rows = processFilterLogic(params).reverse(); 
    const tempSs = SpreadsheetApp.create('TempExport');
    const ts = tempSs.getActiveSheet();
    
    ts.appendRow(['Ngày', 'Loại', 'Nội dung', 'Số tiền']);
    if (rows.length > 0) {
      const dataToSheet = rows.map(r => [r.date, r.type, r.content, Number(r.amount || 0)]);
      ts.getRange(2, 1, dataToSheet.length, 4).setValues(dataToSheet);
      ts.getRange(2, 4, dataToSheet.length, 1).setNumberFormat("#,##0");
    }
    
    SpreadsheetApp.flush();

    // 🔑 CHÌA KHÓA: Lấy dữ liệu file dưới dạng Base64
    const url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const bytes = response.getContent();
    const base64 = Utilities.base64Encode(bytes);
    
    // Xóa file tạm ngay lập tức
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    
    return {
      fileName: 'BaoCao_' + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm") + '.xlsx',
      base64: base64
    };
  } catch (e) {
    throw new Error("Lỗi Server: " + e.toString());
  }
}

/* ===== EXPORT PDF ===== */
function exportFilteredPdf(filterObj){
  const params = {
    type: filterObj.type, mode: filterObj.mode, from: filterObj.dateFrom, to: filterObj.dateTo,
    year: filterObj.year, fromMonth: filterObj.monthFrom, toMonth: filterObj.monthTo, quarter: filterObj.quarter
  };
  const rows = processFilterLogic(params).reverse();
  
  const ss = SpreadsheetApp.create('temp_pdf_' + new Date().getTime());
  const sh = ss.getActiveSheet();
  
  sh.getRange(1,1).setValue('BÁO CÁO THU - CHI').setFontWeight('bold').setFontSize(16);
  sh.getRange(2,1).setValue('Ngày xuất: ' + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"));
  
  let thu=0, chi=0;
  rows.forEach(r => {
    if(r.type === 'Thu') thu += r.amount;
    else chi += r.amount;
  });

  sh.getRange(4,1).setValue('Tổng Thu:'); sh.getRange(4,2).setValue(thu);
  sh.getRange(5,1).setValue('Tổng Chi:'); sh.getRange(5,2).setValue(chi);
  sh.getRange(6,1).setValue('Tồn quỹ:'); sh.getRange(6,2).setValue(thu - chi);
  sh.getRange(4,1,3,1).setFontWeight('bold');

  const startRow = 8;
  sh.getRange(startRow,1,1,4).setValues([['Ngày','Loại','Nội dung','Số tiền']])
    .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');

  let rIdx = startRow + 1;
  rows.forEach(item => {
    sh.getRange(rIdx,1).setValue(item.date);
    sh.getRange(rIdx,2).setValue(item.type);
    sh.getRange(rIdx,3).setValue(item.content);
    sh.getRange(rIdx,4).setValue(item.amount);
    if(item.type === 'Chi') sh.getRange(rIdx,1,1,4).setBackground('#fff0f0').setFontColor('#b91c1c');
    rIdx++;
  });

  sh.autoResizeColumns(1,3);
  SpreadsheetApp.flush();

  const tempFile = DriveApp.getFileById(ss.getId());
  const pdfBlob = tempFile.getAs('application/pdf').setName('BaoCao_' + new Date().getTime() + '.pdf');
  const outFile = DriveApp.createFile(pdfBlob);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tempFile.setTrashed(true);
  return outFile.getUrl();
}

/* ===== CRUD FUNCTIONS ===== */
function getItemById(id){
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]) === String(id)){
      return {
        id: String(data[i][0]),
        date: Utilities.formatDate(new Date(data[i][1]), "GMT+7", 'yyyy-MM-dd'),
        type: data[i][2],
        content: data[i][3],
        amount: Number(data[i][4] || 0)
      };
    }
  }
  return null;
}

function addRow(d){
  const sh = getSheet();
  const jsDate = parseToDate(d.date);
  const id = String(new Date().getTime());
  sh.appendRow([ id, jsDate, d.type||'', d.content||'', Number(d.amount||0), new Date() ]);
  return {ok:true, id};
}

function updateRow(d){
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  const jsDate = parseToDate(d.date);
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(d.id)){
      sh.getRange(i+1,2).setValue(jsDate);
      sh.getRange(i+1,3).setValue(d.type||'');
      sh.getRange(i+1,4).setValue(d.content||'');
      sh.getRange(i+1,5).setValue(Number(d.amount||0));
      return {ok:true};
    }
  }
  return {ok:false};
}

function deleteRow(id){
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(id)){
      sh.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

function summary(){
  const data = getSheet().getDataRange().getValues();
  let thu=0, chi=0;
  for(let i=1;i<data.length;i++){
    const amt = Number(data[i][4]||0);
    if(data[i][2]==='Thu') thu+=amt;
    else if(data[i][2]==='Chi') chi+=amt;
  }
  return {thu, chi, ton: thu-chi};
}

function parseToDate(s){
  if(!s) return new Date();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)){
    const p = s.split('-').map(Number);
    return new Date(p[0], p[1]-1, p[2]);
  }
  if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)){
    const p = s.split('/').map(Number);
    return new Date(p[2], p[1]-1, p[0]);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? new Date() : d;
}