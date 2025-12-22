/* Code.gs - Server side (Apps Script) */
const SS = SpreadsheetApp.getActive();

/* ===== WEB ===== */
function doGet(){
  init();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản lý Thu Chi')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include(f){ return HtmlService.createHtmlOutputFromFile(f).getContent(); }

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
function getSheet(){ init(); return SS.getSheetByName('Data'); }

/* ===== LOGIN ===== */
function login(u,p){
  init();
  const rows = SS.getSheetByName('Users').getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==u && rows[i][1]==p) return {ok:true,name:rows[i][2]};
  }
  return {ok:false};
}

/* ===== BASIC LIST =====
 returns array {id,date(type dd/MM/yyyy),type,content,amount}
*/
function listData(){
  init();
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  const out = [];
  for(let i=1;i<data.length;i++){
    if(!data[i] || data[i].length===0) continue;
    let cellDate = data[i][1];
    let dateStr='';
    try{ dateStr = Utilities.formatDate(new Date(cellDate), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    catch(e){ dateStr = cellDate?String(cellDate):''; }
    out.push({
      id: String(data[i][0]),
      date: dateStr,
      type: data[i][2],
      content: data[i][3],
      amount: Number(data[i][4]||0)
    });
  }
  return out.reverse();
}

/* ===== GET item by ID for edit modal ===== */
function getItemById(id){
  init();
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(id)){
      return {
        id:String(data[i][0]),
        date: Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        type:data[i][2],
        content:data[i][3],
        amount:Number(data[i][4]||0)
      };
    }
  }
  return null;
}

/* ===== ADD / UPDATE / DELETE ===== */
function addRow(d){
  init();
  const sh = getSheet();
  const jsDate = parseToDate(d.date);
  const id = String(new Date().getTime());
  sh.appendRow([ id, jsDate, d.type||'', d.content||'', Number(d.amount||0), new Date() ]);
  return {ok:true, id};
}

function updateRow(d){
  init();
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
  init();
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

/* ===== SUMMARY ===== */
function summary(){
  init();
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  let thu=0, chi=0;
  for(let i=1;i<data.length;i++){
    const amt = Number(data[i][4]||0);
    if(data[i][2]==='Thu') thu+=amt;
    if(data[i][2]==='Chi') chi+=amt;
  }
  return {thu, chi, ton: thu-chi};
}

/* ====== FILTERING ======
 params = {
   mode: 'date'|'month'|'quarter'|'year'|'range' ,
   type: 'Thu'|'Chi'|''(all),
   // for date/range: from:'yyyy-mm-dd', to:'yyyy-mm-dd'
   // for month: year:2024, fromMonth:1, toMonth:3
   // for quarter: year, quarter:1..4
   // for year: year
 }
 Returns array of rows like listData()
*/
function filterData(params){
  const all = listData().reverse(); // chronological
  if(!params || Object.keys(params).length===0) return all.reverse();
  const res = [];
  // convert params to fromDate, toDate
  let from = null, to = null;
  try{
    if(params.mode==='date' || params.mode==='range'){
      if(params.from) from = parseToDate(params.from);
      if(params.to) to = parseToDate(params.to);
    } else if(params.mode==='month'){
      const y = Number(params.year||new Date().getFullYear());
      const fm = Number(params.fromMonth||1);
      const tm = Number(params.toMonth||fm);
      from = new Date(y, fm-1, 1);
      to = new Date(y, tm, 0); // last day of toMonth
    } else if(params.mode==='quarter'){
      const y = Number(params.year||new Date().getFullYear());
      const q = Number(params.quarter||1);
      const fm = (q-1)*3 + 1;
      from = new Date(y, fm-1, 1);
      const tm = fm + 2;
      to = new Date(y, tm, 0);
    } else if(params.mode==='year'){
      const y = Number(params.year||new Date().getFullYear());
      from = new Date(y,0,1);
      to = new Date(y,11,31);
    } else { // default all
      from = null; to = null;
    }
  }catch(e){ from=null; to=null; }

  all.forEach(r=>{
    // r.date is dd/MM/yyyy
    let rDate = null;
    try{ const p = r.date.split('/'); rDate = new Date(Number(p[2]), Number(p[1])-1, Number(p[0])); }catch(e){ rDate=null; }
    let ok = true;
    if(params.type && params.type!=='') ok = (r.type === params.type);
    if(ok && from && rDate) ok = rDate >= from;
    if(ok && to && rDate) ok = rDate <= to;
    if(ok) res.push(r);
  });
  return res.reverse();
}

/* ===== EXPORT filtered data to Excel (xlsx) =====
 Returns public URL to download
*/
function exportFilteredExcel(params){
  const rows = filterData(params).reverse(); // chronological
  const tempSs = SpreadsheetApp.create('temp_export_'+(new Date().getTime()));
  const ts = tempSs.getActiveSheet();
  // write header
  ts.appendRow(['Ngày','Loại','Nội dung','Số tiền']);
  // write rows
  rows.forEach(r=>{
    ts.appendRow([ r.date, r.type, r.content, Number(r.amount||0) ]);
  });
  SpreadsheetApp.flush();
  const tempFile = DriveApp.getFileById(tempSs.getId());
  // convert to xlsx blob
  const xlsxBlob = tempFile.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').setName('BaoCao_ThuChi_'+Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx');
  const outFile = DriveApp.createFile(xlsxBlob);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // trash temp sheet file
  tempFile.setTrashed(true);
  return outFile.getUrl();
}

/* ===== EXPORT filtered to PDF (C1 style simple report) =====
 Returns public URL
*/
function exportFilteredPdf(params){
  const rows = filterData(params).reverse(); // chronological
  // Create temp sheet and write a styled simple report
  const ss = SpreadsheetApp.create('temp_pdf_'+(new Date().getTime()));
  const sh = ss.getActiveSheet();
  // Title
  sh.getRange(1,1).setValue('BÁO CÁO THU – CHI').setFontWeight('bold').setFontSize(16);
  sh.getRange(2,1).setValue('Ngày xuất: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'));
  // Summary
  const summaryObj = summaryFiltered(rows);
  sh.getRange(4,1).setValue('Tổng Thu:').setFontWeight('bold'); sh.getRange(4,2).setValue(summaryObj.thu);
  sh.getRange(5,1).setValue('Tổng Chi:').setFontWeight('bold'); sh.getRange(5,2).setValue(summaryObj.chi);
  sh.getRange(6,1).setValue('Tồn quỹ:').setFontWeight('bold'); sh.getRange(6,2).setValue(summaryObj.ton);
  // table header at row 8
  const startRow = 8;
  sh.getRange(startRow,1,1,4).setValues([['Ngày','Loại','Nội dung','Số tiền']]).setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
  // write rows
  let r = startRow + 1;
  rows.forEach(item=>{
    sh.getRange(r,1).setValue(item.date);
    sh.getRange(r,2).setValue(item.type);
    sh.getRange(r,3).setValue(item.content);
    sh.getRange(r,4).setValue(Number(item.amount||0));
    // style chi rows
    if(item.type === 'Chi'){
      sh.getRange(r,1,1,4).setBackground('#fff0f0').setFontColor('#b91c1c');
    }
    r++;
  });
  // auto resize columns
  sh.autoResizeColumns(1,4);
  SpreadsheetApp.flush();

  const tempFile = DriveApp.getFileById(ss.getId());
  const pdfBlob = tempFile.getAs('application/pdf').setName('BaoCao_ThuChi_'+Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.pdf');
  const outFile = DriveApp.createFile(pdfBlob);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // trash temp
  tempFile.setTrashed(true);
  return outFile.getUrl();
}

/* helper to compute summary for rows Array */
function summaryFiltered(rows){
  let thu=0, chi=0;
  rows.forEach(r=>{
    const amt = Number(r.amount||0);
    if(r.type==='Thu') thu+=amt;
    if(r.type==='Chi') chi+=amt;
  });
  return {thu, chi, ton: thu-chi};
}

/* ===== date parser: accept yyyy-mm-dd (input.date) or dd/MM/yyyy ===== */
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
  if(!isNaN(d.getTime())) return d;
  return new Date();
}
