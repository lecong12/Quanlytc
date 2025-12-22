/* =========================================================================
 * Code.gs - Hoàn chỉnh và Đồng bộ
 * Quản lý các hàm backend cho ứng dụng Thu Chi Gia đình (Web App)
 * ========================================================================= */
const SS = SpreadsheetApp.getActive();

/* -------------------------------------------------------------------------
 * PHẦN 1: KHỞI TẠO, CẤU TRÚC VÀ CÀI ĐẶT CHUNG (INIT, WEB, HELPERS)
 * Chứa các hàm khởi tạo sheet, cấu hình Web App và các tiện ích cơ bản.
 * ------------------------------------------------------------------------- */

/**
 * Khởi tạo sheet 'Data' và 'Users' nếu chưa tồn tại.
 */
function init(){
    // 1. Sheet Data (Dữ liệu giao dịch)
    let sh = SS.getSheetByName('Data');
    if(!sh){
        sh = SS.insertSheet('Data');
        sh.appendRow(['ID','Ngày','Loại','Nội dung','Số tiền','Tạo lúc']);
    }
    // 2. Sheet Users (Tài khoản đăng nhập)
    let us = SS.getSheetByName('Users');
    if(!us){
        us = SS.insertSheet('Users');
        us.appendRow(['Tài khoản','Mật khẩu','Tên']);
        us.appendRow(['admin','admin','Quản trị']);
        us.hideSheet();
    }
}

/**
 * Xử lý yêu cầu truy cập Web App (Web App Entry Point).
 */
function doGet(){
    init();
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('Thu Chi Gia đình')
        .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Hàm tiện ích để nhúng nội dung file HTML/CSS/JS khác vào Index.html.
 */
function include(f){ 
    return HtmlService.createHtmlOutputFromFile(f).getContent(); 
}

/**
 * Lấy sheet dữ liệu chính ('Data'), đảm bảo đã được init.
 */
function getSheet(){ 
    init(); 
    return SS.getSheetByName('Data'); 
}

/**
 * Hàm tiện ích chuyển đổi chuỗi ngày sang đối tượng Date.
 * Hỗ trợ định dạng yyyy-MM-dd và dd/MM/yyyy.
 */
function parseToDate(s){
    if(!s) return new Date();
    // Định dạng yyyy-MM-dd (từ input type="date" của JS)
    if(/^\d{4}-\d{2}-\d{2}$/.test(s)){
        const p = s.split('-').map(Number);
        return new Date(p[0], p[1]-1, p[2]);
    }
    // Định dạng dd/MM/yyyy (từ sheet)
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)){
        const p = s.split('/').map(Number);
        return new Date(p[2], p[1]-1, p[0]);
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? new Date() : d;
}

/* -------------------------------------------------------------------------
 * PHẦN 2: CHỨC NĂNG ĐĂNG NHẬP (AUTHENTICATION)
 * ------------------------------------------------------------------------- */

/**
 * Kiểm tra thông tin đăng nhập.
 */
function login(u,p){
    const rows = SS.getSheetByName('Users').getDataRange().getValues();
    for(let i=1;i<rows.length;i++){
        if(rows[i][0]==u && rows[i][1]==p) return {ok:true, name:rows[i][2]};
    }
    return {ok:false};
}

/* -------------------------------------------------------------------------
 * PHẦN 3: ĐỌC DỮ LIỆU & LỌC (READ & FILTERING LOGIC)
 * Chứa các hàm đọc dữ liệu cơ bản và áp dụng logic lọc nâng cao.
 * ------------------------------------------------------------------------- */

/**
 * Lấy toàn bộ dữ liệu từ sheet và định dạng lại thành mảng objects.
 */
function listData(){
    const sh = getSheet();
    const data = sh.getDataRange().getValues();
    const out = [];
    for(let i=1;i<data.length;i++){
        if(!data[i] || data[i].length===0) continue;
        if(!data[i][1]) continue; // Nếu cột Ngày (cột B) trống thì bỏ qua dòng này
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

/**
 * Áp dụng logic lọc dữ liệu dựa trên các tham số thời gian và loại.
 * @param {object} params - Đối tượng chứa tiêu chí lọc (mode, type, year, month, date,...)
 */
function processFilterLogic(params) {
    // 1. Lấy dữ liệu và đảo ngược
    const all = listData().reverse(); 
    
    // Nếu không có tham số, trả về toàn bộ
    if (!params || Object.keys(params).length === 0) return all;

    const res = [];
    let from = null, to = null;
    
    // 2. Thiết lập thời gian (Giữ nguyên logic cũ của bạn vì nó đang chạy tốt)
    try {
        if (params.mode === 'date' || params.mode === 'range') {
            if (params.from) from = parseToDate(params.from);
            if (params.to) to = parseToDate(params.to);
            if (params.mode === 'date' && params.from) to = parseToDate(params.from);
        } else if (params.mode === 'month') {
            const y = Number(params.year || new Date().getFullYear());
            const fm = Number(params.fromMonth || 1);
            const tm = Number(params.toMonth || fm);
            from = new Date(y, fm - 1, 1);
            to = new Date(y, tm, 0); 
        } else if (params.mode === 'quarter') {
            const y = Number(params.year || new Date().getFullYear());
            const q = Number(params.quarter || 1);
            const fm = (q - 1) * 3 + 1;
            from = new Date(y, fm - 1, 1);
            to = new Date(y, fm + 2, 0); 
        } else if (params.mode === 'year') {
            const y = Number(params.year || new Date().getFullYear());
            from = new Date(y, 0, 1);
            to = new Date(y, 11, 31);
        }
    } catch (e) { }

    // 3. Duyệt và lọc với logic "Lọc tầng"
    all.forEach(r => {
        // --- Bước A: Lọc theo Loại (Thu/Chi) ---
        // Nếu chọn "Tất cả" hoặc không chọn gì, passType mặc định là TRUE
        let passType = true;
        if (params.type && params.type !== "Tất cả" && params.type !== "") {
            passType = (r.type === params.type);
        }

        // --- Bước B: Lọc theo Thời gian ---
        let passDate = true;
        if (from || to) {
            try {
                const p = r.date.split('/');
                const rDate = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])).getTime();
                
                if (from && rDate < from.getTime()) passDate = false;
                if (to && rDate > to.getTime()) passDate = false;
            } catch (e) { passDate = false; }
        }

        // --- Bước C: Kết hợp ---
        // Chỉ khi thỏa mãn CẢ LOẠI và CẢ NGÀY thì mới lấy
        if (passType && passDate) {
            res.push(r);
        }
    });

    return res;
}


/**
 * Cầu nối giữa JavaScript và Server để lấy dữ liệu đã lọc.
 * @param {object} filterObj - Tiêu chí lọc từ client.
 */

function getDataFromServer(filterObj) {
    try {
        // Đảm bảo filterObj luôn tồn tại để không gây lỗi crash server
        const obj = filterObj || { mode: 'all' }; 
        const params = {
            type: obj.type || '',
            mode: obj.mode || 'all',
            from: obj.dateFrom || null,
            to: obj.dateTo || null,
            year: obj.year || new Date().getFullYear(),
            fromMonth: obj.monthFrom || 1,
            toMonth: obj.monthTo || 12,
            quarter: obj.quarter || 1
        };
        return processFilterLogic(params);
    } catch (e) {
        throw new Error("Lỗi Server: " + e.message);
    }
}

/**
 * Tính toán tổng thu, tổng chi và tồn quỹ.
 */
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

/* -------------------------------------------------------------------------
 * PHẦN 4: CHỨC NĂNG THÊM, SỬA, XÓA (CRUD)
 * ------------------------------------------------------------------------- */

/**
 * Lấy một dòng dữ liệu theo ID.
 */
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

/**
 * Thêm một dòng dữ liệu mới.
 */
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
    
    for(let i=1; i < data.length; i++){
        // Cột A là ID (index 0)
        if(String(data[i][0]) === String(d.id)){
            // 1. Ghi dữ liệu mới vào dòng
            sh.getRange(i+1, 2).setValue(jsDate); // Cột B: Ngày
            sh.getRange(i+1, 3).setValue(d.type || '');
            sh.getRange(i+1, 4).setValue(d.content || '');
            sh.getRange(i+1, 5).setValue(Number(d.amount || 0));
            
            // 2. Lệnh cực kỳ quan trọng: BUỘC SHEETS PHẢI SẮP XẾP LẠI
            // Lấy toàn bộ vùng dữ liệu trừ hàng tiêu đề
            const lastRow = sh.getLastRow();
            const lastCol = sh.getLastColumn();
            if (lastRow > 1) {
                const range = sh.getRange(2, 1, lastRow - 1, lastCol);
                // Sắp xếp cột 2 (Cột B - Ngày) theo chiều giảm dần (Mới nhất lên đầu)
                range.sort({column: 2, ascending: true});
            }
            
            return {ok: true};
        }
    }
    return {ok: true};
}



/**
 * Xóa một dòng dữ liệu theo ID.
 */
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

/* -------------------------------------------------------------------------
 * PHẦN 5: XUẤT DỮ LIỆU EXCEL & PDF (EXPORT)
 * Lưu ý: Các hàm export này chỉ mang tính tham khảo, tính năng export thực tế
 * đã được chuyển sang client side (JavaScript) để đơn giản hóa.
 * ------------------------------------------------------------------------- */

/**
 * Export dữ liệu đã lọc sang file Excel (.xlsx) dưới dạng Base64.
 * Lưu ý: Đây là hàm phức tạp, thường nên sử dụng thư viện trên client (JS/XLSX.js)
 */
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

        const url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + token }
        });
        
        const bytes = response.getContent();
        const base64 = Utilities.base64Encode(bytes);
        
        DriveApp.getFileById(tempSs.getId()).setTrashed(true);
        
        return {
            fileName: 'BaoCao_' + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm") + '.xlsx',
            base64: base64
        };
    } catch (e) {
        throw new Error("Lỗi Server: " + e.toString());
    }
}

/**
 * Export dữ liệu đã lọc sang file PDF bằng cách tạo một sheet tạm.
 * Lưu ý: Thường đã được thay thế bằng html2pdf.js trên client.
 */
function exportFilteredPdf(filterObj){
    const params = {
        type: filterObj.type, mode: filterObj.mode, from: filterObj.dateFrom, to: filterObj.dateTo,
        year: filterObj.year, fromMonth: filterObj.monthFrom, toMonth: filterObj.monthTo, quarter: filterObj.quarter
    };
    const rows = processFilterLogic(params).reverse();
    
    // TẠO SHEET TẠM
    const ss = SpreadsheetApp.create('temp_pdf_' + new Date().getTime());
    const sh = ss.getActiveSheet();
    
    // TIÊU ĐỀ & TÓM TẮT
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

    // DỮ LIỆU BẢNG
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

    // XUẤT PDF VÀ DỌN DẸP
    const tempFile = DriveApp.getFileById(ss.getId());
    const pdfBlob = tempFile.getAs('application/pdf').setName('BaoCao_' + new Date().getTime() + '.pdf');
    const outFile = DriveApp.createFile(pdfBlob);
    outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempFile.setTrashed(true);
    return outFile.getUrl();
}

/* ============================================================
    HÀM XỬ LÝ GỬI EMAIL TRÊN SERVER (Code.gs)
   ============================================================ */

function sendEmailFromApp(data, recipientEmail, summary) {
  try {
    // 1. Kiểm tra giá trị Tồn quỹ để quyết định màu sắc (Đỏ nếu âm, Xanh dương nếu dương)
    let tonValue = 0;
    if (summary.ton) {
      // Xóa dấu chấm và dấu phẩy để chuyển chuỗi tiền tệ thành số thuần túy
      tonValue = parseFloat(summary.ton.toString().replace(/\./g, '').replace(/,/g, '')) || 0;
    }
    
    // Màu đỏ (#d63031) cho số âm, màu xanh dương (#0984e3) cho số dương hoặc bằng 0
    let mauTonQuy = tonValue < 0 ? '#d63031' : '#0984e3';

    // 2. Khởi tạo cấu trúc bảng HTML
    let htmlTable = `
      <h2 style="color: #2d3436; text-align: center; text-transform: uppercase;">BÁO CÁO THU CHI GIA ĐÌNH</h2>
      <p style="text-align: center;">Ngày gửi: ${new Date().toLocaleString('vi-VN')}</p>
      <table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%; font-family: sans-serif;">
        <thead style="background-color: #2d3436; color: white;">
          <tr>
            <th style="width: 20%;">Ngày</th>
            <th style="width: 15%;">Loại</th>
            <th style="width: 40%;">Nội dung</th>
            <th style="width: 25%;">Số tiền</th>
          </tr>
        </thead>
        <tbody>
    `;

    // 3. Duyệt dữ liệu danh sách thu chi
    data.forEach(item => {
      const color = item.type === 'Chi' ? '#d63031' : '#27ae60';
      htmlTable += `
        <tr>
          <td>${item.date}</td>
          <td style="color: ${color}; font-weight: normal;">${item.type}</td>
          <td>${item.content}</td>
          <td style="text-align: right;">${item.amount}</td>
        </tr>
      `;
    });

    // 4. Chèn 3 dòng tổng cộng (Thu, Chi, Tồn) vào cuối bảng
    htmlTable += `
      <tr style="background-color: #f1f2f6; font-weight: bold;">
        <td colspan="3" style="text-align: right;">TỔNG THU:</td>
        <td style="text-align: right; color: #27ae60;">${summary.thu}</td>
      </tr>
      <tr style="background-color: #f1f2f6; font-weight: bold;">
        <td colspan="3" style="text-align: right;">TỔNG CHI:</td>
        <td style="text-align: right; color: #d63031;">${summary.chi}</td>
      </tr>
      <tr style="background-color: #dfe4ea; font-weight: bold; font-size: 1.1em;">
        <td colspan="3" style="text-align: right; text-transform: uppercase;">TỒN QUỸ:</td>
        <td style="text-align: right; color: ${mauTonQuy};">${summary.ton}</td>
      </tr>
    `;

    htmlTable += `</tbody></table><p style="text-align: center; color: #636e72;"><i>Báo cáo được gửi tự động từ App quản lý thu chi.</i></p>`;

    // 5. Thực hiện gửi Email qua GmailApp
    GmailApp.sendEmail(recipientEmail, "Báo cáo Thu Chi Gia Đình - " + new Date().toLocaleDateString('vi-VN'), "", {
      htmlBody: htmlTable
    });

    return { ok: true };
  } catch (e) {
    // Trả về lỗi nếu có sự cố (ví dụ: email sai định dạng, hết hạn mức gửi...)
    return { ok: false, error: e.toString() };
  }
}