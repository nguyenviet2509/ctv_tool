// Biến lưu trữ dữ liệu
let masterData = null; // Dữ liệu file mẫu
let masterWorkbook = null; // Workbook gốc
let masterHeaderRowIndex = 0; // Chỉ số dòng tiêu đề
let filesProcessed = 0;
let newCTVCount = 0;
let updatedCTVCount = 0;
let uploadedFiles = []; // Danh sách file đã upload

// Các phần tử DOM
const templateFileInput = document.getElementById('templateFile');
const monthlyFileInput = document.getElementById('monthlyFile');
const templateStatus = document.getElementById('templateStatus');
const monthlyStatus = document.getElementById('monthlyStatus');
const exportBtn = document.getElementById('exportBtn');
const resetBtn = document.getElementById('resetBtn');
const dataTableBody = document.getElementById('dataTableBody');

// Các phần tử thông tin
const totalCTVElement = document.getElementById('totalCTV');
const filesProcessedElement = document.getElementById('filesProcessed');
const newCTVElement = document.getElementById('newCTV');
const updatedCTVElement = document.getElementById('updatedCTV');

// Các cột quan trọng (index bắt đầu từ 0)
const COLUMNS = {
    STT: 0,        // Cột A
    TEN: 1,        // Cột B
    SDT: 2,        // Cột C
    CCCD: 5,       // Cột F (index 5)
    HOA_HONG: 8,   // Cột I (index 8)
    THUE: 9,       // Cột J (index 9)
    TIEN_TRA: 10,  // Cột K (index 10)
    HOA_HONG_DA_CAP: 11,    // Cột L (index 11) - HH đã cấp
    THUE_DA_CAP: 12,        // Cột M (index 12) - Thuế TNCN đã cấp
    TIEN_TRA_DA_CAP: 13,    // Cột N (index 13) - Thực trả đã cấp
    HOA_HONG_CAN_CAP: 14,   // Cột O (index 14) - HH chưa cấp
    THUE_CAN_CAP: 15,       // Cột P (index 15) - Thuế TNCN chưa cấp
    TIEN_TRA_CAN_CAP: 16    // Cột Q (index 16) - Thực trả chưa cấp
};

// Biến lưu thông tin tháng của file hiện tại
let currentMonthInfo = null;

// Xử lý upload file mẫu
templateFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        showStatus(templateStatus, 'Đang đọc file mẫu...', 'info');
        
        const data = await readExcelFile(file);
        masterWorkbook = data.workbook;
        masterData = data.rows;
        
        // Tìm dòng tiêu đề (dòng có chứa "Tên" hoặc "CMND")
        if (masterData && masterData.length > 0) {
            let headerRowIndex = 0;
            for (let i = 0; i < Math.min(5, masterData.length); i++) {
                const row = masterData[i];
                const rowStr = row.join('').toLowerCase();
                if (rowStr.includes('tên') || rowStr.includes('cmnd') || rowStr.includes('cccd')) {
                    headerRowIndex = i;
                    break;
                }
            }
            masterHeaderRowIndex = headerRowIndex; // lưu toàn cục
            
            // Thêm tiêu đề cho 6 cột (3 cột cũ + 3 cột mới)
            const headerRow = masterData[headerRowIndex];
            setCellValue(headerRow, COLUMNS.HOA_HONG, 'Tổng HH Năm');
            setCellValue(headerRow, COLUMNS.THUE, 'Tổng Thuế TNCN Năm');
            setCellValue(headerRow, COLUMNS.TIEN_TRA, 'Tổng Thực Trả Năm');
            setCellValue(headerRow, COLUMNS.HOA_HONG_DA_CAP, 'HH Đã Cấp');
            setCellValue(headerRow, COLUMNS.THUE_DA_CAP, 'Thuế TNCN Đã Cấp');
            setCellValue(headerRow, COLUMNS.TIEN_TRA_DA_CAP, 'Thực Trả Đã Cấp');
            setCellValue(headerRow, COLUMNS.HOA_HONG_CAN_CAP, 'HH Chưa Cấp');
            setCellValue(headerRow, COLUMNS.THUE_CAN_CAP, 'Thuế TNCN Chưa Cấp');
            setCellValue(headerRow, COLUMNS.TIEN_TRA_CAN_CAP, 'Thực Trả Chưa Cấp');
            
            // Khởi tạo giá trị cho 6 cột mới dựa trên điều kiện Thuế TNCN
            for (let i = headerRowIndex + 1; i < masterData.length; i++) {
                const row = masterData[i];
                const hhVal  = getCellValue(row, COLUMNS.HOA_HONG);
                const taxVal = getCellValue(row, COLUMNS.THUE);
                const traVal = getCellValue(row, COLUMNS.TIEN_TRA);
                
                // Check xem Thuế TNCN có dữ liệu hay không
                if (hasThueData(row)) {
                    // Có thuế: đã cấp = tổng năm, chưa cấp = 0
                    setCellValue(row, COLUMNS.HOA_HONG_DA_CAP,  hhVal);
                    setCellValue(row, COLUMNS.THUE_DA_CAP,       taxVal);
                    setCellValue(row, COLUMNS.TIEN_TRA_DA_CAP,   traVal);
                    setCellValue(row, COLUMNS.HOA_HONG_CAN_CAP,  '0');
                    setCellValue(row, COLUMNS.THUE_CAN_CAP,       '0');
                    setCellValue(row, COLUMNS.TIEN_TRA_CAN_CAP,   '0');
                } else {
                    // Không có thuế: đã cấp = 0, chưa cấp = tổng năm
                    setCellValue(row, COLUMNS.HOA_HONG_DA_CAP,  '0');
                    setCellValue(row, COLUMNS.THUE_DA_CAP,       '0');
                    setCellValue(row, COLUMNS.TIEN_TRA_DA_CAP,   '0');
                    setCellValue(row, COLUMNS.HOA_HONG_CAN_CAP,  hhVal  || '0');
                    setCellValue(row, COLUMNS.THUE_CAN_CAP,       taxVal || '0');
                    setCellValue(row, COLUMNS.TIEN_TRA_CAN_CAP,   traVal || '0');
                }
            }
        }
        
        filesProcessed = 1;
        updateUI();
        renderTable();
        saveToLocalStorage();
        
        showStatus(templateStatus, `✓ Đã tải file mẫu thành công! (${masterData.length} CTV)`, 'success');
        exportBtn.disabled = false;
        
        // Khóa không cho upload file mẫu lần 2
        templateFileInput.disabled = true;
        templateFileInput.style.opacity = '0.5';
        templateFileInput.style.cursor = 'not-allowed';
        
    } catch (error) {
        showStatus(templateStatus, `✗ Lỗi: ${error.message}`, 'error');
        console.error(error);
    }
});

// Xử lý upload file tháng tiếp theo
monthlyFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!masterData) {
        showStatus(monthlyStatus, '✗ Vui lòng upload file mẫu trước!', 'error');
        monthlyFileInput.value = '';
        return;
    }

    // Validate tên file
    const monthInfo = extractMonthFromFileName(file.name);
    if (!monthInfo.isValid) {
        monthlyStatus.innerHTML = showMonthValidationError(file.name);
        monthlyStatus.className = 'status-message error';
        monthlyStatus.style.display = 'block';
        monthlyFileInput.value = '';
        return;
    }

    // Lưu thông tin tháng
    currentMonthInfo = monthInfo;

    try {
        showStatus(monthlyStatus, `Đang xử lý file tháng ${monthInfo.month}...`, 'info');
        
        const data = await readExcelFile(file);
        const monthlyRows = data.rows;
        
        // Tạo hash từ dữ liệu thực tế (danh sách CCCD) để phát hiện file trùng
        const cccdList = monthlyRows
            .map(row => getCellValue(row, COLUMNS.CCCD))
            .filter(cccd => cccd)
            .sort()
            .join('|');
        const fileHash = await simpleHash(cccdList);
        
        // Kiểm tra file đã upload chưa
        if (uploadedFiles.includes(fileHash)) {
            showStatus(monthlyStatus, '⚠️ File này đã được upload rồi! Không thể upload trùng để tránh cộng dồn sai số liệu.', 'error');
            monthlyFileInput.value = '';
            return;
        }
        
        // Reset bộ đếm cho file mới
        let currentNewCTV = 0;
        let currentUpdatedCTV = 0;
        let updatedCTVList = [];
        let newCTVList = [];
        
        // Xử lý từng dòng trong file tháng mới
        for (let i = 0; i < monthlyRows.length; i++) {
            const monthlyRow = monthlyRows[i];
            const cccd = getCellValue(monthlyRow, COLUMNS.CCCD);
            
            if (!cccd) continue; // Bỏ qua nếu không có CCCD
            
            // Check điều kiện: Cột Thuế TNCN có dữ liệu hay không?
            const hasThue = hasThueData(monthlyRow);
            
            // Tìm CCCD trong masterData
            const existingIndex = masterData.findIndex(row => 
                getCellValue(row, COLUMNS.CCCD) === cccd
            );
            
            if (existingIndex !== -1) {
                // CCCD đã tồn tại - Cộng dồn các giá trị
                const masterRow = masterData[existingIndex];
                
                // Cộng dồn Hoa hồng
                const oldHoaHong = parseNumber(getCellValue(masterRow, COLUMNS.HOA_HONG));
                const newHoaHong = parseNumber(getCellValue(monthlyRow, COLUMNS.HOA_HONG));
                const totalHoaHong = oldHoaHong + newHoaHong;
                setCellValue(masterRow, COLUMNS.HOA_HONG, totalHoaHong > 0 ? formatCurrencyForExcel(totalHoaHong) : '');
                
                // Cộng dồn Thuế
                const oldThue = parseNumber(getCellValue(masterRow, COLUMNS.THUE));
                const newThue = parseNumber(getCellValue(monthlyRow, COLUMNS.THUE));
                const totalThue = oldThue + newThue;
                setCellValue(masterRow, COLUMNS.THUE, totalThue > 0 ? formatCurrencyForExcel(totalThue) : '');
                
                // Cộng dồn Tiền trả
                const oldTienTra = parseNumber(getCellValue(masterRow, COLUMNS.TIEN_TRA));
                const newTienTra = parseNumber(getCellValue(monthlyRow, COLUMNS.TIEN_TRA));
                const totalTienTra = oldTienTra + newTienTra;
                setCellValue(masterRow, COLUMNS.TIEN_TRA, totalTienTra > 0 ? formatCurrencyForExcel(totalTienTra) : '');
                
                // Tính cột Đã Cấp và Chưa Cấp
                const prevDaCapHH    = parseNumber(getCellValue(masterRow, COLUMNS.HOA_HONG_DA_CAP));
                const prevDaCapThue  = parseNumber(getCellValue(masterRow, COLUMNS.THUE_DA_CAP));
                const prevDaCapTra   = parseNumber(getCellValue(masterRow, COLUMNS.TIEN_TRA_DA_CAP));
                
                let newDaCapHH   = prevDaCapHH;
                let newDaCapThue = prevDaCapThue;
                let newDaCapTra  = prevDaCapTra;
                
                if (hasThue) {
                    // Có thuế: đã cấp += giá trị tháng hiện tại
                    newDaCapHH   = prevDaCapHH   + newHoaHong;
                    newDaCapThue = prevDaCapThue  + newThue;
                    newDaCapTra  = prevDaCapTra   + newTienTra;
                }
                // Không có thuế: đã cấp giữ nguyên (cộng thêm 0)
                
                setCellValue(masterRow, COLUMNS.HOA_HONG_DA_CAP,   newDaCapHH   > 0 ? formatCurrencyForExcel(newDaCapHH)   : '0');
                setCellValue(masterRow, COLUMNS.THUE_DA_CAP,        newDaCapThue > 0 ? formatCurrencyForExcel(newDaCapThue) : '0');
                setCellValue(masterRow, COLUMNS.TIEN_TRA_DA_CAP,    newDaCapTra  > 0 ? formatCurrencyForExcel(newDaCapTra)  : '0');
                
                // Chưa cấp = Tổng năm - Đã cấp
                const chuaCapHH   = totalHoaHong - newDaCapHH;
                const chuaCapThue = totalThue    - newDaCapThue;
                const chuaCapTra  = totalTienTra - newDaCapTra;
                
                setCellValue(masterRow, COLUMNS.HOA_HONG_CAN_CAP,  chuaCapHH   > 0 ? formatCurrencyForExcel(chuaCapHH)   : '0');
                setCellValue(masterRow, COLUMNS.THUE_CAN_CAP,       chuaCapThue > 0 ? formatCurrencyForExcel(chuaCapThue) : '0');
                setCellValue(masterRow, COLUMNS.TIEN_TRA_CAN_CAP,   chuaCapTra  > 0 ? formatCurrencyForExcel(chuaCapTra)  : '0');
                
                currentUpdatedCTV++;
                updatedCTVList.push({
                    ten: getCellValue(monthlyRow, COLUMNS.TEN),
                    cccd: cccd,
                    hoaHong: formatCurrency(newHoaHong),
                    thue: formatCurrency(newThue),
                    tienTra: formatCurrency(newTienTra),
                    hasThue: hasThue
                });
            } else {
                // CCCD mới - Thêm hàng mới
                const newRow = [...monthlyRow];
                
                const newHH  = parseNumber(getCellValue(newRow, COLUMNS.HOA_HONG));
                const newTax = parseNumber(getCellValue(newRow, COLUMNS.THUE));
                const newTra = parseNumber(getCellValue(newRow, COLUMNS.TIEN_TRA));
                
                if (hasThue) {
                    // Có thuế: đã cấp = giá trị tháng này, chưa cấp = 0
                    setCellValue(newRow, COLUMNS.HOA_HONG_DA_CAP,  newHH  > 0 ? formatCurrencyForExcel(newHH)  : '0');
                    setCellValue(newRow, COLUMNS.THUE_DA_CAP,       newTax > 0 ? formatCurrencyForExcel(newTax) : '0');
                    setCellValue(newRow, COLUMNS.TIEN_TRA_DA_CAP,   newTra > 0 ? formatCurrencyForExcel(newTra) : '0');
                    setCellValue(newRow, COLUMNS.HOA_HONG_CAN_CAP,  '0');
                    setCellValue(newRow, COLUMNS.THUE_CAN_CAP,       '0');
                    setCellValue(newRow, COLUMNS.TIEN_TRA_CAN_CAP,   '0');
                } else {
                    // Không có thuế: đã cấp = 0, chưa cấp = tổng
                    setCellValue(newRow, COLUMNS.HOA_HONG_DA_CAP,  '0');
                    setCellValue(newRow, COLUMNS.THUE_DA_CAP,       '0');
                    setCellValue(newRow, COLUMNS.TIEN_TRA_DA_CAP,   '0');
                    setCellValue(newRow, COLUMNS.HOA_HONG_CAN_CAP,  newHH  > 0 ? formatCurrencyForExcel(newHH)  : '0');
                    setCellValue(newRow, COLUMNS.THUE_CAN_CAP,       newTax > 0 ? formatCurrencyForExcel(newTax) : '0');
                    setCellValue(newRow, COLUMNS.TIEN_TRA_CAN_CAP,   newTra > 0 ? formatCurrencyForExcel(newTra) : '0');
                }
                
                masterData.push(newRow);
                currentNewCTV++;
                newCTVList.push({
                    ten: getCellValue(monthlyRow, COLUMNS.TEN),
                    cccd: cccd,
                    hoaHong: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.HOA_HONG))),
                    thue: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.THUE))),
                    tienTra: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.TIEN_TRA))),
                    hasThue: hasThue
                });
            }
        }
        
        // Cập nhật STT
        updateSTT();
        
        // Cập nhật bộ đếm tổng
        newCTVCount += currentNewCTV;
        updatedCTVCount += currentUpdatedCTV;
        filesProcessed++;
        
        updateUI();
        renderTable();
        
        // Lưu hash của file vào danh sách đã upload
        uploadedFiles.push(fileHash);
        saveToLocalStorage();
        
        // Hiển thị bảng chi tiết các CTV đã cập nhật và thêm mới
        let detailHtml = `<div style="margin-top:15px; padding:15px; background:#f8f9fa; border-radius:8px;">`;
        detailHtml += `<div style="margin-bottom:10px; font-size:16px; font-weight:bold; color:#28a745;">✓ Xử lý thành công! (Tháng ${monthInfo.month})</div>`;
        
        // Bảng CTV cập nhật
        if (updatedCTVList.length > 0) {
            detailHtml += `
                <div style="margin-bottom:15px;">
                    <h4 style="color:#007bff; margin-bottom:10px;">🔄 CTV Đã Cập Nhật (${currentUpdatedCTV})</h4>
                    <div style="max-height:300px; overflow-y:auto; border:1px solid #dee2e6; border-radius:5px;">
                        <table style="width:100%; border-collapse:collapse; background:white; font-size:13px;">
                            <thead style="background:#007bff; color:white; position:sticky; top:0;">
                                <tr>
                                    <th style="padding:8px; text-align:left; border:1px solid #dee2e6;">Tên</th>
                                    <th style="padding:8px; text-align:left; border:1px solid #dee2e6;">CCCD/ID</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Tiền HH</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Thuế</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Tiền Trả</th>
                                </tr>
                            </thead>
                            <tbody>`;
            
            updatedCTVList.forEach(ctv => {
                detailHtml += `
                    <tr style="border-bottom:1px solid #dee2e6;">
                        <td style="padding:6px 8px; border:1px solid #dee2e6;">${ctv.ten}</td>
                        <td style="padding:6px 8px; border:1px solid #dee2e6;">${ctv.cccd}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.hoaHong}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.thue}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.tienTra}</td>
                    </tr>`;
            });
            
            detailHtml += `</tbody></table></div></div>`;
        }
        
        // Bảng CTV mới thêm
        if (newCTVList.length > 0) {
            detailHtml += `
                <div>
                    <h4 style="color:#28a745; margin-bottom:10px;">➕ CTV Mới Thêm (${currentNewCTV})</h4>
                    <div style="max-height:300px; overflow-y:auto; border:1px solid #dee2e6; border-radius:5px;">
                        <table style="width:100%; border-collapse:collapse; background:white; font-size:13px;">
                            <thead style="background:#28a745; color:white; position:sticky; top:0;">
                                <tr>
                                    <th style="padding:8px; text-align:left; border:1px solid #dee2e6;">Tên</th>
                                    <th style="padding:8px; text-align:left; border:1px solid #dee2e6;">CCCD/ID</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Tiền HH</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Thuế</th>
                                    <th style="padding:8px; text-align:right; border:1px solid #dee2e6;">Tiền Trả</th>
                                </tr>
                            </thead>
                            <tbody>`;
            
            newCTVList.forEach(ctv => {
                detailHtml += `
                    <tr style="border-bottom:1px solid #dee2e6;">
                        <td style="padding:6px 8px; border:1px solid #dee2e6;">${ctv.ten}</td>
                        <td style="padding:6px 8px; border:1px solid #dee2e6;">${ctv.cccd}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.hoaHong}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.thue}</td>
                        <td style="padding:6px 8px; text-align:right; border:1px solid #dee2e6;">${ctv.tienTra}</td>
                    </tr>`;
            });
            
            detailHtml += `</tbody></table></div></div>`;
        }
        
        detailHtml += `</div>`;
        
        monthlyStatus.innerHTML = detailHtml;
        monthlyStatus.className = 'status-message success';
        monthlyStatus.style.display = 'block';
        
        // Reset input để có thể upload file tiếp theo
        monthlyFileInput.value = '';
        
    } catch (error) {
        showStatus(monthlyStatus, `✗ Lỗi: ${error.message}`, 'error');
        console.error(error);
    }
});

// Xử lý xuất file Excel
exportBtn.addEventListener('click', () => {
    if (!masterData || masterData.length === 0) {
        alert('Không có dữ liệu để xuất!');
        return;
    }

    try {
        // Đảm bảo dòng tiêu đề luôn có nhãn đúng trước khi xuất
        if (masterData[masterHeaderRowIndex]) {
            const hr = masterData[masterHeaderRowIndex];
            setCellValue(hr, COLUMNS.HOA_HONG,         'Tổng HH Năm');
            setCellValue(hr, COLUMNS.THUE,              'Tổng Thuế TNCN Năm');
            setCellValue(hr, COLUMNS.TIEN_TRA,          'Tổng Thực Trả Năm');
            setCellValue(hr, COLUMNS.HOA_HONG_DA_CAP,   'HH Đã Cấp');
            setCellValue(hr, COLUMNS.THUE_DA_CAP,        'Thuế TNCN Đã Cấp');
            setCellValue(hr, COLUMNS.TIEN_TRA_DA_CAP,    'Thực Trả Đã Cấp');
            setCellValue(hr, COLUMNS.HOA_HONG_CAN_CAP,   'HH Chưa Cấp');
            setCellValue(hr, COLUMNS.THUE_CAN_CAP,        'Thuế TNCN Chưa Cấp');
            setCellValue(hr, COLUMNS.TIEN_TRA_CAN_CAP,    'Thực Trả Chưa Cấp');
        }

        // Các cột số tiền cần xuất dạng number
        const CURRENCY_COLS = [
            COLUMNS.HOA_HONG,
            COLUMNS.THUE,
            COLUMNS.TIEN_TRA,
            COLUMNS.HOA_HONG_DA_CAP,
            COLUMNS.THUE_DA_CAP,
            COLUMNS.TIEN_TRA_DA_CAP,
            COLUMNS.HOA_HONG_CAN_CAP,
            COLUMNS.THUE_CAN_CAP,
            COLUMNS.TIEN_TRA_CAN_CAP
        ];

        // Tạo bản sao dữ liệu, chuyển cột tiền sang number
        const exportData = masterData.map(row => {
            const newRow = [...row];
            CURRENCY_COLS.forEach(col => {
                const raw = newRow[col];
                if (raw === undefined || raw === null || raw === '') return;
                const num = parseNumber(raw);
                // Chỉ chuyển sang number nếu thực sự là số
                // (tránh overwrite tiêu đề cột bị parseNumber trả về 0)
                if (num > 0 || raw === 0 || raw === '0') {
                    newRow[col] = num;
                }
            });
            return newRow;
        });

        // Tạo worksheet từ dữ liệu đã chuẩn hoá
        const ws = XLSX.utils.aoa_to_sheet(exportData);

        // Áp định dạng #,##0 cho tất cả ô thuộc cột tiền (bỏ qua dòng tiêu đề)
        const range = XLSX.utils.decode_range(ws['!ref']);
        const currencyFmt = '#,##0';
        const currencyColSet = new Set(CURRENCY_COLS);

        for (let R = range.s.r + 1; R <= range.e.r; R++) {
            CURRENCY_COLS.forEach(C => {
                const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
                if (ws[cellAddr] && ws[cellAddr].t === 'n') {
                    ws[cellAddr].z = currencyFmt;
                }
            });
        }

        // Tính độ rộng cột: lấy độ dài ký tự lớn nhất trong mỗi cột
        const numCols = range.e.c + 1;
        const colWidths = Array(numCols).fill(0);

        for (let R = range.s.r; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellAddr];
                if (!cell) continue;
                let displayLen;
                if (currencyColSet.has(C) && cell.t === 'n') {
                    // Hiển thị với dấu phẩy để ước tính độ rộng
                    displayLen = Math.floor(cell.v).toLocaleString('en-US').length;
                } else {
                    displayLen = (cell.v !== undefined && cell.v !== null)
                        ? cell.v.toString().length : 0;
                }
                if (displayLen > colWidths[C]) colWidths[C] = displayLen;
            }
        }

        ws['!cols'] = colWidths.map(w => ({ wch: Math.min(w + 2, 60) }));

        // Lấy tên sheet từ workbook gốc hoặc dùng mặc định
        const sheetName = masterWorkbook ? 
            masterWorkbook.SheetNames[0] : 'Sheet1';
        
        // Tạo workbook mới
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        
        // Tạo tên file với timestamp
        const date = new Date();
        const timestamp = `${date.getFullYear()}${(date.getMonth()+1).toString().padStart(2,'0')}${date.getDate().toString().padStart(2,'0')}`;
        const fileName = `DuLieu_CTV_TongHop_${timestamp}.xlsx`;
        
        // Xuất file
        XLSX.writeFile(wb, fileName);
        
        showStatus(monthlyStatus, `✓ Đã xuất file: ${fileName}`, 'success');
        
    } catch (error) {
        alert(`Lỗi khi xuất file: ${error.message}`);
        console.error(error);
    }
});

// Xử lý reset
resetBtn.addEventListener('click', () => {
    if (confirm('Bạn có chắc muốn bắt đầu lại? Tất cả dữ liệu sẽ bị xóa.')) {
        masterData = null;
        masterWorkbook = null;
        filesProcessed = 0;
        newCTVCount = 0;
        updatedCTVCount = 0;
        uploadedFiles = [];
        currentMonthInfo = null;
        
        templateFileInput.value = '';
        monthlyFileInput.value = '';
        
        templateStatus.style.display = 'none';
        monthlyStatus.style.display = 'none';
        
        exportBtn.disabled = true;
        
        localStorage.removeItem('tool_ctv_data');
        
        // Mở khóa input file mẫu
        templateFileInput.disabled = false;
        templateFileInput.style.opacity = '1';
        templateFileInput.style.cursor = 'pointer';
        
        updateUI();
        renderTable();
    }
});

// Hàm đọc file Excel
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Lấy sheet đầu tiên
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Chuyển thành array of arrays
                const rows = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '',
                    raw: false
                });
                
                // Lọc bỏ các dòng hoàn toàn trống
                const filteredRows = rows.filter(row => 
                    row.some(cell => cell !== null && cell !== undefined && cell !== '')
                );
                
                resolve({ workbook, rows: filteredRows });
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Không thể đọc file'));
        reader.readAsArrayBuffer(file);
    });
}

// Hàm tạo hash từ chuỗi (dựa trên dữ liệu thực tế)
async function simpleHash(str) {
    const encoder = new TextEncoder();
    const data = encoder.encode(str);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex;
}

// Hàm lấy giá trị cell
function getCellValue(row, columnIndex) {
    if (!row || !row[columnIndex]) return '';
    return row[columnIndex];
}

// Hàm set giá trị cell
function setCellValue(row, columnIndex, value) {
    if (!row) return;
    row[columnIndex] = value;
}

// Hàm kiểm tra xem cột Thuế TNCN có dữ liệu hay không
function hasThueData(row) {
    const thueValue = getCellValue(row, COLUMNS.THUE);
    
    // Nếu không có giá trị hoặc là dấu "-"
    if (!thueValue || thueValue === '-' || thueValue.toString().trim() === '') {
        return false;
    }
    
    // Parse giá trị và check xem có phải số > 0 không
    const thueNum = parseNumber(thueValue);
    return thueNum > 0;
}

// Hàm validate và extract số tháng từ tên file
function extractMonthFromFileName(fileName) {
    // Loại bỏ phần mở rộng file (.xlsx, .xls, v.v.)
    const nameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
    
    // Các pattern để match tháng: Thang1, Thang_1, T1, thang1, thang_01, Thang01, v.v.
    // Pattern: (Thang|T|thang|THANG)(_)?(\d{1,2})
    const patterns = [
        /(?:thang|t)[\s_-]?(\d{1,2})/i,  // Thang1, Thang_1, T1, thang 1, thang-1, v.v.
    ];
    
    for (let pattern of patterns) {
        const match = nameWithoutExt.match(pattern);
        if (match) {
            const monthNum = parseInt(match[1], 10);
            
            // Validate tháng phải từ 1-12
            if (monthNum >= 1 && monthNum <= 12) {
                return {
                    isValid: true,
                    month: monthNum,
                    monthFormatted: monthNum.toString().padStart(2, '0')
                };
            }
        }
    }
    
    return {
        isValid: false,
        month: null,
        monthFormatted: null
    };
}

// Hàm hiển thị lỗi validation tháng
function showMonthValidationError(fileName) {
    const errorMsg = `
        <div style="padding:15px; background:#fff3cd; border:1px solid #ffc107; border-radius:8px; color:#856404;">
            <strong>⚠️ Tên file không hợp lệ!</strong><br>
            <span style="font-size:0.95em;">Tên file phải chứa thông tin tháng theo một trong các dạng sau:</span>
            <ul style="margin:10px 0 0 20px; font-size:0.9em;">
                <li>Thang1, Thang_1, Thang-1, Thang 1</li>
                <li>thang1, thang_1, thang-1, thang 1</li>
                <li>T1, T_1, T-1, T 1</li>
                <li>Thang01, thang_01, T01, v.v.</li>
            </ul>
            <div style="margin-top:10px; font-size:0.9em;">
                <strong>File của bạn:</strong> ${fileName}
            </div>
        </div>
    `;
    return errorMsg;
}

// Hàm parse số (xử lý cả số có dấu phẩy, chấm)
function parseNumber(value) {
    if (typeof value === 'number') return value;
    if (!value) return 0;
    
    // Xóa tất cả dấu phẩy (ngăn cách hàng nghìn) và các ký tự không phải số
    const cleaned = value.toString().replace(/,/g, '').replace(/[^\d.-]/g, '');
    const parsed = parseFloat(cleaned);
    
    return isNaN(parsed) ? 0 : parsed;
}

// Hàm format số thành tiền (dùng dấu phẩy cho hàng nghìn)
function formatCurrency(value) {
    const num = parseNumber(value);
    return num.toLocaleString('en-US');
}

// Hàm format số tiền cho Excel (với dấu phẩy)
function formatCurrencyForExcel(value) {
    const num = parseNumber(value);
    if (num === 0) return 0; // Giữ nguyên 0 thay vì trả về chuỗi rỗng
    return num.toLocaleString('en-US');
}

// Hàm cập nhật STT
function updateSTT() {
    if (!masterData) return;
    
    for (let i = 0; i < masterData.length; i++) {
        setCellValue(masterData[i], COLUMNS.STT, i + 1);
    }
}

// Hàm cập nhật UI
function updateUI() {
    totalCTVElement.textContent = masterData ? masterData.length : 0;
    filesProcessedElement.textContent = filesProcessed;
    newCTVElement.textContent = newCTVCount;
    updatedCTVElement.textContent = updatedCTVCount;
}

// Hàm render bảng dữ liệu
function renderTable() {
    if (!masterData || masterData.length === 0) {
        dataTableBody.innerHTML = `
            <tr>
                <td colspan="10" class="empty-message">Chưa có dữ liệu. Vui lòng upload file mẫu.</td>
            </tr>
        `;
        return;
    }
    
    let html = '';
    
    // Bỏ qua dòng tiêu đề (dòng đầu tiên)
    for (let i = 2; i < masterData.length; i++) {
        const row = masterData[i];
        
        html += `
            <tr>
                <td>${getCellValue(row, COLUMNS.STT)}</td>
                <td>${getCellValue(row, COLUMNS.TEN)}</td>
                <td>${getCellValue(row, COLUMNS.SDT)}</td>
                <td>${getCellValue(row, COLUMNS.CCCD)}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.HOA_HONG))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.THUE))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.TIEN_TRA))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.HOA_HONG_DA_CAP))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.THUE_DA_CAP))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.TIEN_TRA_DA_CAP))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.HOA_HONG_CAN_CAP))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.THUE_CAN_CAP))}</td>
                <td>${formatCurrency(getCellValue(row, COLUMNS.TIEN_TRA_CAN_CAP))}</td>
            </tr>
        `;
    }
    
    dataTableBody.innerHTML = html;
}

// Hàm hiển thị status
function showStatus(element, message, type) {
    element.textContent = message;
    element.className = `status-message ${type}`;
    element.style.display = 'block';
}

// Hàm lưu dữ liệu vào localStorage
function saveToLocalStorage() {
    try {
        const dataToSave = {
            masterData: masterData,
            masterHeaderRowIndex: masterHeaderRowIndex,
            filesProcessed: filesProcessed,
            newCTVCount: newCTVCount,
            updatedCTVCount: updatedCTVCount,
            uploadedFiles: uploadedFiles,
            currentMonthInfo: currentMonthInfo
        };
        localStorage.setItem('tool_ctv_data', JSON.stringify(dataToSave));
    } catch (error) {
        console.error('Lỗi khi lưu dữ liệu:', error);
    }
}

// Hàm load dữ liệu từ localStorage
function loadFromLocalStorage() {
    try {
        const savedData = localStorage.getItem('tool_ctv_data');
        if (savedData) {
            const data = JSON.parse(savedData);
            masterData = data.masterData;
            masterHeaderRowIndex = data.masterHeaderRowIndex || 0;
            filesProcessed = data.filesProcessed || 0;
            newCTVCount = data.newCTVCount || 0;
            updatedCTVCount = data.updatedCTVCount || 0;
            uploadedFiles = data.uploadedFiles || [];
            currentMonthInfo = data.currentMonthInfo || null;
            
            if (masterData && masterData.length > 0) {
                exportBtn.disabled = false;
                showStatus(templateStatus, '✓ Đã khôi phục dữ liệu từ phiên làm việc trước', 'success');
                
                // Khóa input file mẫu vì đã có dữ liệu
                templateFileInput.disabled = true;
                templateFileInput.style.opacity = '0.5';
                templateFileInput.style.cursor = 'not-allowed';
            }
            
            updateUI();
            renderTable();
        }
    } catch (error) {
        console.error('Lỗi khi load dữ liệu:', error);
    }
}

// Khởi tạo
document.addEventListener('DOMContentLoaded', () => {
    loadFromLocalStorage();
    updateUI();
});
