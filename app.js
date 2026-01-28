// Biến lưu trữ dữ liệu
let masterData = null; // Dữ liệu file mẫu
let masterWorkbook = null; // Workbook gốc
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
    TIEN_TRA: 10   // Cột K (index 10)
};

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
            
            // Thêm tiêu đề cho 3 cột chính
            const headerRow = masterData[headerRowIndex];
            setCellValue(headerRow, COLUMNS.HOA_HONG, 'Tiền Hoa Hồng');
            setCellValue(headerRow, COLUMNS.THUE, 'Thuế TNCN');
            setCellValue(headerRow, COLUMNS.TIEN_TRA, 'Số Tiền Trả CTV');
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

    try {
        showStatus(monthlyStatus, 'Đang xử lý file...', 'info');
        
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
            
            // Tìm CCCD trong masterData
            const existingIndex = masterData.findIndex(row => 
                getCellValue(row, COLUMNS.CCCD) === cccd
            );
            
            if (existingIndex !== -1) {
                // Case 1: CCCD đã tồn tại - Cộng dồn các giá trị
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
                
                currentUpdatedCTV++;
                updatedCTVList.push({
                    ten: getCellValue(monthlyRow, COLUMNS.TEN),
                    cccd: cccd,
                    hoaHong: formatCurrency(newHoaHong),
                    thue: formatCurrency(newThue),
                    tienTra: formatCurrency(newTienTra)
                });
            } else {
                // Case 2: CCCD mới - Thêm hàng mới
                masterData.push([...monthlyRow]);
                currentNewCTV++;
                newCTVList.push({
                    ten: getCellValue(monthlyRow, COLUMNS.TEN),
                    cccd: cccd,
                    hoaHong: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.HOA_HONG))),
                    thue: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.THUE))),
                    tienTra: formatCurrency(parseNumber(getCellValue(monthlyRow, COLUMNS.TIEN_TRA)))
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
        detailHtml += `<div style="margin-bottom:10px; font-size:16px; font-weight:bold; color:#28a745;">✓ Xử lý thành công!</div>`;
        
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
        // Tạo worksheet từ dữ liệu
        const ws = XLSX.utils.aoa_to_sheet(masterData);
        
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
    if (!value || value === 0) return '';
    return value.toLocaleString('en-US');
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
                <td colspan="7" class="empty-message">Chưa có dữ liệu. Vui lòng upload file mẫu.</td>
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
            filesProcessed: filesProcessed,
            newCTVCount: newCTVCount,
            updatedCTVCount: updatedCTVCount,
            uploadedFiles: uploadedFiles
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
            filesProcessed = data.filesProcessed || 0;
            newCTVCount = data.newCTVCount || 0;
            updatedCTVCount = data.updatedCTVCount || 0;
            uploadedFiles = data.uploadedFiles || [];
            
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
