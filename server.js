const express = require("express");
const ExcelJS = require("exceljs");
const multer = require("multer");

const app = express();
const PORT = process.env.PORT || 3000;

// Multer lưu file tạm vào RAM
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 15 * 1024 * 1024 }
});

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Dữ liệu mặc định ban đầu
let excelHeaders = [
    "Company", "PartNuM", "PartDescription", "UOMClassID", "IUM", 
    "PUM", "SalesUM", "TypeCode", "ClassID", "ProdCode"
];

let excelData = [
    { Company: "AQT", PartNuM: "R01.01.001", PartDescription: "Aluminium (AL 6061 1000x1000x8 mm)", UOMClassID: "Count", IUM: "EA", PUM: "EA", SalesUM: "EA", TypeCode: "P", ClassID: "O301", ProdCode: "10.0" },
    { Company: "AQT", PartNuM: "R01.01.002", PartDescription: "Aluminium (AL 6061 600x600x8 mm)", UOMClassID: "Count", IUM: "EA", PUM: "EA", SalesUM: "EA", TypeCode: "P", ClassID: "O301", ProdCode: "10.0" },
    { Company: "AQT", PartNuM: "R01.01.024", PartDescription: "Aluminium (AL 6061 1200x1000x24 mm)", UOMClassID: "Count", IUM: "EA", PUM: "EA", SalesUM: "EA", TypeCode: "P", ClassID: "O301", ProdCode: "10.0" },
    { Company: "AQT", PartNuM: "R02.04.015", PartDescription: "Steel Plate 24mm Heavy Duty", UOMClassID: "Count", IUM: "EA", PUM: "EA", SalesUM: "EA", TypeCode: "M", ClassID: "S102", ProdCode: "12.0" }
];

// CSS ĐÃ SỬA CHUẨN GIAO DIỆN
const style = `
<style>
    * { box-sizing: border-box; }
    html, body {
        height: 100vh;
        width: 100vw;
        margin: 0;
        padding: 0;
        overflow: hidden; /* Cố định chuẩn 1 màn hình */
        font-family: Arial, sans-serif;
        background-color: #f4f5f7;
    }
    
    .wrapper {
        display: flex;
        flex-direction: column;
        height: 100vh;
        padding: 20px 30px;
    }

    /* Nút Import excel trên cùng bên phải */
    .top-actions {
        display: flex;
        justify-content: flex-end;
        margin-bottom: 15px;
        flex-shrink: 0;
    }

    /* Layout 2 cột co giãn theo chiều cao màn hình */
    .main-layout {
        display: grid;
        grid-template-columns: 280px 1fr;
        gap: 25px;
        flex: 1;
        min-height: 0;
    }

    /* Ô xanh bên trái: Căn giữa nội dung "Nhập sp cần mua" */
    .left-panel {
        background-color: #4a77d4;
        color: white;
        padding: 30px 20px;
        border-radius: 4px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
        height: 100%;
    }
    .left-panel label {
        font-size: 18px;
        font-weight: bold;
        margin-bottom: 20px;
        text-align: center;
    }
    .left-panel input {
        width: 100%;
        padding: 12px 14px;
        border: none;
        border-radius: 4px;
        font-size: 15px;
        outline: none;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }

    /* Cột bên phải */
    .right-section {
        display: flex;
        flex-direction: column;
        height: 100%;
        min-height: 0;
    }

    /* Khung xanh File gốc */
    .file-goc-panel {
        background-color: #4a77d4;
        color: white;
        padding: 20px;
        border-radius: 4px;
        display: flex;
        flex-direction: column;
        flex: 1;
        min-height: 0;
        position: relative;
    }
    
    /* Chữ "File gốc" căn chính giữa đẹp mắt */
    .file-goc-header {
        font-size: 22px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 15px;
        flex-shrink: 0;
    }
    
    /* Bảng dữ liệu tự tạo thanh cuộn nội bộ */
    .table-container {
        background: white;
        color: #333;
        flex: 1;
        border-radius: 4px;
        overflow: auto;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        white-space: nowrap;
    }
    th, td {
        padding: 10px 14px;
        border: 1px solid #e0e0e0;
        text-align: left;
    }
    th {
        background: #f8f9fa;
        color: #333;
        font-weight: bold;
        position: sticky;
        top: 0;
        z-index: 2;
    }
    tr:nth-child(even) { background-color: #fcfcfc; }
    tr:hover { background-color: #eaf2ff; }

    /* Nút Thêm sp mới phía dưới */
    .bottom-actions {
        display: flex;
        justify-content: flex-end;
        margin-top: 15px;
        flex-shrink: 0;
    }

    .btn {
        padding: 10px 22px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        background: #ffffff;
        color: #333;
        font-size: 14px;
        font-weight: 500;
        box-shadow: 0 1px 4px rgba(0,0,0,0.15);
    }
    .btn:hover { background: #e8e8e8; }
    .btn-primary { background: #4a77d4; color: white; }
    .btn-primary:hover { background: #3b63b8; }

    /* Modal Form Thêm SP mới */
    .modal {
        display: none;
        position: fixed;
        top: 0; left: 0; width: 100%; height: 100%;
        background: rgba(0,0,0,0.5);
        justify-content: center;
        align-items: center;
        z-index: 1000;
    }
    .modal-content {
        background: white;
        padding: 25px;
        border-radius: 6px;
        width: 450px;
        max-height: 85vh;
        overflow-y: auto;
    }
    .modal-content input {
        width: 100%;
        padding: 8px 10px;
        margin: 4px 0 12px;
        border: 1px solid #ccc;
        border-radius: 4px;
    }
</style>
`;

// API LỌC DỮ LIỆU REALTIME
app.get("/api/filter", (req, res) => {
    const q = String(req.query.q || "").toLowerCase().trim();
    
    if (!q) {
        return res.json(excelData);
    }

    const filtered = excelData.filter(row => {
        return Object.values(row).some(val => 
            String(val || "").toLowerCase().includes(q)
        );
    });

    res.json(filtered);
});

// TRANG CHÍNH
app.get("/", (req, res) => {
    const headerHtml = excelHeaders.map(h => `<th>${h}</th>`).join("");

    const rowsHtml = excelData.map(row => {
        const cells = excelHeaders.map(h => `<td>${row[h] !== undefined ? row[h] : ""}</td>`).join("");
        return `<tr>${cells}</tr>`;
    }).join("") || `<tr><td colspan="${excelHeaders.length || 1}" style="text-align:center">Chưa có dữ liệu</td></tr>`;

    const modalInputsHtml = excelHeaders.map(h => `
        <label style="font-size:12px; font-weight:bold;">${h}:</label>
        <input type="text" name="${h}" placeholder="Nhập ${h}">
    `).join("");

    res.send(`
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Quản lý Sản Phẩm Excel</title>
            ${style}
        </head>
        <body>
            <div class="wrapper">
                
                <!-- TOP ACTION: Import Excel -->
                <div class="top-actions">
                    <form id="importForm" action="/import-excel" method="POST" enctype="multipart/form-data">
                        <input type="file" id="excelFile" name="excelFile" accept=".xlsx" style="display:none" onchange="document.getElementById('importForm').submit()">
                        <button type="button" class="btn" onclick="document.getElementById('excelFile').click()">Import excel</button>
                    </form>
                </div>

                <!-- MAIN LAYOUT 2 CỘT -->
                <div class="main-layout">
                    
                    <!-- BÊN TRÁI: Ô NHẬP SP CẦN MUA -->
                    <div class="left-panel">
                        <label for="searchInput">Nhập sp cần mua</label>
                        <input 
                            type="text" 
                            id="searchInput" 
                            placeholder="Gõ bất kỳ (VD: 24, AL)..." 
                            autocomplete="off"
                        >
                    </div>

                    <!-- BÊN PHẢI: KHUNG FILE GỐC -->
                    <div class="right-section">
                        <div class="file-goc-panel">
                            <div class="file-goc-header">File gốc</div>
                            <div class="table-container">
                                <table>
                                    <thead>
                                        <tr id="tableHeader">
                                            ${headerHtml}
                                        </tr>
                                    </thead>
                                    <tbody id="tableBody">
                                        ${rowsHtml}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- BOTTOM ACTIONS: Thêm sp mới -->
                        <div class="bottom-actions">
                            <button type="button" class="btn" onclick="openModal()">Thêm sp mới</button>
                        </div>
                    </div>

                </div>
            </div>

            <!-- MODAL THÊM SP MỚI -->
            <div id="addModal" class="modal">
                <div class="modal-content">
                    <h3 style="margin-top:0; font-size:16px;">Thêm sản phẩm mới vào file gốc</h3>
                    <form action="/add-product" method="POST">
                        ${modalInputsHtml}
                        <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:15px;">
                            <button type="button" class="btn" onclick="closeModal()">Hủy</button>
                            <button type="submit" class="btn btn-primary">Lưu sản phẩm</button>
                        </div>
                    </form>
                </div>
            </div>

            <!-- JAVASCRIPT LỌC DỮ LIỆU REALTIME -->
            <script>
                const searchInput = document.getElementById("searchInput");
                const tableBody = document.getElementById("tableBody");
                const headers = ${JSON.stringify(excelHeaders)};

                searchInput.addEventListener("input", async function() {
                    const query = this.value.trim();

                    try {
                        const res = await fetch("/api/filter?q=" + encodeURIComponent(query));
                        const data = await res.json();

                        if (data.length > 0) {
                            tableBody.innerHTML = data.map(row => {
                                const cells = headers.map(h => \`<td>\${row[h] !== undefined ? row[h] : ""}</td>\`).join("");
                                return \`<tr>\${cells}</tr>\`;
                            }).join("");
                        } else {
                            tableBody.innerHTML = \`<tr><td colspan="\${headers.length}" style="text-align:center; padding:15px; color:#888;">Không tìm thấy dữ liệu phù hợp với "\${query}"</td></tr>\`;
                        }
                    } catch (e) {
                        console.error("Lỗi lọc dữ liệu:", e);
                    }
                });

                function openModal() { document.getElementById("addModal").style.display = "flex"; }
                function closeModal() { document.getElementById("addModal").style.display = "none"; }
            </script>
        </body>
        </html>
    `);
});

// THÊM SẢN PHẨM MỚI
app.post("/add-product", (req, res) => {
    const newRow = {};
    excelHeaders.forEach(h => {
        newRow[h] = (req.body[h] || "").trim();
    });

    excelData.unshift(newRow);
    res.redirect("/");
});

// IMPORT EXCEL (.XLSX)
app.post("/import-excel", upload.single("excelFile"), async (req, res) => {
    try {
        if (!req.file) return res.redirect("/");

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const sheet = workbook.worksheets[0];

        if (sheet) {
            const newHeaders = [];
            const newList = [];

            const headerRow = sheet.getRow(1);
            headerRow.eachCell((cell, colNumber) => {
                newHeaders[colNumber] = String(cell.value || "").trim();
            });

            const validHeaders = newHeaders.filter(Boolean);

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;

                const rowData = {};
                let hasData = false;

                headerRow.eachCell((cell, colNumber) => {
                    const headerName = newHeaders[colNumber];
                    if (headerName) {
                        const cellValue = row.getCell(colNumber).value;
                        const valStr = (cellValue && typeof cellValue === 'object' && cellValue.result !== undefined) 
                            ? String(cellValue.result) 
                            : String(cellValue ?? "").trim();

                        rowData[headerName] = valStr;
                        if (valStr) hasData = true;
                    }
                });

                if (hasData) {
                    newList.push(rowData);
                }
            });

            if (validHeaders.length > 0 && newList.length > 0) {
                excelHeaders = validHeaders;
                excelData = newList;
            }
        }
    } catch (err) {
        console.error("Lỗi khi đọc file Excel:", err);
    }
    res.redirect("/");
});

app.listen(PORT, () => {
    console.log(`Server đang chạy tại: http://localhost:${PORT}`);
});