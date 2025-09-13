/**
 * @OnlyCurrentDoc
 */

// --- THAY ĐỔI 1: Thêm ID của bảng tính của bạn vào đây ---
// ID này được lấy từ URL bạn đã cung cấp.
const SPREADSHEET_ID = '1aFoYyCCiE0yUT6jGlmTujha34CHZTPAW3-IVmiTI5Gw';

// Tên các sheet trong bảng tính
const NHAPXUAT_SHEET = 'nhapxuat';
const CHITIET_SHEET = 'chitiet';
const SANPHAM_SHEET = 'sanpham';

function doGet() {
  return HtmlService.createTemplateFromFile('WebAppInterface').evaluate()
    .setTitle('Giao Diện Nhập Liệu Kho')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getInitialData() {
  try {
    // --- THAY ĐỔI 2: Mở bảng tính cụ thể bằng ID ---
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sanphamSheet = ss.getSheetByName(SANPHAM_SHEET);
    const nhapxuatSheet = ss.getSheetByName(NHAPXUAT_SHEET);

    // Thêm kiểm tra để đảm bảo các sheet tồn tại
    if (!sanphamSheet || !nhapxuatSheet) {
        throw new Error("Không tìm thấy một trong các sheet cần thiết (sanpham, nhapxuat). Vui lòng kiểm tra lại tên sheet.");
    }

    // Lấy dữ liệu sản phẩm để tìm kiếm (Mã SP, Tên SP, Mã nội bộ)
    const lastRow = sanphamSheet.getLastRow();
    if (lastRow < 2) { // Nếu không có sản phẩm nào
        return { productList: [], existingTransactionCodes: [] };
    }
    const productData = sanphamSheet.getRange('A2:H' + lastRow).getValues();
    const productList = productData.map(row => ({
      id: row[0].toString().trim(),
      name: row[1].toString().trim(),
      internalCode: row[7] ? row[7].toString().trim() : '' // Cột H
    })).filter(p => p.id); // Lọc ra những hàng có Mã SP

    // Lấy mã giao dịch đã tồn tại
    const lastTransRow = nhapxuatSheet.getLastRow();
    const existingTransactionCodes = lastTransRow < 2 ? [] : nhapxuatSheet.getRange('A2:A' + lastTransRow).getValues().flat().filter(String).map(code => code.toString().trim());

    return { 
      productList: productList,
      existingTransactionCodes: existingTransactionCodes
    };
  } catch (e) {
    Logger.log('Lỗi khi lấy dữ liệu ban đầu: ' + e.toString());
    throw new Error('Không thể tải dữ liệu từ Google Sheet. ' + e.message);
  }
}

function getTransactionDetails(orderCode) {
  try {
    if (!orderCode) return null;

    // --- THAY ĐỔI 3: Mở bảng tính cụ thể bằng ID ---
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nhapxuatSheet = ss.getSheetByName(NHAPXUAT_SHEET);
    const chitietSheet = ss.getSheetByName(CHITIET_SHEET);
    
    const nhapxuatData = nhapxuatSheet.getDataRange().getValues();
    const chitietData = chitietSheet.getDataRange().getValues();

    // Tìm giao dịch trong sheet nhapxuat (bỏ qua header)
    let transactionRow = null;
    for (let i = 1; i < nhapxuatData.length; i++) {
      if (nhapxuatData[i][0].toString().trim() === orderCode.toString().trim()) {
        transactionRow = nhapxuatData[i];
        break;
      }
    }

    if (!transactionRow) return null; // Không tìm thấy giao dịch

    const transactionType = transactionRow[1]; // Cột B: Loại (Nhập/Xuất)
    const products = {};

    // Tìm các sản phẩm liên quan trong sheet chitiet (bỏ qua header)
    for (let i = 1; i < chitietData.length; i++) {
      const row = chitietData[i];
      if (row[2].toString().trim() === orderCode.toString().trim()) { // Cột C: mã nx
        const productCode = row[3].toString().trim(); // Cột D: Mã SP
        const productType = row[7]; // Cột H: Loại
        const productKey = `${productCode}_${productType}`;
        
        products[productKey] = {
          code: productCode,
          name: row[4],  // Cột E: Tên SP
          qty: parseInt(row[5], 10),    // Cột F: SLNX
          type: productType
        };
      }
    }

    return {
      type: transactionType,
      products: products
    };
  } catch (e) {
    Logger.log(`Lỗi khi lấy chi tiết giao dịch ${orderCode}: ` + e.toString());
    return null;
  }
}


function saveTransactions(transactions) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Đợi tối đa 30 giây

  try {
    // --- THAY ĐỔI 4: Mở bảng tính cụ thể bằng ID ---
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nhapxuatSheet = ss.getSheetByName(NHAPXUAT_SHEET);
    const chitietSheet = ss.getSheetByName(CHITIET_SHEET);
    const sanphamSheet = ss.getSheetByName(SANPHAM_SHEET);

    const sanphamData = sanphamSheet.getDataRange().getValues();
    const productIndexMap = new Map();
    for(let i=1; i<sanphamData.length; i++) {
        if(sanphamData[i][0]) {
            productIndexMap.set(sanphamData[i][0].toString().trim(), i);
        }
    }
    
    let chitietRowsToAdd = [];
    let nhapxuatRowsToAdd = [];
    let newTransactionCodes = [];
    let cancellationOccurred = false;

    transactions.forEach(entry => {
      const now = new Date();
      const user = Session.getActiveUser().getEmail() || 'Không xác định';

      if (entry.transactionType === 'Hủy') {
        cancellationOccurred = true;
        handleCancellation(entry, sanphamData, productIndexMap, ss);
      } else {
        const orderCodes = entry.type === 'single' ? [entry.orderCode] : entry.orderCodes;
        const totalProductTypes = new Set(Object.values(entry.products).map(p => p.code)).size;
        
        orderCodes.forEach(orderCode => {
          nhapxuatRowsToAdd.push([
            orderCode,
            entry.transactionType,
            totalProductTypes,
            '', // Nhập thủ công
            'Đã Soạn Hàng',
            `Đã Soạn Hàng - ${user} : ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yy HH:mm:ss')}`,
            '' // DVVC
          ]);
          newTransactionCodes.push(orderCode.toString().trim());

          for (const productKey in entry.products) {
            const product = entry.products[productKey];
            const qtyPerOrder = (entry.type === 'bulk' && orderCodes.length > 0) 
                                ? Math.round(product.qty / orderCodes.length) 
                                : product.qty;
            
            chitietRowsToAdd.push([
              Utilities.getUuid(),
              entry.transactionType,
              orderCode,
              product.code, 
              product.name,
              qtyPerOrder,
              now,
              product.type
            ]);

            const productRowIndex = productIndexMap.get(product.code.toString().trim());
            if (productRowIndex !== undefined) {
              const stockCol = product.type === 'Bình thường' ? 3 : 5; // Cột D (index 3) hoặc F (index 5)
              const currentStock = sanphamData[productRowIndex][stockCol] || 0;
              if (entry.transactionType === 'Nhập') {
                sanphamData[productRowIndex][stockCol] = currentStock + qtyPerOrder;
              } else { // Xuất
                sanphamData[productRowIndex][stockCol] = currentStock - qtyPerOrder;
              }
            }
          }
        });
      }
    });

    if (nhapxuatRowsToAdd.length > 0) {
      nhapxuatSheet.getRange(nhapxuatSheet.getLastRow() + 1, 1, nhapxuatRowsToAdd.length, nhapxuatRowsToAdd[0].length).setValues(nhapxuatRowsToAdd);
    }
    if (chitietRowsToAdd.length > 0) {
      chitietSheet.getRange(chitietSheet.getLastRow() + 1, 1, chitietRowsToAdd.length, chitietRowsToAdd[0].length).setValues(chitietRowsToAdd);
    }

    sanphamSheet.getRange(1, 1, sanphamData.length, sanphamData[0].length).setValues(sanphamData);
    
    const finalTransactionCodes = (cancellationOccurred || newTransactionCodes.length > 0) 
        ? nhapxuatSheet.getRange('A2:A' + nhapxuatSheet.getLastRow()).getValues().flat().filter(String).map(code => code.toString().trim())
        : null;

    return { 
      success: true, 
      message: 'Đã xử lý thành công!',
      updatedTransactionCodes: finalTransactionCodes
    };

  } catch (e) {
    Logger.log('Lỗi khi lưu giao dịch: ' + e.toString() + "\n" + e.stack);
    return { success: false, message: 'Lỗi máy chủ: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function handleCancellation(entry, sanphamData, productIndexMap, spreadsheet) {
  const orderCode = entry.orderCode;
  const originalType = entry.originalType;
  const products = entry.products;

  // 1. Cập nhật lại tồn kho
  for (const productKey in products) {
    const product = products[productKey];
    const productRowIndex = productIndexMap.get(product.code.toString().trim());
    if (productRowIndex !== undefined) {
      const stockCol = product.type === 'Bình thường' ? 3 : 5; // Cột D (index 3) hoặc F (index 5)
      const currentStock = sanphamData[productRowIndex][stockCol] || 0;
      
      if (originalType === 'Nhập') {
        sanphamData[productRowIndex][stockCol] = currentStock - product.qty;
      } else { // Xuất
        sanphamData[productRowIndex][stockCol] = currentStock + product.qty;
      }
    }
  }

  // 2. Xóa dòng trong sheet `nhapxuat` và `chitiet`
  const nhapxuatSheet = spreadsheet.getSheetByName(NHAPXUAT_SHEET);
  const chitietSheet = spreadsheet.getSheetByName(CHITIET_SHEET);

  deleteRowsByValue(nhapxuatSheet, 0, orderCode);
  deleteRowsByValue(chitietSheet, 2, orderCode);
}

function deleteRowsByValue(sheet, columnIndex, valueToDelete) {
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    data.forEach((row, index) => {
        if (row[columnIndex] && row[columnIndex].toString().trim() === valueToDelete.toString().trim()) {
            rowsToDelete.push(index + 1);
        }
    });

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
    }
}
