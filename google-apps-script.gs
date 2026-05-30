// ============================================
// Zero's 訂單接收 — Google Apps Script
//
// 用途:接收前端 POST 過來的訂單,寫入這份試算表
// 部署:右上「部署」→「新增部署作業」→ 選「網頁應用程式」
//      執行身分=我、誰可存取=任何人 → 部署 → 複製網址貼回前端
// ============================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 第一次寫入前,自動建立標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '訂單編號', '訂單時間', '姓名', 'Email', '電話',
        '配送方式', '地址 / 取貨門市', '備註',
        '商品明細', '小計', '折扣', '運費', '滿額贈', '總計', '狀態'
      ]);
      const headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setFontWeight('bold').setBackground('#cfbedd');
      sheet.setFrozenRows(1);
      // 自動調整欄寬
      sheet.autoResizeColumns(1, 15);
    }

    // 寫入訂單資料
    sheet.appendRow([
      data.orderId,
      data.timestamp,
      data.name,
      data.email,
      data.phone,
      data.shippingMethod,
      data.address,
      data.note,
      data.itemsText,
      data.subtotal,
      data.discount,
      data.shipping,
      data.gift || '',
      data.total,
      '待付款'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderId: data.orderId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 測試用 — 在 Apps Script 編輯器執行此 function 可確認權限與寫入正常
function testRun() {
  const dummy = {
    orderId: 'TEST-001',
    timestamp: new Date().toLocaleString('zh-TW'),
    name: '測試客人',
    email: 'test@example.com',
    phone: '0912345678',
    shippingMethod: '宅配',
    address: '台北市測試路 1 號',
    note: '這是測試訂單',
    itemsText: '撞色滾邊運動背心 × 1\n  顏色:藍撞咖 / 尺寸:L  NT$ 680',
    subtotal: 680,
    discount: 0,
    shipping: 80,
    gift: '',
    total: 760
  };
  doPost({ postData: { contents: JSON.stringify(dummy) }});
  Logger.log('測試完成,請查看試算表');
}
