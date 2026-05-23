// ============================================
// Zero's 訂單接收 — Google Apps Script
// 部署為網頁應用程式後,把網址貼回前端 fetch POST 即可
// ============================================

const SHOP_NAME = "小玲の運動系列選品 — Zero's";
const SELLER_LINE = "@xiaoling.active";  // 改成你實際的 LINE ID
const BANK_NAME = "[請填寫你的銀行名稱]";
const BANK_ACCOUNT = "[請填寫你的銀行帳號]";

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
      // 標題列樣式
      const headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setFontWeight('bold').setBackground('#cfbedd');
      sheet.setFrozenRows(1);
    }

    // 寫入訂單
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

    // 自動寄訂單確認信給客人
    if (data.email) {
      sendCustomerEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderId: data.orderId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendCustomerEmail(data) {
  const subject = `${SHOP_NAME} · 訂單確認 #${data.orderId}`;
  const html = `
    <div style="font-family: -apple-system, 'PingFang TC', sans-serif; max-width: 600px; margin: 0 auto; padding: 24px; background: #f3ebdd; color: #1a1a1a;">
      <div style="background: #cfbedd; padding: 24px; border: 1px solid #1a1a1a;">
        <h1 style="margin: 0 0 8px; font-size: 28px;">訂單已成立 ♡</h1>
        <p style="margin: 0; font-size: 14px;">${SHOP_NAME}</p>
      </div>

      <div style="background: #fff; padding: 20px; border: 1px solid #1a1a1a; border-top: none;">
        <p>嗨 ${data.name},感謝你的訂購 ♡</p>
        <p>以下是訂單明細:</p>

        <table style="width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 13px;">
          <tr><td style="padding: 6px 0; color: #7a7a7a;">訂單編號</td><td style="padding: 6px 0; font-weight: 700;">${data.orderId}</td></tr>
          <tr><td style="padding: 6px 0; color: #7a7a7a;">時間</td><td style="padding: 6px 0;">${data.timestamp}</td></tr>
          <tr><td style="padding: 6px 0; color: #7a7a7a;">配送方式</td><td style="padding: 6px 0;">${data.shippingMethod}</td></tr>
          <tr><td style="padding: 6px 0; color: #7a7a7a;">收件資訊</td><td style="padding: 6px 0;">${data.address}</td></tr>
        </table>

        <h3 style="border-bottom: 2px dashed #1a1a1a; padding-bottom: 6px;">商品明細</h3>
        <pre style="background: #f3ebdd; padding: 12px; font-size: 13px; white-space: pre-wrap; border: 1px solid #d8cdb8;">${data.itemsText}</pre>

        <table style="width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 13px;">
          <tr><td>小計</td><td style="text-align: right;">NT$ ${data.subtotal.toLocaleString()}</td></tr>
          <tr><td style="color: #c66;">優惠折扣</td><td style="text-align: right; color: #c66;">- NT$ ${data.discount.toLocaleString()}</td></tr>
          <tr><td>運費</td><td style="text-align: right;">${data.shipping === 0 ? '免運' : 'NT$ ' + data.shipping}</td></tr>
          ${data.gift ? `<tr><td>滿額贈</td><td style="text-align: right;">${data.gift}</td></tr>` : ''}
          <tr><td style="border-top: 2px solid #1a1a1a; padding-top: 10px; font-size: 16px; font-weight: 700;">總計</td><td style="text-align: right; border-top: 2px solid #1a1a1a; padding-top: 10px; font-size: 22px; font-weight: 700;">NT$ ${data.total.toLocaleString()}</td></tr>
        </table>

        <h3 style="border-bottom: 2px dashed #1a1a1a; padding-bottom: 6px;">付款說明</h3>
        <p>請於 3 日內完成轉帳:</p>
        <ul style="padding-left: 18px; line-height: 1.8;">
          <li>銀行:<strong>${BANK_NAME}</strong></li>
          <li>帳號:<strong>${BANK_ACCOUNT}</strong></li>
          <li>金額:<strong>NT$ ${data.total.toLocaleString()}</strong></li>
        </ul>
        <p>完成轉帳後,請傳訊息至 LINE <strong>${SELLER_LINE}</strong>,並附上轉帳收據 + 訂單編號 <strong>${data.orderId}</strong>。</p>

        <p style="margin-top: 24px; color: #7a7a7a; font-size: 12px;">這封信由系統自動寄出,請勿直接回覆。如有問題請聯絡小玲 LINE ${SELLER_LINE}。</p>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: html
  });
}

// 測試用 — 在 Apps Script 編輯器執行這個 function 可確認權限正常
function testRun() {
  const dummy = {
    orderId: 'TEST-001',
    timestamp: new Date().toLocaleString('zh-TW'),
    name: '測試客人',
    email: '',  // 留空避免發信
    phone: '0912345678',
    shippingMethod: '宅配',
    address: '台北市測試路 1 號',
    note: '這是測試',
    itemsText: '撞色滾邊運動背心 × 1\n  顏色:藍撞咖 / 尺寸:L',
    subtotal: 680,
    discount: 0,
    shipping: 80,
    gift: '',
    total: 760
  };
  doPost({ postData: { contents: JSON.stringify(dummy) }});
}
