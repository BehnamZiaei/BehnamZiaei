function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({success: false, error: 'درخواست نامعتبر است.'});
  }

  var data = JSON.parse(e.postData.contents);
  var phone = String(data.phone || '').replace(/[\s-]/g, '');

  if (!/^09\d{9}$/.test(phone)) {
    return jsonResponse({success: false, error: 'شماره موبایل معتبر نیست.'});
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()[0]
    .appendRow([new Date(), phone]);

  return jsonResponse({success: true});
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
