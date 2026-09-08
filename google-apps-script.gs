function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({success: false, error: 'درخواست نامعتبر است.'});
  }

  var data = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (data.type === 'collaboration') {
    var name = String(data.name || '').trim();
    var email = String(data.email || '').trim();
    var project = String(data.project || '').trim();
    var budget = String(data.budget || '').trim();
    var details = String(data.details || '').trim();

    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !project || !details) {
      return jsonResponse({success: false, error: 'اطلاعات فرم کامل یا معتبر نیست.'});
    }

    sheet.appendRow([new Date(), 'درخواست همکاری', name, email, project, budget, details]);
    return jsonResponse({success: true});
  }

  var phone = String(data.phone || '').replace(/[\s-]/g, '');

  if (!/^09\d{9}$/.test(phone)) {
    return jsonResponse({success: false, error: 'شماره موبایل معتبر نیست.'});
  }

  sheet.appendRow([new Date(), 'دریافت شماره', phone]);

  return jsonResponse({success: true});
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
