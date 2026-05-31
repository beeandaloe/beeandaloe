// ============================================================
// BEE & ALOE — Google Apps Script (Order Receiver)
// Paste this into: Google Sheet → Extensions → Apps Script
// Then: Deploy → New Deployment → Web App → Anyone → Deploy
// Copy the URL and paste it into index.html as APPS_SCRIPT_URL
// ============================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Customer Name',
        'Phone',
        'Email',
        'Delivery Address',
        'Items Ordered',
        'Total Amount',
        'Notes',
        'Status'
      ]);
      // Style header row
      var header = sheet.getRange(1, 1, 1, 9);
      header.setBackground('#2C5F2E');
      header.setFontColor('#FFFFFF');
      header.setFontWeight('bold');
    }

    // Parse incoming order data
    var data = JSON.parse(e.postData.contents);

    // Append the new order row
    sheet.appendRow([
      data.timestamp,
      data.name,
      data.phone,
      data.email,
      data.address,
      data.items,
      data.total,
      data.notes,
      'New Order' // default status
    ]);

    // Auto-resize columns for readability
    sheet.autoResizeColumns(1, 9);

    // Send email notification to Bee & Aloe
    MailApp.sendEmail({
      to: 'beeandaloebiz@gmail.com',
      subject: '🛒 New Order — Bee & Aloe',
      body:
        'You have a new order!\n\n' +
        'Customer: ' + data.name + '\n' +
        'Phone: ' + data.phone + '\n' +
        'Email: ' + data.email + '\n' +
        'Address: ' + data.address + '\n\n' +
        'Items:\n' + data.items + '\n\n' +
        'Total: ' + data.total + '\n\n' +
        'Notes: ' + data.notes + '\n\n' +
        'Time: ' + data.timestamp
    });

    // Return success
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET requests (for testing the script is live)
function doGet(e) {
  return ContentService
    .createTextOutput('Bee & Aloe order receiver is active ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}
