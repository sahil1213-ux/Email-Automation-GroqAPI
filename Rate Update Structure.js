function setQuotationHeaders() {
    const ss = SpreadsheetApp.openById(
      "1ZGzg1M8oQT5AFT7Vkpq5fL9vbay-IyQHNui8eG5cRX4"
    );
    const sheet = ss.getSheetByName("Rate Updates");
  
      const headers = [
      "Sender Email",
      "Sender Name",
      "Received Date & Time",
      "Company",
      "Subject",
      "Origin",
      "Destination",
      "20 ft Rate($)",
      "40 ft Rate($)",
      "Carrier",
      "Transit Time",
      "Free Time(Days)",
      "Additional Charges",
      "Validity / Urgency",
      "Full Body / Notes"
    ];
  
     // Set header values in row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📩 Email Actions")
    .addItem("🔄 Refresh Rate Updates", "processRateUpdates")
    .addItem("🚚 Refresh Multi Route Prices", "processMultiRoutePrices")
    .addToUi();
}

