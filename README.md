# Heat-Pump Google app Script

// ═══════════════════════════════════════════════════════════════
// HeatPro — Google Apps Script (Paste this into your Google Sheet)
// 
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Go to Extensions → Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click Deploy → New Deployment
// 5. Type: "Web app"
// 6. Execute as: "Me"
// 7. Who has access: "Anyone"
// 8. Click Deploy and copy the Web App URL
// 9. Paste that URL into your HeatPro tool's settings
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'HeatPro Data';

// Column headers for the sheet
const HEADERS = [
  'S.No',
  'Timestamp',
  'Customer Name',
  'Contact Person',
  'Mobile',
  'Email',
  'Site Address',
  'City',
  'State',
  'Facility Type',
  'Purpose of Hot Water',
  'Hot Water Temp (°C)',
  'Cold Water Temp (°C)',
  'Daily Demand (L/day)',
  'Demand Mode',
  'Bath Users',
  'Bath LPD',
  'Kitchen Users',
  'Kitchen LPD',
  'Laundry Kg',
  'Laundry LPD',
  'Process Users',
  'Process LPD',
  'Peak Pattern',
  'Operating Hours',
  'Existing System',
  'Existing Capacity',
  'Problems Faced',
  'Installation Space',
  'Electrical Phase',
  'Voltage',
  'Water Quality',
  'Ambient Temp',
  'Priority',
  'Budget',
  'Notes',
  // Calculated results
  'Delta T (°C)',
  'Heat Load (kWh/day)',
  'COP',
  'HP Capacity (kW)',
  'HP Capacity (TR)',
  'No. of Units',
  'Unit Size (kW)',
  'Storage Tank (L)',
  'Daily Elec Input (kWh)',
  'Daily Cost HP (₹)',
  'Daily Cost Geyser (₹)',
  'Annual Saving (₹)',
  'Saving %',
  'Est. Payback (years)',
  'CO2 Reduction (tonnes/yr)'
];

/**
 * Handles incoming POST requests from the HeatPro tool
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    
    // Get next serial number
    const lastRow = sheet.getLastRow();
    const serialNo = lastRow <= 1 ? 1 : lastRow; // Row 1 is header
    
    // Build row data matching HEADERS order
    const row = [
      serialNo,
      new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      data.customerName || '',
      data.contactPerson || '',
      data.mobile || '',
      data.email || '',
      data.address || '',
      data.city || '',
      data.state || '',
      data.facilityType || '',
      data.purposes || '',
      data.hotTemp || '',
      data.coldTemp || '',
      data.dailyLitres || '',
      data.demandMode || '',
      data.bathUsers || '',
      data.bathLpd || '',
      data.kitchenUsers || '',
      data.kitchenLpd || '',
      data.laundryKg || '',
      data.laundryLpd || '',
      data.processUsers || '',
      data.processLpd || '',
      data.peakPattern || '',
      data.opHours || '',
      data.existingSystem || '',
      data.existingCapacity || '',
      data.problems || '',
      data.installSpace || '',
      data.elecPhase || '',
      data.voltage || '',
      data.waterQuality || '',
      data.ambientTemp || '',
      data.priority || '',
      data.budget || '',
      data.notes || '',
      // Calculated results
      data.deltaT || '',
      data.heatLoadKWh || '',
      data.cop || '',
      data.hpCapacityKW || '',
      data.hpCapacityTR || '',
      data.numUnits || '',
      data.unitSizeKW || '',
      data.storageLitres || '',
      data.dailyElecInput || '',
      data.dailyCostHP || '',
      data.dailyCostGeyser || '',
      data.annualSaving || '',
      data.savingPct || '',
      data.paybackYears || '',
      data.co2Reduction || ''
    ];
    
    sheet.appendRow(row);
    
    // Auto-resize columns on first entry
    if (lastRow <= 1) {
      for (let i = 1; i <= HEADERS.length; i++) {
        sheet.autoResizeColumn(i);
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', row: serialNo }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles GET requests (for testing the endpoint)
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: 'ok', 
      message: 'HeatPro Google Sheets API is active',
      version: '2.0'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gets the data sheet, creating it with headers if it doesn't exist
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    
    // Add headers
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    
    // Style the header row
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a2545');
    headerRange.setFontColor('#ff6b1a');
    headerRange.setFontFamily('Arial');
    headerRange.setFontSize(10);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths for key columns
    sheet.setColumnWidth(1, 50);   // S.No
    sheet.setColumnWidth(2, 160);  // Timestamp
    sheet.setColumnWidth(3, 200);  // Customer Name
  }
  
  return sheet;
}
