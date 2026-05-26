const BARCODE_RECORD_SPREADSHEET_ID = '1OX5aKM-TuJPJIT_7v3Lq5WcvixxUo3ppt0c17Wqz8iE';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Production Update')
    .addItem('Update Actual Scan H9', 'updateActualScanH9')
    .addToUi();
}

function updateActualScanH9() {
  const targetSS = SpreadsheetApp.getActiveSpreadsheet();
  const recordSS = SpreadsheetApp.openById(BARCODE_RECORD_SPREADSHEET_ID);

  const recordSheet = recordSS.getSheetByName('barcode record');
  const targetSheet = targetSS.getSheetByName('Plan');

  if (!recordSheet) {
    throw new Error('ไม่พบชีต "barcode record" ในไฟล์ Barcode Record');
  }

  if (!targetSheet) {
    throw new Error('ไม่พบชีต "Plan" ในไฟล์ลงยอด H9');
  }

  const recordData = recordSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const countByJobOrder = {};

  // barcode record:
  // A = Date/Time
  // B = Job Order
  // C = Model
  // D = Barcode
  // E = Status
  for (let i = 1; i < recordData.length; i++) {
    const jobOrder = String(recordData[i][1]).trim();
    const status = String(recordData[i][4]).trim().toUpperCase();

    if (!jobOrder) continue;
    if (status !== 'OK') continue;

    countByJobOrder[jobOrder] = (countByJobOrder[jobOrder] || 0) + 1;
  }

  const actualScanValues = [];
  const diffValues = [];
  const progressValues = [];

  // ชีต Plan ในไฟล์ลงยอด H9:
  // D = Job Order
  // H = Plan
  // I = Actual เดิม ห้ามแก้
  // J = Actual Scan
  // K = diff
  // L = Progress %
  for (let i = 1; i < targetData.length; i++) {
    const jobOrder = String(targetData[i][3]).trim(); // column D
    const plan = Number(targetData[i][7]) || 0; // column H
    const actualScan = countByJobOrder[jobOrder] || 0;

    actualScanValues.push([actualScan]);
    diffValues.push([actualScan - plan]);

    const progress = plan > 0 ? actualScan / plan : 0;
    progressValues.push([progress]);
  }

  const startRow = 2;
  const rowCount = actualScanValues.length;

  // J = Actual Scan
  targetSheet.getRange(startRow, 10, rowCount, 1).setValues(actualScanValues);

  // K = diff
  targetSheet.getRange(startRow, 11, rowCount, 1).setValues(diffValues);

  // L = Progress %
  targetSheet.getRange(startRow, 12, rowCount, 1).setValues(progressValues);
  targetSheet.getRange(startRow, 12, rowCount, 1).setNumberFormat('0.00%');
}
