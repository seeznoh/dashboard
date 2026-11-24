function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? String(e.parameter.page).trim() : 'index';
  const templateName = page === 'input' ? 'input' : 'index';

  const template = HtmlService.createTemplateFromFile(templateName);
  template.appUrl = ScriptApp.getService().getUrl();

  return template
    .evaluate()
    .setTitle(page === 'input' ? '데이터 입력 및 조회 - 주식 자산 관리' : '종합 대시보드 - 주식 자산 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPortfolioData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('자산_입력_및_지수비교');
  if (!sheet) {
    throw new Error('자산_입력_및_지수비교 시트를 찾을 수 없습니다.');
  }

  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  const lastRow = sheet.getLastRow();
  const values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 11).getValues() : [];
  const items = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let month = row[0];

    if (month instanceof Date && !isNaN(month.getTime())) {
      month = Utilities.formatDate(month, tz, 'yyyy-MM');
    } else {
      month = month ? String(month).trim() : '';
    }

    if (!month) continue;

    const cleanNum = (val) => Number(String(val ?? '').replace(/,/g, '').trim()) || 0;

    items.push({
      month: month,
      deposit: cleanNum(row[1]),
      cumulativePrincipal: cleanNum(row[2]),
      totalAsset: cleanNum(row[3]),
      profit: cleanNum(row[4]),
      dividend: cleanNum(row[5]),
      sp500: row[6] === '' ? null : Number(row[6]),
      nasdaq: row[7] === '' ? null : Number(row[7]),
      kospi: row[8] === '' ? null : Number(row[8]),
      cumReturn: row[9] === '' ? null : Number(row[9]),
      ytdReturn: row[10] === '' ? null : Number(row[10])
    });
  }

  return {
    updatedAt: new Date().toISOString(),
    items: items
  };
}

function addPortfolioRow(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('자산_입력_및_지수비교');
  if (!sheet) {
    throw new Error('자산_입력_및_지수비교 시트를 찾을 수 없습니다.');
  }

  let month = data && data.month ? String(data.month).trim().replace(/[./]/g, '-') : '';
  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error('연월 형식은 YYYY-MM 이어야 합니다. (예: 2026-12)');
  }

  const deposit = Number(data.deposit) || 0;
  const cumulativePrincipal = Number(data.cumulativePrincipal) || 0;
  const totalAsset = Number(data.totalAsset) || 0;
  const profit = Number(data.profit) || 0;
  const dividend = Number(data.dividend) || 0;

  const prevLastRow = sheet.getLastRow();
  const newRow = prevLastRow + 1;

  const monthCell = sheet.getRange(newRow, 1);
  monthCell.setNumberFormat('@');
  monthCell.setValue(String(month));

  sheet.getRange(newRow, 2, 1, 5).setValues([[
    deposit,
    cumulativePrincipal,
    totalAsset,
    profit,
    dividend
  ]]);

  if (prevLastRow >= 2 && newRow > prevLastRow) {
    sheet.getRange(prevLastRow, 7, 1, 5).copyTo(sheet.getRange(newRow, 7, 1, 5));
  }

  return { added: true, month: month, row: newRow };
}

function deletePortfolioRow(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('자산_입력_및_지수비교');
  if (!sheet) {
    throw new Error('자산_입력_및_지수비교 시트를 찾을 수 없습니다.');
  }

  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  const cleanTarget = targetMonth ? String(targetMonth).trim() : '';
  const lastRow = sheet.getLastRow();
  let deletedCount = 0;

  for (let r = lastRow; r >= 2; r--) {
    const cellVal = sheet.getRange(r, 1).getValue();
    let mStr = '';
    if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
      mStr = Utilities.formatDate(cellVal, tz, 'yyyy-MM');
    } else {
      mStr = cellVal ? String(cellVal).trim() : '';
    }
    if (mStr === cleanTarget) {
      sheet.deleteRow(r);
      deletedCount++;
    }
  }

  return { deleted: true, month: cleanTarget, deletedCount: deletedCount };
}