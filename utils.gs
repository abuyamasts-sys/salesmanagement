function getSpreadsheet_() {
  return SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
}

function getSheetByName_(sheetName) {
  var sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet tidak ditemukan: ' + sheetName);
  }
  return sheet;
}

function getSheetByNameOrNull_(sheetName) {
  return getSpreadsheet_().getSheetByName(sheetName);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function generateDocNumber_(prefix) {
  var stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMddHHmmss');
  return prefix + '-' + stamp;
}

function toBooleanLabel_(value) {
  return value ? 'Ya' : 'Tidak';
}

function getSheetData_(sheetName) {
  var sheet = getSheetByName_(sheetName);
  var values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    return [];
  }

  var headers = (values[0] || []).map(function(header) {
    return String(header || '').trim();
  });
  var rows = values.slice(1);

  return rows
    .filter(function(row) {
      return row.join('').toString().trim() !== '';
    })
    .map(function(row) {
      return mapRowToObject_(headers, row);
    });
}

function mapRowToObject_(headers, row) {
  var obj = {};

  headers.forEach(function(header, index) {
    var key = String(header || '').trim();
    if (!key) {
      return;
    }
    obj[key] = row[index];
  });

  return obj;
}

function normalizeText_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeSheetDateToYmd_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function appendRowByHeaders_(sheetName, data) {
  var sheet = getSheetByName_(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var row = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(data, header) ? data[header] : '';
  });

  sheet.appendRow(row);
}

function getNowParts_() {
  var now = new Date();

  return {
    tanggal: Utilities.formatDate(now, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    jam: Utilities.formatDate(now, APP_CONFIG.TIMEZONE, 'HH:mm:ss'),
    timestamp: now
  };
}

function updateRowByKey_(sheetName, keyField, keyValue, updates) {
  var sheet = getSheetByName_(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (!lastRow || lastRow < 2 || !lastCol) {
    throw new Error('Sheet belum memiliki data: ' + sheetName);
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var keyIndex = headers.indexOf(String(keyField || '').trim());

  if (keyIndex === -1) {
    throw new Error('Kolom key tidak ditemukan: ' + keyField);
  }

  var normalizedKey = String(keyValue).trim();
  var keyRange = sheet.getRange(2, keyIndex + 1, lastRow - 1, 1);
  var finder = keyRange.createTextFinder(normalizedKey).matchEntireCell(true);
  var match = finder.findNext();

  if (!match) {
    throw new Error('Data tidak ditemukan di ' + sheetName + ' untuk ' + keyField + ': ' + keyValue);
  }

  var rowNumber = match.getRow();
  var updatesObj = updates || {};
  var updatedAny = false;
  var rowRange = sheet.getRange(rowNumber, 1, 1, lastCol);
  var rowValues = rowRange.getValues()[0];

  headers.forEach(function(header, columnIndex) {
    if (!Object.prototype.hasOwnProperty.call(updatesObj, header)) {
      return;
    }

    rowValues[columnIndex] = updatesObj[header];
    updatedAny = true;
  });

  if (!updatedAny) {
    return mapRowToObject_(headers, rowValues);
  }

  rowRange.setValues([rowValues]);
  return mapRowToObject_(headers, rowValues);
}

function ensureSheetHeadersContain_(sheetName, headers) {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var safeHeaders = headers || [];
  var currentHeaders;
  var missingHeaders;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    if (safeHeaders.length) {
      sheet.getRange(1, 1, 1, safeHeaders.length).setValues([safeHeaders]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  missingHeaders = safeHeaders.filter(function(header) {
    return currentHeaders.indexOf(header) === -1;
  });

  if (missingHeaders.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}
