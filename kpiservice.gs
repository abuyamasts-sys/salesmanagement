function normalizeMonthKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIMEZONE, 'yyyy-MM');
  }

  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})-(\d{2})/);
  if (!match) {
    return '';
  }
  return match[1] + '-' + match[2];
}

function getCurrentMonthKey_() {
  var now = getNowParts_();
  return normalizeMonthKey_(now.tanggal);
}

function listSalesUsers_() {
  return getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).filter(function(row) {
    var roleKey = normalizeText_(row.role);
    var statusKey = normalizeText_(row.status_aktif);
    var isSales = roleKey === 'sales' || roleKey.indexOf('sales') !== -1;
    var isActive = statusKey === 'aktif' || statusKey === '';

    return isSales && isActive;
  }).map(function(row) {
    return {
      user_id: String(row.user_id || '').trim(),
      nama_user: String(row.nama_user || '').trim(),
      role: String(row.role || '').trim()
    };
  }).filter(function(row) {
    return row.user_id;
  }).sort(function(left, right) {
    return String(left.nama_user || '').localeCompare(String(right.nama_user || ''), 'id-ID');
  });
}

function upsertSalesKpiTarget_(bulan, salesId, targetQty, catatan, approverId) {
  var monthKey = normalizeMonthKey_(bulan);
  var salesKey = String(salesId || '').trim();
  var qty = Number(targetQty || 0);

  if (!monthKey) {
    throw new Error('Bulan KPI tidak valid. Format wajib YYYY-MM.');
  }

  if (!salesKey) {
    throw new Error('Sales ID wajib dipilih.');
  }

  if (!(qty >= 0)) {
    throw new Error('Target KPI harus berupa angka.');
  }

  var now = getNowParts_();
  var sheetName = APP_CONFIG.SHEETS.KPI_TARGET_SALES;
  var headers = APP_CONFIG.HEADERS.KPI_TARGET_SALES;
  ensureSheetWithHeaders_(sheetName, headers);
  var sheet = getSheetByName_(sheetName);
  var data = sheet.getDataRange().getValues();
  var headerRow = data.length ? data[0].map(function(value) { return String(value || '').trim(); }) : headers;
  var bulanIndex = headerRow.indexOf('bulan');
  var salesIndex = headerRow.indexOf('sales_id');
  var existingRowNumber = 0;

  if (bulanIndex === -1 || salesIndex === -1) {
    throw new Error('Header KPI target tidak lengkap.');
  }

  for (var r = 1; r < data.length; r += 1) {
    var row = data[r] || [];
    var rowBulan = normalizeMonthKey_(row[bulanIndex]);
    var rowSales = String(row[salesIndex] || '').trim();
    if (rowBulan === monthKey && rowSales === salesKey) {
      existingRowNumber = r + 1;
      break;
    }
  }

  var payload = {
    bulan: monthKey,
    sales_id: salesKey,
    target_qty: qty,
    catatan: String(catatan || '').trim(),
    tanggal_input: now.tanggal + ' ' + now.jam,
    diinput_oleh: String(approverId || '').trim()
  };

  if (!existingRowNumber) {
    appendRowByHeaders_(sheetName, payload);
    SpreadsheetApp.flush();
    return payload;
  }

  var rowValues = headerRow.map(function(header) {
    return Object.prototype.hasOwnProperty.call(payload, header) ? payload[header] : '';
  });
  sheet.getRange(existingRowNumber, 1, 1, headerRow.length).setValues([rowValues]);
  SpreadsheetApp.flush();
  return payload;
}

function upsertSalesKpiTargetForAllSales_(bulan, targetQty, catatan, approverId) {
  var monthKey = normalizeMonthKey_(bulan);
  var qty = Number(targetQty || 0);
  var note = String(catatan || '').trim();
  var salesUsers = listSalesUsers_();
  var now = getNowParts_();
  var sheetName = APP_CONFIG.SHEETS.KPI_TARGET_SALES;
  var headers = APP_CONFIG.HEADERS.KPI_TARGET_SALES;
  var headerIndexMap = {};
  var updates = [];
  var newRows = [];
  var existingRowMap = {};
  var values;
  var headerRow;
  var results = [];

  if (!salesUsers.length) {
    throw new Error('Belum ada sales aktif untuk diupdate.');
  }

  if (!monthKey) {
    throw new Error('Bulan KPI tidak valid. Format wajib YYYY-MM.');
  }

  if (!(qty >= 0)) {
    throw new Error('Target KPI harus berupa angka.');
  }

  ensureSheetWithHeaders_(sheetName, headers);

  var sheet = getSheetByName_(sheetName);
  values = sheet.getDataRange().getValues();
  headerRow = values.length ? values[0].map(function(value) {
    return String(value || '').trim();
  }) : headers.slice();

  headerRow.forEach(function(header, index) {
    headerIndexMap[header] = index;
  });

  for (var r = 1; r < values.length; r += 1) {
    var row = values[r] || [];
    var rowBulan = normalizeMonthKey_(row[headerIndexMap.bulan]);
    var rowSales = String(row[headerIndexMap.sales_id] || '').trim();

    if (!rowBulan || !rowSales) {
      continue;
    }

    existingRowMap[rowBulan + '|' + rowSales] = r + 1;
  }

  salesUsers.forEach(function(user) {
    var salesKey = String(user.user_id || '').trim();
    var mapKey = monthKey + '|' + salesKey;
    var payload = {
      bulan: monthKey,
      sales_id: salesKey,
      target_qty: qty,
      catatan: note,
      tanggal_input: now.tanggal + ' ' + now.jam,
      diinput_oleh: String(approverId || '').trim()
    };
    var rowValues = headerRow.map(function(header) {
      return Object.prototype.hasOwnProperty.call(payload, header) ? payload[header] : '';
    });

    results.push(payload);

    if (existingRowMap[mapKey]) {
      updates.push({
        rowNumber: existingRowMap[mapKey],
        values: rowValues
      });
      return;
    }

    newRows.push(rowValues);
  });

  updates.forEach(function(item) {
    sheet.getRange(item.rowNumber, 1, 1, headerRow.length).setValues([item.values]);
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headerRow.length).setValues(newRows);
  }

  SpreadsheetApp.flush();

  return results;
}

function listSalesKpiTargetsByMonth_(bulan) {
  var monthKey = normalizeMonthKey_(bulan);
  var sheetName = APP_CONFIG.SHEETS.KPI_TARGET_SALES;

  if (!monthKey) {
    return [];
  }

  ensureSheetWithHeaders_(sheetName, APP_CONFIG.HEADERS.KPI_TARGET_SALES);

  var sheet = getSheetByName_(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (!lastRow || lastRow < 2 || !lastCol) {
    return [];
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values.slice(1).filter(function(row) {
    return row.join('').toString().trim() !== '';
  }).map(function(row) {
    return mapRowToObject_(headers, row);
  }).filter(function(row) {
    return normalizeMonthKey_(row.bulan) === monthKey;
  }).map(function(row) {
    return {
      bulan: normalizeMonthKey_(row.bulan),
      sales_id: String(row.sales_id || '').trim(),
      target_qty: Number(row.target_qty || 0),
      catatan: String(row.catatan || '').trim(),
      tanggal_input: String(row.tanggal_input || '').trim(),
      diinput_oleh: String(row.diinput_oleh || '').trim()
    };
  });
}

function getSalesKpiTargetQty_(bulan, salesId) {
  var monthKey = normalizeMonthKey_(bulan);
  var salesKey = String(salesId || '').trim();
  var sheetName = APP_CONFIG.SHEETS.KPI_TARGET_SALES;

  if (!monthKey || !salesKey) {
    return 0;
  }

  ensureSheetWithHeaders_(sheetName, APP_CONFIG.HEADERS.KPI_TARGET_SALES);

  var sheet = getSheetByName_(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (!lastRow || lastRow < 2 || !lastCol) {
    return 0;
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function(header) {
    return String(header || '').trim();
  });
  var bulanIndex = headers.indexOf('bulan');
  var salesIndex = headers.indexOf('sales_id');
  var targetIndex = headers.indexOf('target_qty');

  if (bulanIndex === -1 || salesIndex === -1 || targetIndex === -1) {
    return 0;
  }

  for (var i = 1; i < values.length; i += 1) {
    var row = values[i] || [];
    var rowBulan = normalizeMonthKey_(row[bulanIndex]);
    var rowSales = String(row[salesIndex] || '').trim();

    if (rowBulan === monthKey && rowSales === salesKey) {
      return Number(row[targetIndex] || 0);
    }
  }

  return 0;
}

function computeOrderQtyTotal_(noSo) {
  var details = getSalesOrderDetailsByNoSo_(noSo);

  return (details || []).reduce(function(result, row) {
    return result + Number(row.qty || 0);
  }, 0);
}

function recordKpiLogForOrderIfEligible_(noSo, actorId, tanggalSiapKirim) {
  var noSoKey = String(noSo || '').trim();
  if (!noSoKey) {
    return { recorded: false, reason: 'no_so kosong' };
  }

  var salesOrder = findSalesOrderByNoSo_(noSoKey) || null;
  if (!salesOrder) {
    return { recorded: false, reason: 'order tidak ditemukan' };
  }

  if (normalizeText_(salesOrder.status_order) !== 'siap kirim') {
    return { recorded: false, reason: 'status bukan siap kirim' };
  }

  if (normalizeText_(salesOrder.jenis_customer) !== 'baru') {
    return { recorded: false, reason: 'bukan customer baru' };
  }

  var sheetName = APP_CONFIG.SHEETS.KPI_LOG;
  ensureSheetWithHeaders_(sheetName, APP_CONFIG.HEADERS.KPI_LOG);

  var exists = getSheetData_(sheetName).some(function(row) {
    return String(row.no_so || '').trim() === noSoKey;
  });

  if (exists) {
    return { recorded: false, reason: 'sudah tercatat' };
  }

  var tanggalText = String(tanggalSiapKirim || '').trim();
  var monthKey = normalizeMonthKey_(tanggalText);
  if (!monthKey) {
    monthKey = getCurrentMonthKey_();
  }

  var qtyTotal = computeOrderQtyTotal_(noSoKey);
  var payload = {
    bulan: monthKey,
    no_so: noSoKey,
    sales_id: String(salesOrder.sales_id || '').trim(),
    jenis_customer: String(salesOrder.jenis_customer || '').trim(),
    qty_total: qtyTotal,
    tanggal_siap_kirim: tanggalText || (getNowParts_().tanggal + ' ' + getNowParts_().jam),
    dicatat_oleh: String(actorId || '').trim()
  };

  appendRowByHeaders_(sheetName, payload);
  return { recorded: true, payload: payload };
}

function getSalesKpiSummary_(bulan, salesId) {
  var monthKey = normalizeMonthKey_(bulan) || getCurrentMonthKey_();
  var salesKey = String(salesId || '').trim();

  var targetQty = getSalesKpiTargetQty_(monthKey, salesKey);

  var sheetName = APP_CONFIG.SHEETS.KPI_LOG;
  ensureSheetWithHeaders_(sheetName, APP_CONFIG.HEADERS.KPI_LOG);

  var rows = getSheetData_(sheetName).filter(function(row) {
    return normalizeMonthKey_(row.bulan) === monthKey &&
      String(row.sales_id || '').trim() === salesKey &&
      normalizeText_(row.jenis_customer) === 'baru';
  });

  var achieved = rows.reduce(function(result, row) {
    return result + Number(row.qty_total || 0);
  }, 0);

  var remaining = Math.max(Number(targetQty || 0) - achieved, 0);
  var percent = targetQty > 0 ? Math.min(Math.round((achieved / targetQty) * 100), 999) : 0;

  return {
    bulan: monthKey,
    sales_id: salesKey,
    target_qty: targetQty,
    achieved_qty: achieved,
    remaining_qty: remaining,
    achieved_percent: percent,
    orders: rows.map(function(row) {
      return {
        no_so: String(row.no_so || '').trim(),
        qty_total: Number(row.qty_total || 0),
        tanggal_siap_kirim: String(row.tanggal_siap_kirim || '').trim()
      };
    }).sort(function(left, right) {
      return String(right.tanggal_siap_kirim || '').localeCompare(String(left.tanggal_siap_kirim || ''));
    })
  };
}
