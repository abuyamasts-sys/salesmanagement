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
    var tipeSalesKey = normalizeText_(row.tipe_sales);
    var userIdKey = normalizeText_(row.user_id);
    var isFreelance = tipeSalesKey === 'freelance';
    var isSales = roleKey === 'sales' || roleKey.indexOf('sales') !== -1;
    var isActive = statusKey === 'aktif' || statusKey === '';
    var isSlfMurni = tipeSalesKey.indexOf('slfm') === 0 || userIdKey.indexOf('slfm') === 0;
    var channelSalesKey = normalizeText_(row.channel_sales_default || (isFreelance || isSlfMurni ? 'SLF' : 'SLS'));

    return isSales && isActive && !isSlfMurni && channelSalesKey === 'sls';
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
  return computeOrderDeliveredQtyTotal_(noSo);
}

function hasKpiNumericValue_(value) {
  return value !== null && typeof value !== 'undefined' && String(value).trim() !== '';
}

function getKpiNumber_(value) {
  var number = Number(value);

  return isNaN(number) ? 0 : number;
}

function getKpiDeliveredQtyFromDetail_(detail) {
  var source = detail || {};

  if (hasKpiNumericValue_(source.qty_terkirim)) {
    return getKpiNumber_(source.qty_terkirim);
  }

  return getKpiNumber_(source.qty);
}

function computeOrderDeliveredQtyTotal_(noSo) {
  var details = getSalesOrderDetailsByNoSo_(noSo);

  return (details || []).reduce(function(result, row) {
    return result + getKpiDeliveredQtyFromDetail_(row);
  }, 0);
}

function buildKpiDeliveredQtyByNoSo_() {
  var result = {};

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, APP_CONFIG.HEADERS.SALES_ORDER_DETAIL);

  getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL).forEach(function(detail) {
    var noSo = String(detail.no_so || '').trim();

    if (!noSo) {
      return;
    }

    result[noSo] = Number(result[noSo] || 0) + getKpiDeliveredQtyFromDetail_(detail);
  });

  return result;
}

function buildSalesUserMapForKpi_() {
  var result = {};

  listSalesUsers_().forEach(function(user) {
    var userId = String(user.user_id || '').trim();

    if (userId) {
      result[userId] = user;
    }
  });

  return result;
}

function isSalesOrderDeliveredForKpi_(order) {
  var status = normalizeText_(order && order.status_order);

  return status === 'terkirim' || status === 'selesai';
}

function getSalesOrderKpiMonth_(order) {
  var source = order || {};

  return normalizeMonthKey_(source.tanggal_order || source.tanggal_kirim_rencana || source.tanggal_selesai || '');
}

function getKpiOrderQtyTotal_(order, deliveredQtyByNoSo) {
  var noSo = String(order && order.no_so || '').trim();
  var qtyMap = deliveredQtyByNoSo || {};

  if (noSo && Object.prototype.hasOwnProperty.call(qtyMap, noSo)) {
    return Number(qtyMap[noSo] || 0);
  }

  return getKpiNumber_(order && order.qty);
}

function listDeliveredSalesOrdersForKpi_(bulan, salesId) {
  var monthKey = normalizeMonthKey_(bulan) || getCurrentMonthKey_();
  var salesKey = String(salesId || '').trim();
  var salesUserMap = buildSalesUserMapForKpi_();
  var deliveredQtyByNoSo = buildKpiDeliveredQtyByNoSo_();

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);

  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(order) {
    var orderSalesId = String(order.sales_id || '').trim();

    if (salesKey && orderSalesId !== salesKey) {
      return false;
    }

    return !!(
      orderSalesId &&
      salesUserMap[orderSalesId] &&
      isSalesOrderDeliveredForKpi_(order) &&
      getSalesOrderKpiMonth_(order) === monthKey
    );
  }).map(function(order) {
    return {
      no_so: String(order.no_so || '').trim(),
      sales_id: String(order.sales_id || '').trim(),
      sales_name: String(order.sales_nama || '').trim(),
      nama_customer: String(order.nama_customer_input || order.nama_customer || '').trim(),
      jenis_customer: String(order.jenis_customer || '').trim(),
      qty_total: getKpiOrderQtyTotal_(order, deliveredQtyByNoSo),
      tanggal_order: normalizeSheetDateToYmd_(order.tanggal_order),
      jam_order: String(order.jam_order || '').trim(),
      tanggal_kirim_rencana: normalizeSheetDateToYmd_(order.tanggal_kirim_rencana),
      status_order: String(order.status_order || '').trim()
    };
  }).sort(function(left, right) {
    var leftStamp = String(left.tanggal_order || '') + ' ' + String(left.jam_order || '');
    var rightStamp = String(right.tanggal_order || '') + ' ' + String(right.jam_order || '');

    if (leftStamp !== rightStamp) {
      return String(rightStamp).localeCompare(String(leftStamp));
    }

    return String(right.no_so || '').localeCompare(String(left.no_so || ''));
  });
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

  if (!isSalesOrderDeliveredForKpi_(salesOrder)) {
    return { recorded: false, reason: 'status belum terkirim' };
  }

  if (!isActiveSalesUserForKpi_(salesOrder.sales_id)) {
    return { recorded: false, reason: 'sales_id bukan user SLS aktif' };
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
  var monthKey = getSalesOrderKpiMonth_(salesOrder);
  if (!monthKey) {
    monthKey = normalizeMonthKey_(tanggalText) || getCurrentMonthKey_();
  }

  var qtyTotal = computeOrderDeliveredQtyTotal_(noSoKey);
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

function isActiveSalesUserForKpi_(salesId) {
  var profile = getCurrentUserProfile(salesId);

  return !!(
    profile &&
    profile.authorized &&
    normalizeRoleKey_(profile.role) === 'sales' &&
    (normalizeText_(profile.status_aktif) === 'aktif' || normalizeText_(profile.status_aktif) === '') &&
    !profile.is_slfm &&
    normalizeText_(profile.channel_sales_default) === 'sls'
  );
}

function getSalesKpiSummary_(bulan, salesId) {
  var monthKey = normalizeMonthKey_(bulan) || getCurrentMonthKey_();
  var salesKey = String(salesId || '').trim();
  if (!isActiveSalesUserForKpi_(salesKey)) {
    return {
      bulan: monthKey,
      sales_id: salesKey,
      target_qty: 0,
      achieved_qty: 0,
      remaining_qty: 0,
      achieved_percent: 0,
      orders: []
    };
  }

  var targetQty = getSalesKpiTargetQty_(monthKey, salesKey);
  var rows = listDeliveredSalesOrdersForKpi_(monthKey, salesKey);

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
        nama_customer: String(row.nama_customer || '').trim(),
        qty_total: Number(row.qty_total || 0),
        tanggal_order: String(row.tanggal_order || '').trim(),
        jam_order: String(row.jam_order || '').trim(),
        tanggal_kirim_rencana: String(row.tanggal_kirim_rencana || '').trim(),
        status_order: String(row.status_order || '').trim()
      };
    }).sort(function(left, right) {
      var leftStamp = String(left.tanggal_order || '') + ' ' + String(left.jam_order || '');
      var rightStamp = String(right.tanggal_order || '') + ' ' + String(right.jam_order || '');

      return String(rightStamp).localeCompare(String(leftStamp));
    })
  };
}

function deriveSalesKpiProgressStatus_(targetQty, achievedQty) {
  var target = Number(targetQty || 0);
  var achieved = Number(achievedQty || 0);

  if (!(target > 0)) {
    return 'Belum Ada Target';
  }

  if (!(achieved > 0)) {
    return 'Belum Mulai';
  }

  if (achieved < target) {
    return 'Berjalan';
  }

  if (achieved === target) {
    return 'Tercapai';
  }

  return 'Melebihi';
}

function buildApproverKpiProgressSummary_(rows) {
  var safeRows = Array.isArray(rows) ? rows : [];
  var rowsWithTarget = safeRows.filter(function(row) {
    return Number(row.target_qty || 0) > 0;
  });
  var achievedRows = rowsWithTarget.filter(function(row) {
    var status = String(row.status_kpi || '').trim();
    return status === 'Tercapai' || status === 'Melebihi';
  });
  var averageProgress = 0;

  if (rowsWithTarget.length) {
    averageProgress = Math.round(rowsWithTarget.reduce(function(sum, row) {
      return sum + Number(row.progress_percent || 0);
    }, 0) / rowsWithTarget.length);
  }

  return {
    total_sales: safeRows.length,
    sales_with_target: rowsWithTarget.length,
    sales_achieved: achievedRows.length,
    sales_not_achieved: Math.max(rowsWithTarget.length - achievedRows.length, 0),
    sales_without_target: Math.max(safeRows.length - rowsWithTarget.length, 0),
    total_target_qty: safeRows.reduce(function(sum, row) {
      return sum + Number(row.target_qty || 0);
    }, 0),
    total_achieved_qty: safeRows.reduce(function(sum, row) {
      return sum + Number(row.achieved_qty || 0);
    }, 0),
    average_progress_percent: averageProgress
  };
}

function buildApproverKpiProgressRows_(bulan) {
  var monthKey = normalizeMonthKey_(bulan) || getCurrentMonthKey_();
  var salesUsers = listSalesUsers_();
  var targets = listSalesKpiTargetsByMonth_(monthKey);
  var targetMap = {};
  var deliveredRows = listDeliveredSalesOrdersForKpi_(monthKey);
  var orderMap = {};

  targets.forEach(function(target) {
    var salesId = String(target.sales_id || '').trim();
    if (!salesId) {
      return;
    }

    targetMap[salesId] = target;
  });

  deliveredRows.forEach(function(row) {
    var salesId = String(row.sales_id || '').trim();
    var orderRow;

    if (!salesId) {
      return;
    }

    if (!orderMap[salesId]) {
      orderMap[salesId] = [];
    }

    orderRow = {
      no_so: String(row.no_so || '').trim(),
      qty_total: Number(row.qty_total || 0),
      tanggal_order: String(row.tanggal_order || '').trim(),
      jam_order: String(row.jam_order || '').trim(),
      tanggal_kirim_rencana: String(row.tanggal_kirim_rencana || '').trim(),
      status_order: String(row.status_order || '').trim()
    };
    orderMap[salesId].push(orderRow);
  });

  return salesUsers.map(function(user) {
    var salesId = String(user.user_id || '').trim();
    var targetRow = targetMap[salesId] || null;
    var orders = (orderMap[salesId] || []).slice().sort(function(left, right) {
      var leftStamp = String(left.tanggal_order || '') + ' ' + String(left.jam_order || '');
      var rightStamp = String(right.tanggal_order || '') + ' ' + String(right.jam_order || '');

      return String(rightStamp).localeCompare(String(leftStamp));
    });
    var targetQty = Number(targetRow && targetRow.target_qty || 0);
    var achievedQty = orders.reduce(function(sum, row) {
      return sum + Number(row.qty_total || 0);
    }, 0);
    var remainingQty = Math.max(targetQty - achievedQty, 0);
    var progressPercent = targetQty > 0 ? Math.round((achievedQty / targetQty) * 100) : 0;

    return {
      sales_id: salesId,
      sales_name: String(user.nama_user || salesId || '').trim(),
      target_qty: targetQty,
      achieved_qty: achievedQty,
      remaining_qty: remainingQty,
      progress_percent: progressPercent,
      order_count: orders.length,
      status_kpi: deriveSalesKpiProgressStatus_(targetQty, achievedQty),
      last_target_update: targetRow ? String(targetRow.tanggal_input || '').trim() : '',
      last_target_updated_by: targetRow ? String(targetRow.diinput_oleh || '').trim() : '',
      catatan_target: targetRow ? String(targetRow.catatan || '').trim() : '',
      orders: orders
    };
  });
}

function getApproverKpiProgressData_(bulan) {
  var monthKey = normalizeMonthKey_(bulan) || getCurrentMonthKey_();
  var rows = buildApproverKpiProgressRows_(monthKey);

  rows.sort(function(left, right) {
    var statusOrder = {
      'Belum Ada Target': 1,
      'Belum Mulai': 2,
      'Berjalan': 3,
      'Tercapai': 4,
      'Melebihi': 5
    };
    var leftRank = statusOrder[left.status_kpi] || 99;
    var rightRank = statusOrder[right.status_kpi] || 99;

    if (leftRank !== rightRank) {
      return leftRank - rightRank;
    }

    if (Number(left.progress_percent || 0) !== Number(right.progress_percent || 0)) {
      return Number(left.progress_percent || 0) - Number(right.progress_percent || 0);
    }

    return String(left.sales_name || '').localeCompare(String(right.sales_name || ''), 'id-ID');
  });

  return {
    bulan: monthKey,
    salesUsers: listSalesUsers_(),
    targets: listSalesKpiTargetsByMonth_(monthKey),
    summary: buildApproverKpiProgressSummary_(rows),
    rows: rows
  };
}
