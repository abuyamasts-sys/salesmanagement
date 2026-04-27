function getControllerDashboardData(userId, filters) {
  requireCurrentUserRole_(['Controller', 'MTR'], userId);

  var currentUser = getCurrentUserProfile(userId);
  var filter = buildControllerDashboardFilter_(filters);
  var salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER);
  var salesOrderDetails = getControllerSalesOrderDetails_();
  var deliveryOrders = getControllerDeliveryOrders_();
  var billingData = getAdminBillingData_(currentUser, {});
  var summary = buildControllerDashboardSummary_(salesOrders, salesOrderDetails, filter);
  var statusSummary = buildControllerStatusSummary_(salesOrders);
  var opsSummary = buildControllerOpsSummary_(salesOrders, deliveryOrders, billingData);
  var trend = buildControllerTrend7Days_(salesOrders, filter.today);
  var salesKpiSnapshot = buildControllerSalesKpiSnapshot_();
  var topSales = buildControllerTopSales_(salesOrders);
  var topCustomers = buildControllerTopCustomers_(salesOrders);
  var topItems = buildControllerTopItems_(salesOrderDetails);
  var fieldMonitoring = getFieldMonitoringDashboardData(userId, {
    tanggal: normalizeSheetDateToYmd_((filters || {}).monitoringDate || '') || filter.today
  });

  return toClientValue_({
    currentUser: currentUser,
    filter: filter,
    summary: summary,
    salesKpiSnapshot: salesKpiSnapshot,
    statusSummary: statusSummary,
    opsSummary: opsSummary,
    trend: trend,
    topSales: topSales,
    topCustomers: topCustomers,
    topItems: topItems,
    fieldMonitoring: fieldMonitoring,
    lastUpdated: getControllerDashboardLastUpdated_()
  });
}

function buildControllerDashboardFilter_(filters) {
  var now = getNowParts_().timestamp;
  var startOfWeek = getControllerStartOfWeek_(now);
  var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  var payload = filters || {};

  return {
    preset: String(payload.preset || '').trim() || 'today',
    startDate: normalizeSheetDateToYmd_(payload.startDate || ''),
    endDate: normalizeSheetDateToYmd_(payload.endDate || ''),
    today: normalizeSheetDateToYmd_(now),
    weekStart: normalizeSheetDateToYmd_(startOfWeek),
    monthStart: normalizeSheetDateToYmd_(startOfMonth)
  };
}

function buildControllerDashboardSummary_(salesOrders, salesOrderDetails, filter) {
  var orders = Array.isArray(salesOrders) ? salesOrders : [];
  var details = Array.isArray(salesOrderDetails) ? salesOrderDetails : [];
  var detailQtyByNoSo = buildControllerDetailQtyByNoSo_(details);
  var today = String(filter.today || '').trim();
  var weekStart = String(filter.weekStart || '').trim();
  var monthStart = String(filter.monthStart || '').trim();

  return {
    so_today: buildControllerOrderCount_(orders, today, today),
    omzet_today: buildControllerOmzetTotal_(orders, today, today),
    customers_today: buildControllerUniqueCustomerCount_(orders, today, today),
    qty_today: buildControllerQtyTotal_(orders, detailQtyByNoSo, today, today),
    so_week: buildControllerOrderCount_(orders, weekStart, today),
    omzet_week: buildControllerOmzetTotal_(orders, weekStart, today),
    so_month: buildControllerOrderCount_(orders, monthStart, today),
    omzet_month: buildControllerOmzetTotal_(orders, monthStart, today)
  };
}

function buildControllerSalesKpiSnapshot_() {
  var monthKey = getCurrentMonthKey_();
  var progressData = getApproverKpiProgressData_(monthKey);
  var rows = Array.isArray(progressData && progressData.rows) ? progressData.rows : [];
  var summary = progressData && progressData.summary ? progressData.summary : {};
  var totalTarget = Number(summary.total_target_qty || 0);
  var totalAchieved = Number(summary.total_achieved_qty || 0);
  var totalRemaining = rows.reduce(function(sum, row) {
    return sum + Number(row.remaining_qty || 0);
  }, 0);
  var totalProgressPercent = totalTarget > 0 ? Math.round((totalAchieved / totalTarget) * 100) : 0;
  var statusOrder = {
    'Belum Ada Target': 1,
    'Belum Mulai': 2,
    'Berjalan': 3,
    'Tercapai': 4,
    'Melebihi': 5
  };

  return {
    bulan: monthKey,
    summary: {
      total_target_qty: totalTarget,
      total_achieved_qty: totalAchieved,
      total_remaining_qty: totalRemaining,
      total_progress_percent: totalProgressPercent
    },
    rows: rows.slice().sort(function(left, right) {
      var leftRank = statusOrder[String(left.status_kpi || '').trim()] || 99;
      var rightRank = statusOrder[String(right.status_kpi || '').trim()] || 99;

      if (leftRank !== rightRank) {
        return leftRank - rightRank;
      }

      if (Number(left.progress_percent || 0) !== Number(right.progress_percent || 0)) {
        return Number(left.progress_percent || 0) - Number(right.progress_percent || 0);
      }

      return String(left.sales_name || left.sales_id || '').localeCompare(String(right.sales_name || right.sales_id || ''), 'id-ID');
    }).map(function(row) {
      return {
        sales_id: String(row.sales_id || '').trim(),
        sales_name: String(row.sales_name || row.sales_id || '-').trim() || '-',
        target_qty: Number(row.target_qty || 0),
        achieved_qty: Number(row.achieved_qty || 0),
        remaining_qty: Number(row.remaining_qty || 0),
        progress_percent: Number(row.progress_percent || 0),
        status_kpi: String(row.status_kpi || 'Belum Ada Target').trim() || 'Belum Ada Target'
      };
    })
  };
}

function getControllerDashboardLastUpdated_() {
  var now = getNowParts_();
  return now.tanggal + ' ' + now.jam;
}

function getControllerSalesOrderDetails_() {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);

  if (!sheet) {
    return [];
  }

  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);
}

function getControllerDeliveryOrders_() {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.SURAT_JALAN);

  if (!sheet) {
    return [];
  }

  return getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN);
}

function buildControllerDetailQtyByNoSo_(details) {
  return (details || []).reduce(function(result, detail) {
    var noSo = String(detail.no_so || '').trim();

    if (!noSo) {
      return result;
    }

    result[noSo] = (result[noSo] || 0) + Number(detail.qty || 0);
    return result;
  }, {});
}

function buildControllerOrderCount_(orders, startDate, endDate) {
  return filterControllerOrdersByDateRange_(orders, startDate, endDate).length;
}

function buildControllerOmzetTotal_(orders, startDate, endDate) {
  return filterControllerOrdersByDateRange_(orders, startDate, endDate).reduce(function(sum, order) {
    return sum + getControllerOrderAmount_(order);
  }, 0);
}

function buildControllerUniqueCustomerCount_(orders, startDate, endDate) {
  var customers = {};

  filterControllerOrdersByDateRange_(orders, startDate, endDate).forEach(function(order) {
    var customerKey = getControllerCustomerKey_(order);

    if (customerKey) {
      customers[customerKey] = true;
    }
  });

  return Object.keys(customers).length;
}

function buildControllerQtyTotal_(orders, detailQtyByNoSo, startDate, endDate) {
  return filterControllerOrdersByDateRange_(orders, startDate, endDate).reduce(function(sum, order) {
    var noSo = String(order.no_so || '').trim();
    return sum + Number(detailQtyByNoSo[noSo] || 0);
  }, 0);
}

function filterControllerOrdersByDateRange_(orders, startDate, endDate) {
  var startKey = String(startDate || '').trim();
  var endKey = String(endDate || '').trim();

  return (orders || []).filter(function(order) {
    var orderDate = normalizeSheetDateToYmd_(order.tanggal_order || '');

    if (!orderDate) {
      return false;
    }

    if (startKey && orderDate < startKey) {
      return false;
    }

    if (endKey && orderDate > endKey) {
      return false;
    }

    return true;
  });
}

function getControllerOrderAmount_(order) {
  var source = order || {};
  var totalFinal = Number(source.total_final || 0);
  var totalOrder = Number(source.total || 0);

  if (!isNaN(totalFinal) && totalFinal > 0) {
    return totalFinal;
  }

  if (!isNaN(totalOrder) && totalOrder > 0) {
    return totalOrder;
  }

  return 0;
}

function getControllerCustomerKey_(order) {
  var source = order || {};
  var customerId = String(source.customer_id || '').trim();
  var customerName = String(source.nama_customer_input || '').trim();

  return customerId || customerName;
}

function getControllerStartOfWeek_(dateValue) {
  var date = new Date(dateValue);
  var dayOfWeek;
  var diffToMonday;

  date.setHours(0, 0, 0, 0);
  dayOfWeek = date.getDay();
  diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  date.setDate(date.getDate() + diffToMonday);
  return date;
}

function buildControllerStatusSummary_(salesOrders) {
  var statusConfig = [
    { key: 'draft', label: 'Draft' },
    { key: 'menunggu persetujuan', label: 'Menunggu Persetujuan' },
    { key: 'disetujui', label: 'Disetujui' },
    { key: 'ditolak', label: 'Ditolak' },
    { key: 'siap kirim', label: 'Siap Kirim' },
    { key: 'terkirim', label: 'Terkirim' },
    { key: 'selesai', label: 'Selesai' }
  ];
  var counts = {};

  (salesOrders || []).forEach(function(order) {
    var statusKey = normalizeText_(order.status_order || '');

    if (!statusKey) {
      return;
    }

    counts[statusKey] = (counts[statusKey] || 0) + 1;
  });

  return statusConfig.map(function(item) {
    return {
      key: item.key.replace(/\s+/g, '_'),
      label: item.label,
      value: Number(counts[item.key] || 0)
    };
  });
}

function buildControllerOpsSummary_(salesOrders, deliveryOrders, billingData) {
  var deliveryRows = Array.isArray(deliveryOrders) ? deliveryOrders : [];
  var billingSummary = billingData && billingData.billingSummary ? billingData.billingSummary : {};

  return {
    ready_orders: countControllerOrdersByStatus_(salesOrders, 'siap kirim'),
    sj_ready: countControllerDeliveryByStatus_(deliveryRows, 'siap kirim'),
    sj_delivered: countControllerDeliveryByStatus_(deliveryRows, 'terkirim'),
    sj_completed: countControllerDeliveryByStatus_(deliveryRows, 'selesai'),
    unpaid: Number(billingSummary.belum_lunas || 0) + Number(billingSummary.sebagian || 0)
  };
}

function countControllerOrdersByStatus_(salesOrders, statusKey) {
  return (salesOrders || []).filter(function(order) {
    return normalizeText_(order.status_order || '') === normalizeText_(statusKey);
  }).length;
}

function countControllerDeliveryByStatus_(deliveryOrders, statusKey) {
  return (deliveryOrders || []).filter(function(row) {
    return normalizeText_(row.status_kirim || '') === normalizeText_(statusKey);
  }).length;
}

function buildControllerTrend7Days_(salesOrders, todayKey) {
  var endDate = parseControllerDateKey_(todayKey) || getNowParts_().timestamp;
  var rows = [];
  var countsByDate = {};
  var omzetByDate = {};
  var cursor;
  var dateKey;

  (salesOrders || []).forEach(function(order) {
    var orderDate = normalizeSheetDateToYmd_(order.tanggal_order || '');

    if (!orderDate) {
      return;
    }

    countsByDate[orderDate] = (countsByDate[orderDate] || 0) + 1;
    omzetByDate[orderDate] = (omzetByDate[orderDate] || 0) + getControllerOrderAmount_(order);
  });

  for (cursor = 6; cursor >= 0; cursor -= 1) {
    dateKey = normalizeSheetDateToYmd_(shiftControllerDate_(endDate, -cursor));
    rows.push({
      tanggal: dateKey,
      so_count: Number(countsByDate[dateKey] || 0),
      omzet: Number(omzetByDate[dateKey] || 0)
    });
  }

  return rows;
}

function buildControllerTopSales_(salesOrders) {
  var salesUsers = getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).filter(function(user) {
    return normalizeText_(user.role) === 'sales' &&
      normalizeText_(user.status_aktif || 'aktif') === 'aktif';
  }).map(function(user) {
    var userId = String(user.user_id || '').trim();
    var salesCode = String(user.kode_sales || '').trim();
    var displayName = String(user.nama_user || salesCode || userId || 'Tanpa Sales').trim();

    return {
      sales_id: userId,
      sales_kode: salesCode,
      sales_nama: displayName,
      so_count: 0,
      omzet: 0
    };
  });
  var grouped = salesUsers.reduce(function(result, user) {
    if (!user.sales_id) {
      return result;
    }

    result[user.sales_id] = user;
    return result;
  }, {});

  (salesOrders || []).forEach(function(order) {
    var salesId = String(order.sales_id || '').trim();
    var target = grouped[salesId];

    if (!target) {
      return;
    }

    target.so_count += 1;
    target.omzet += getControllerOrderAmount_(order);
  });

  return Object.keys(grouped).map(function(key) {
    return grouped[key];
  }).sort(function(left, right) {
    if (Number(right.omzet || 0) !== Number(left.omzet || 0)) {
      return Number(right.omzet || 0) - Number(left.omzet || 0);
    }

    if (Number(right.so_count || 0) !== Number(left.so_count || 0)) {
      return Number(right.so_count || 0) - Number(left.so_count || 0);
    }

    return String(left.sales_nama || left.sales_id || '').localeCompare(String(right.sales_nama || right.sales_id || ''), 'id-ID');
  });
}

function buildControllerTopCustomers_(salesOrders) {
  var grouped = {};

  (salesOrders || []).forEach(function(order) {
    var customerKey = getControllerCustomerKey_(order);
    var customerId = String(order.customer_id || '').trim();
    var customerName = String(order.nama_customer_input || '').trim() || customerId || 'Tanpa Customer';

    if (!customerKey) {
      return;
    }

    if (!grouped[customerKey]) {
      grouped[customerKey] = {
        customer_id: customerId,
        nama_customer: customerName,
        so_count: 0,
        omzet: 0
      };
    }

    grouped[customerKey].so_count += 1;
    grouped[customerKey].omzet += getControllerOrderAmount_(order);
  });

  return Object.keys(grouped).map(function(key) {
    return grouped[key];
  }).sort(function(left, right) {
    if (Number(right.omzet || 0) !== Number(left.omzet || 0)) {
      return Number(right.omzet || 0) - Number(left.omzet || 0);
    }

    if (Number(right.so_count || 0) !== Number(left.so_count || 0)) {
      return Number(right.so_count || 0) - Number(left.so_count || 0);
    }

    return String(left.nama_customer || '').localeCompare(String(right.nama_customer || ''), 'id-ID');
  }).slice(0, 5);
}

function buildControllerTopItems_(salesOrderDetails) {
  var grouped = {};

  (salesOrderDetails || []).forEach(function(detail) {
    var itemName = String(detail.nama_item || '').trim();
    var noSo = String(detail.no_so || '').trim();

    if (!itemName) {
      return;
    }

    if (!grouped[itemName]) {
      grouped[itemName] = {
        nama_item: itemName,
        qty: 0,
        order_map: {}
      };
    }

    grouped[itemName].qty += Number(detail.qty || 0);
    if (noSo) {
      grouped[itemName].order_map[noSo] = true;
    }
  });

  return Object.keys(grouped).map(function(key) {
    var item = grouped[key];
    return {
      nama_item: item.nama_item,
      qty: Number(item.qty || 0),
      order_count: Object.keys(item.order_map || {}).length
    };
  }).sort(function(left, right) {
    if (Number(right.qty || 0) !== Number(left.qty || 0)) {
      return Number(right.qty || 0) - Number(left.qty || 0);
    }

    if (Number(right.order_count || 0) !== Number(left.order_count || 0)) {
      return Number(right.order_count || 0) - Number(left.order_count || 0);
    }

    return String(left.nama_item || '').localeCompare(String(right.nama_item || ''), 'id-ID');
  }).slice(0, 5);
}

function parseControllerDateKey_(value) {
  var match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    return null;
  }

  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function shiftControllerDate_(dateValue, dayDelta) {
  var date = new Date(dateValue);
  date.setDate(date.getDate() + Number(dayDelta || 0));
  return date;
}
