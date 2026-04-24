function doGet() {
  return HtmlService.createTemplateFromFile('layout')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppBootstrap() {
  ensureMasterUserPasswords_();
  var currentUser = getCurrentUserProfile();

  return {
    appName: APP_CONFIG.APP_NAME,
    currentUser: currentUser,
    brandLogoDataUrl: getBrandLogoDataUrl_(),
    sheets: APP_CONFIG.SHEETS,
    orderStatus: APP_CONFIG.ORDER_STATUS,
    customerStatus: APP_CONFIG.CUSTOMER_STATUS,
    approvalStatus: APP_CONFIG.APPROVAL_STATUS,
    deliveryPriority: APP_CONFIG.DELIVERY_PRIORITY,
    printConfig: APP_CONFIG.PRINT
  };
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .getContent();
}

function toClientValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP_CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  }

  if (Array.isArray(value)) {
    return value.map(function(item) {
      return toClientValue_(item);
    });
  }

  if (value && typeof value === 'object') {
    var result = {};
    Object.keys(value).forEach(function(key) {
      result[key] = toClientValue_(value[key]);
    });
    return result;
  }

  return value;
}

function getSalesOrderFormData(userId) {
  requireCurrentUserRole_(['Sales'], userId);
  syncCompletedSlfCashOrdersToReadyCommission_();
  var currentUser = getCurrentUserProfile(userId);
  var customers = currentUser.is_freelance
    ? getCustomersOwnedBySales_(currentUser.user_id).filter(function(customer) {
      return normalizeText_(customer.sales_owner_id) === normalizeText_(currentUser.user_id) &&
        normalizeText_(customer.status_customer) !== 'ditahan';
    })
    : getActiveCustomers();

  return toClientValue_({
    customers: customers,
    currentUser: currentUser,
    users: [currentUser],
    products: getProductCatalog_(),
    salesHistoryOrders: getSalesOrderHistoryForSales_(currentUser.user_id, 30),
    salesPayoutBatches: listSlfPayoutBatchesForSales_(currentUser.user_id),
    slfMinPayout: getSlfMinPayoutAmount_(),
    deliveryPriority: APP_CONFIG.DELIVERY_PRIORITY,
    customerType: APP_CONFIG.CUSTOMER_TYPE
  });
}

function getSalesOrderHistoryForSales_(salesId, limit) {
  var safeSalesId = String(salesId || '').trim();
  var maxRows = Math.max(1, Number(limit || 30));
  var salesOrders;
  var neededNoSo = {};
  var salesOrderDetailsByNoSo;

  if (!safeSalesId) {
    return [];
  }

  salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(row) {
    return String(row.sales_id || '').trim() === safeSalesId;
  });

  salesOrders.sort(function(left, right) {
    return getSalesOrderHistorySortValue_(right) - getSalesOrderHistorySortValue_(left);
  });

  salesOrders = salesOrders.slice(0, maxRows);
  salesOrders.forEach(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    if (noSoKey) {
      neededNoSo[noSoKey] = true;
    }
  });

  salesOrderDetailsByNoSo = getSalesOrderDetailsMapForNoSo_(neededNoSo);

  return salesOrders.map(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    return buildSalesOrderClientRowFromDetails_(order, salesOrderDetailsByNoSo[noSoKey] || null);
  });
}

function getSalesOrderHistorySortValue_(order) {
  var source = order || {};
  var candidates = [
    source.tanggal_order,
    source.tanggal_kirim_rencana,
    source.tanggal_selesai,
    source.no_so
  ];
  var index;
  var candidate;
  var parsed;

  for (index = 0; index < candidates.length; index += 1) {
    candidate = candidates[index];

    if (candidate instanceof Date && !isNaN(candidate.getTime())) {
      return candidate.getTime();
    }

    parsed = new Date(candidate);
    if (!isNaN(parsed.getTime())) {
      return parsed.getTime();
    }
  }

  return 0;
}

function loginFromDashboard(formData) {
  var payload = formData || {};

  return toClientValue_(loginWithPassword(payload.user_id || '', payload.password || ''));
}

function getApproverDashboardData(userId) {
  requireCurrentUserRole_(['Approver'], userId);
  maybeBackfillCompletedOrdersVerification_();
  syncCompletedSlfCashOrdersToReadyCommission_();
  var currentUser = getCurrentUserProfile(userId);
  var backupData = getBackupDashboardData_();
  var approvals = getSheetData_(APP_CONFIG.SHEETS.APPROVAL_ORDER).filter(function(row) {
    return normalizeText_(row.status_approval) === 'menunggu';
  });
  var exportOrders = getReadyKledoExportOrders_();
  var approverCommissionData = getApproverCommissionSlfData(userId);

  return toClientValue_({
    currentUser: currentUser,
    approvers: [currentUser],
    products: getProductCatalog_(),
    approverCommissionMaster: approverCommissionData.master_komisi_slf || [],
    approverCommissionReadyToPay: approverCommissionData.ready_to_pay || [],
    exportOrders: exportOrders,
    backupSummary: backupData.summary,
    backupHistory: backupData.history,
    approvals: approvals.map(function(approval) {
      var order = buildSalesOrderClientRow_(findSalesOrderByNoSo_(approval.no_so) || {});
      var customer = findCustomerByCode_(order.customer_id);

      return {
        no_so: approval.no_so,
        approval_id: approval.approval_id,
        tanggal_pengajuan: approval.tanggal_pengajuan,
        diajukan_oleh: approval.diajukan_oleh,
        alasan_approval: approval.alasan_approval,
        status_approval: approval.status_approval,
        customer: order.nama_customer_input || '',
        alamat_kirim: order.alamat_kirim || '',
        item: order.item_summary || order.item || '',
        qty: order.qty_summary || order.qty || '',
        details: order.details || [],
        total: order.total_order || order.total || '',
        sales_nama: order.sales_nama || '',
        status_order: order.status_order || '',
        status_customer: customer ? (customer.status_customer || '') : '',
        status_pembayaran_customer: order.status_pembayaran_customer || '',
        total_tunggakan: order.total_tunggakan || '',
        jumlah_nota_overdue: order.jumlah_nota_overdue || '',
        tanggal_jatuh_tempo_terdekat: order.tanggal_jatuh_tempo_terdekat || '',
        catatan_piutang: order.catatan_piutang || ''
      };
    })
  });
}

function getReadyKledoExportOrders_() {
  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(row) {
    var statusOrder = normalizeText_(row.status_order);
    var verificationStatus = String(row.status_verifikasi_cs || '').trim();
    var exportStatus = String(row.status_export_kledo || '').trim();

    return statusOrder === 'selesai' &&
      verificationStatus === 'Sudah Dicek' &&
      exportStatus !== 'Sudah Export';
  }).map(function(orderRow) {
    var order = buildSalesOrderClientRow_(orderRow || {});
    var suratJalan = findSuratJalanByNoSo_(order.no_so) || {};

    return {
      no_so: order.no_so || '',
      customer: order.nama_customer_input || '',
      alamat_kirim: order.alamat_kirim || '',
      no_hp_customer: order.no_hp_customer || '',
      sales_nama: order.sales_nama || '',
      tanggal_order: order.tanggal_order || '',
      tanggal_selesai: order.tanggal_selesai || '',
      tanggal_verifikasi_cs: order.tanggal_verifikasi_cs || '',
      tanggal_jatuh_tempo: order.tanggal_jatuh_tempo || '',
      tanggal_kirim: suratJalan.tanggal_kirim || order.tanggal_kirim_rencana || '',
      term_pembayaran: order.term_pembayaran || '',
      item: order.item_summary || order.item || '',
      qty: order.qty_summary || order.qty || '',
      details: order.details || [],
      total_final: order.total_final || order.total || 0,
      status_export_kledo: order.status_export_kledo || 'Siap Export',
      catatan_order: order.catatan || '',
      catatan_verifikasi_cs: order.catatan_verifikasi_cs || ''
    };
  });
}

function backfillCompletedOrdersVerification_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  var salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER);
  var suratJalanByNoSo = {};

  getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN).forEach(function(row) {
    var noSoKey = String(row.no_so || '').trim();

    if (noSoKey) {
      suratJalanByNoSo[noSoKey] = row;
    }
  });

  salesOrders.forEach(function(order) {
    var noSo = String(order.no_so || '').trim();
    var statusOrder = normalizeText_(order.status_order);
    var statusVerifikasi = String(order.status_verifikasi_cs || '').trim();
    var statusExport = String(order.status_export_kledo || '').trim();
    var tanggalSelesai = normalizeSheetDateToYmd_(order.tanggal_selesai);
    var suratJalan = suratJalanByNoSo[noSo] || {};
    var updates = {};

    if (!noSo || statusOrder !== 'selesai') {
      return;
    }

    if (!statusVerifikasi) {
      updates.status_verifikasi_cs = 'Sudah Dicek';
      updates.catatan_verifikasi_cs = 'Backfill otomatis untuk order selesai sebelum verifikasi CS diwajibkan.';
    }

    if (!statusExport) {
      updates.status_export_kledo = 'Siap Export';
    }

    if (!tanggalSelesai) {
      updates.tanggal_selesai = normalizeSheetDateToYmd_(
        suratJalan.tanggal_kirim ||
        order.tanggal_verifikasi_cs ||
        order.tanggal_order
      );
    }

    if (!Object.keys(updates).length) {
      return;
    }

    updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, updates);
  });
}

function maybeBackfillCompletedOrdersVerification_() {
  // Backfill ini berguna untuk data lama, tapi berat jika dijalankan setiap request.
  // Batasi maksimal 1x per hari (berdasarkan timezone aplikasi).
  var props = PropertiesService.getScriptProperties();
  var todayKey = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  var propKey = 'backfill_completed_verification_last_run';

  try {
    if (props.getProperty(propKey) === todayKey) {
      return;
    }

    backfillCompletedOrdersVerification_();
    props.setProperty(propKey, todayKey);
  } catch (error) {
    // Jangan blokir dashboard jika backfill gagal.
    // Data tetap bisa dipakai; backfill bisa dicoba lagi nanti.
  }
}

function getAdminDashboardData(userId) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  maybeBackfillCompletedOrdersVerification_();
  syncCompletedSlfCashOrdersToReadyCommission_();
  var currentUser = getCurrentUserProfile(userId);
  var salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER);
  var deliveryOrders = getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN);
  var salesOrderDetails = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);
  var salesOrderByNoSo = {};
  var salesOrderDetailsByNoSo = {};
  var suratJalanByNoSo = {};

  salesOrders.forEach(function(order) {
    salesOrderByNoSo[String(order.no_so || '').trim()] = order;
  });

  salesOrderDetails.forEach(function(detail) {
    var noSo = String(detail.no_so || '').trim();

    if (!noSo) {
      return;
    }

    if (!salesOrderDetailsByNoSo[noSo]) {
      salesOrderDetailsByNoSo[noSo] = [];
    }

    salesOrderDetailsByNoSo[noSo].push(detail);
  });

  Object.keys(salesOrderDetailsByNoSo).forEach(function(noSo) {
    salesOrderDetailsByNoSo[noSo].sort(function(left, right) {
      return Number(left.urutan_item || 0) - Number(right.urutan_item || 0);
    });
  });

  deliveryOrders.forEach(function(row) {
    suratJalanByNoSo[String(row.no_so || '').trim()] = true;
  });

  var readyOrders = salesOrders.filter(function(row) {
    var statusOrder = normalizeText_(row.status_order);
    return statusOrder === 'siap kirim' || statusOrder === 'pending kirim';
  }).filter(function(row) {
    return !suratJalanByNoSo[String(row.no_so || '').trim()];
  }).map(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    return buildSalesOrderClientRowFromDetails_(order, salesOrderDetailsByNoSo[noSoKey] || null);
  });
  var tomorrowPlanDate = getTomorrowDateString_();
  var tomorrowOrders = salesOrders.filter(function(row) {
    return shouldIncludeTomorrowOrder_(row, tomorrowPlanDate);
  }).map(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    return buildSalesOrderClientRowFromDetails_(order, salesOrderDetailsByNoSo[noSoKey] || null);
  });

  return toClientValue_({
    currentUser: currentUser,
    admins: [currentUser],
    customers: getActiveCustomers(),
    products: getProductCatalog_(),
    tomorrowPlanDate: tomorrowPlanDate,
    tomorrowOrders: tomorrowOrders,
    readyOrders: readyOrders,
    deliveryOrders: deliveryOrders.map(function(row) {
      var noSoKey = String(row.no_so || '').trim();
      var sourceOrder = salesOrderByNoSo[noSoKey] || {};
      var order = buildSalesOrderClientRowFromDetails_(sourceOrder, salesOrderDetailsByNoSo[noSoKey] || null);
      var result = {};

      Object.keys(row || {}).forEach(function(key) {
        result[key] = row[key];
      });

      result.item_summary = order.item_summary || '';
      result.qty_summary = order.qty_summary || '';
      result.details = order.details || [];
      result.subtotal_final = order.subtotal_final || 0;
      result.diskon_final = order.diskon_final || 0;
      result.total_final = order.total_final || 0;
      result.status_verifikasi_cs = order.status_verifikasi_cs || 'Belum Dicek';
      result.diverifikasi_oleh = order.diverifikasi_oleh || '';
      result.tanggal_verifikasi_cs = order.tanggal_verifikasi_cs || '';
      result.catatan_verifikasi_cs = order.catatan_verifikasi_cs || '';

      return result;
    })
  });
}

function getAdminOperationsData(userId, options) {
  requireCurrentUserRole_(['CS/Admin'], userId);

  var salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER);
  var deliveryOrders = getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN);
  var salesOrderByNoSo = {};
  var suratJalanByNoSo = {};
  var neededNoSo = {};
  var salesOrderDetailsByNoSo;
  var readyOrders;
  var deliverySummary;
  var filterDate;
  var filteredDeliveryOrders;

  salesOrders.forEach(function(order) {
    salesOrderByNoSo[String(order.no_so || '').trim()] = order;
  });

  deliveryOrders.forEach(function(row) {
    suratJalanByNoSo[String(row.no_so || '').trim()] = true;
  });

  deliverySummary = deliveryOrders.reduce(function(result, row) {
    var status = normalizeText_(row.status_kirim);
    if (status === 'siap kirim') result.siap_kirim += 1;
    if (status === 'terkirim') result.terkirim += 1;
    if (status === 'selesai') result.selesai += 1;
    return result;
  }, { siap_kirim: 0, terkirim: 0, selesai: 0 });

  filterDate = normalizeSheetDateToYmd_(options && options.delivery_filter_date);
  filteredDeliveryOrders = filterDate ? deliveryOrders.filter(function(row) {
    return normalizeSheetDateToYmd_(row.tanggal_kirim || row.tanggal_cetak || '') === filterDate;
  }) : deliveryOrders.slice();

  readyOrders = salesOrders.filter(function(row) {
    var statusOrder = normalizeText_(row.status_order);
    return statusOrder === 'siap kirim' || statusOrder === 'pending kirim';
  }).filter(function(row) {
    return !suratJalanByNoSo[String(row.no_so || '').trim()];
  });

  readyOrders.forEach(function(order) {
    var key = String(order.no_so || '').trim();
    if (key) neededNoSo[key] = true;
  });

  salesOrderDetailsByNoSo = getSalesOrderDetailsMapForNoSo_(neededNoSo);

  readyOrders = readyOrders.map(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    return buildAdminReadyOrderListRow_(salesOrderByNoSo[noSoKey] || order, salesOrderDetailsByNoSo[noSoKey] || null);
  });

  return toClientValue_({
    readyOrders: readyOrders,
    deliverySummary: deliverySummary,
    products: getProductCatalog_(),
    deliveryOrders: filteredDeliveryOrders.map(function(row) {
      var noSoKey = String(row.no_so || '').trim();
      return buildAdminDeliveryListRow_(row, salesOrderByNoSo[noSoKey] || {});
    })
  });
}

function buildAdminReadyOrderListRow_(order, rawDetails) {
  var row = buildSalesOrderClientRowFromDetails_(order || {}, rawDetails || null);

  return {
    no_so: row.no_so || '',
    tanggal_order: row.tanggal_order || '',
    jam_order: row.jam_order || '',
    tanggal_kirim_rencana: row.tanggal_kirim_rencana || '',
    nama_customer_input: row.nama_customer_input || '',
    nama_customer: row.nama_customer || '',
    sales_nama: row.sales_nama || '',
    item_summary: row.item_summary || row.item || '',
    qty_summary: row.qty_summary || row.qty || '',
    total_order: row.total_order || row.total || 0,
    total_final: row.total_final || row.total || 0,
    status_order: row.status_order || '',
    catatan: row.catatan || '',
    alasan_hold: row.alasan_hold || '',
    alasan_batal: row.alasan_batal || ''
  };
}

function buildAdminDeliveryListRow_(suratJalan, order) {
  var sourceDelivery = suratJalan || {};
  var sourceOrder = order || {};
  var tanggalKirimEfektif = resolveEffectiveSuratJalanTanggalKirim_(sourceDelivery, sourceOrder);

  return {
    no_surat_jalan: sourceDelivery.no_surat_jalan || '',
    no_so: sourceDelivery.no_so || '',
    tanggal_kirim: tanggalKirimEfektif || sourceDelivery.tanggal_cetak || '',
    tanggal_cetak: sourceDelivery.tanggal_cetak || '',
    nama_customer: sourceDelivery.nama_customer || sourceOrder.nama_customer_input || '',
    nama_customer_input: sourceOrder.nama_customer_input || sourceDelivery.nama_customer || '',
    status_kirim: sourceDelivery.status_kirim || '',
    catatan_kirim: sourceDelivery.catatan_kirim || '',
    status_verifikasi_cs: sourceOrder.status_verifikasi_cs || 'Belum Dicek',
    diverifikasi_oleh: sourceOrder.diverifikasi_oleh || '',
    tanggal_verifikasi_cs: sourceOrder.tanggal_verifikasi_cs || '',
    catatan_verifikasi_cs: sourceOrder.catatan_verifikasi_cs || ''
  };
}

function getDeliveryVerificationData_(noSo) {
  var noSoKey = String(noSo || '').trim();
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var sourceOrder = findSalesOrderByNoSo_(noSoKey) || {};
  var order = buildSalesOrderClientRow_(sourceOrder);
  var tanggalKirimEfektif;

  if (!suratJalan) {
    throw new Error('Data surat jalan tidak ditemukan untuk diverifikasi.');
  }

  tanggalKirimEfektif = resolveEffectiveSuratJalanTanggalKirim_(suratJalan, sourceOrder);

  return {
    no_surat_jalan: suratJalan.no_surat_jalan || '',
    no_so: suratJalan.no_so || noSoKey,
    tanggal_kirim: tanggalKirimEfektif || suratJalan.tanggal_cetak || '',
    tanggal_cetak: suratJalan.tanggal_cetak || '',
    nama_customer: suratJalan.nama_customer || sourceOrder.nama_customer_input || '',
    nama_customer_input: sourceOrder.nama_customer_input || suratJalan.nama_customer || '',
    status_kirim: suratJalan.status_kirim || '',
    catatan_kirim: suratJalan.catatan_kirim || '',
    status_verifikasi_cs: order.status_verifikasi_cs || 'Belum Dicek',
    diverifikasi_oleh: order.diverifikasi_oleh || '',
    tanggal_verifikasi_cs: order.tanggal_verifikasi_cs || '',
    catatan_verifikasi_cs: order.catatan_verifikasi_cs || '',
    details: order.details || []
  };
}

function getAdminBillingData(userId) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  var currentUser = getCurrentUserProfile(userId);
  var payload = getAdminBillingData_(currentUser, {});
  return toClientValue_(payload);
}

function markTagihanLunasFromDashboard(userId, payload) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  var currentUser = getCurrentUserProfile(userId);
  return toClientValue_(markTagihanLunas_(currentUser, payload || {}));
}

function recordTagihanPaymentFromDashboard(userId, payload) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  var currentUser = getCurrentUserProfile(userId);
  return toClientValue_(recordTagihanPayment_(currentUser, payload || {}));
}

function getSalesOrderDetailsMapForNoSo_(neededNoSoMap) {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);
  var needed = neededNoSoMap || {};
  var result = {};

  if (!Object.keys(needed).length) {
    return result;
  }

  if (!sheet) {
    return result;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (!lastRow || lastRow < 2 || !lastCol) {
    return result;
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0] || [];
  var noSoIndex = headers.indexOf('no_so');

  if (noSoIndex === -1) {
    return result;
  }

  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var noSo = String(row[noSoIndex] || '').trim();

    if (!noSo || !needed[noSo]) {
      continue;
    }

    if (!result[noSo]) {
      result[noSo] = [];
    }

    result[noSo].push(mapRowToObject_(headers, row));
  }

  Object.keys(result).forEach(function(noSoKey) {
    result[noSoKey].sort(function(left, right) {
      return Number(left.urutan_item || 0) - Number(right.urutan_item || 0);
    });
  });

  return result;
}

function getAdminPlanningData(userId) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  maybeBackfillCompletedOrdersVerification_();

  var salesOrders = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER);
  var salesOrderDetails = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);
  var salesOrderDetailsByNoSo = {};
  var tomorrowPlanDate = getTomorrowDateString_();

  salesOrderDetails.forEach(function(detail) {
    var noSo = String(detail.no_so || '').trim();

    if (!noSo) {
      return;
    }

    if (!salesOrderDetailsByNoSo[noSo]) {
      salesOrderDetailsByNoSo[noSo] = [];
    }

    salesOrderDetailsByNoSo[noSo].push(detail);
  });

  Object.keys(salesOrderDetailsByNoSo).forEach(function(noSo) {
    salesOrderDetailsByNoSo[noSo].sort(function(left, right) {
      return Number(left.urutan_item || 0) - Number(right.urutan_item || 0);
    });
  });

  var tomorrowOrders = salesOrders.filter(function(row) {
    return shouldIncludeTomorrowOrder_(row, tomorrowPlanDate);
  }).map(function(order) {
    var noSoKey = String(order.no_so || '').trim();
    return buildSalesOrderClientRowFromDetails_(order, salesOrderDetailsByNoSo[noSoKey] || null);
  });

  return toClientValue_({
    tomorrowPlanDate: tomorrowPlanDate,
    tomorrowOrders: tomorrowOrders
  });
}

function getAdminAgentFormData(userId) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  var currentUser = getCurrentUserProfile(userId);

  return toClientValue_({
    currentUser: currentUser,
    customers: getActiveCustomers(),
    products: getProductCatalog_()
  });
}

function getTomorrowDateString_() {
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  return Utilities.formatDate(tomorrow, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function shouldIncludeTomorrowOrder_(row, tomorrowDate) {
  var tanggalRencana = normalizeSheetDateToYmd_(row.tanggal_kirim_rencana);
  var statusOrder = normalizeText_(row.status_order);

  if (tanggalRencana !== String(tomorrowDate || '').trim()) {
    return false;
  }

  return ['draft', 'menunggu persetujuan', 'disetujui', 'siap kirim'].indexOf(statusOrder) !== -1;
}

function submitSalesOrderFromForm(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Sales'], userId);
  var currentUserProfile = getCurrentUserProfile(currentUser.user_id);
  var payload = {
    sales_id: currentUser.user_id,
    sales_nama: currentUser.nama_user,
    tipe_sales: currentUserProfile.tipe_sales,
    channel_sales: currentUserProfile.channel_sales_default,
    is_freelance: currentUserProfile.is_freelance,
    jenis_customer: formData.jenis_customer,
    customer_id: formData.customer_id,
    nama_customer_input: formData.nama_customer_input,
    alamat_kirim: formData.alamat_kirim,
    link_google_maps: formData.link_google_maps,
    latitude: formData.latitude,
    longitude: formData.longitude,
    pic_customer: formData.pic_customer,
    no_hp_customer: formData.no_hp_customer,
    items: Array.isArray(formData.items) ? formData.items : [],
    subtotal: Number(formData.subtotal || 0),
    total: Number(formData.total || 0),
    term_pembayaran: formData.term_pembayaran,
    tanggal_jatuh_tempo: formData.tanggal_jatuh_tempo,
    tanggal_kirim_rencana: formData.tanggal_kirim_rencana,
    catatan: formData.catatan
  };

  validateSalesNewCustomerFields_(payload);

  return submitSalesOrder(payload);
}

function approveOrderFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  return approveOrder(formData.no_so, currentUser.user_id, formData.catatan_approval || '');
}

function rejectOrderFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  return rejectOrder(formData.no_so, currentUser.user_id, formData.catatan_approval || '');
}

function updateProductBasePriceFromApprover(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  var kodeItem = String(formData && formData.kode_item || '').trim();
  var hargaDasar = Number(formData && formData.harga_dasar || 0);

  if (!kodeItem) {
    throw new Error('Kode item wajib dipilih.');
  }

  if (hargaDasar <= 0) {
    throw new Error('Harga dasar harus lebih dari 0.');
  }

  return toClientValue_(updateProductBasePrice_(kodeItem, hargaDasar, currentUser));
}

function getApproverCommissionSlfDataFromDashboard(userId) {
  requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(getApproverCommissionSlfData(userId));
}

function submitSlfPayoutRequestFromDashboard(userId, formData) {
  requireCurrentUserRole_(['Sales'], userId);
  return toClientValue_(createSlfPayoutBatch_(userId, formData || {}));
}

function updateSalesBankAccountFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Sales'], userId);
  var payload = formData || {};
  var bankNama = String(payload.bank_nama || '').trim();
  var bankNoRekening = String(payload.bank_no_rekening || '').trim();
  var bankNamaPemilik = String(payload.bank_nama_pemilik || '').trim();

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_USER, APP_CONFIG.HEADERS.MASTER_USER);

  if (!bankNama || !bankNoRekening || !bankNamaPemilik) {
    throw new Error('Data rekening wajib lengkap: nama bank, nomor rekening, dan nama pemilik rekening.');
  }

  updateRowByKey_(APP_CONFIG.SHEETS.MASTER_USER, 'user_id', currentUser.user_id, {
    bank_nama: bankNama,
    bank_no_rekening: bankNoRekening,
    bank_nama_pemilik: bankNamaPemilik
  });

  return toClientValue_(getCurrentUserProfile(currentUser.user_id));
}

function getSalesSlfPayoutBatchesFromDashboard(userId) {
  var currentUser = requireCurrentUserRole_(['Sales'], userId);
  return toClientValue_({
    min_payout: getSlfMinPayoutAmount_(),
    batches: listSlfPayoutBatchesForSales_(currentUser.user_id),
    ready_rows: getSlfReadyToPayoutOrdersBySales_(currentUser.user_id)
  });
}

function getApproverSlfPayoutBatchesFromDashboard(userId) {
  requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_({
    min_payout: getSlfMinPayoutAmount_(),
    batches: listSlfPayoutBatchesForApprover_()
  });
}

function getSlfPayoutBatchDetailFromDashboard(userId, payoutBatchId) {
  var currentUser = requireCurrentUserRole_(['Sales', 'Approver'], userId);
  var detail = getSlfPayoutBatchDetail_(payoutBatchId);

  if (normalizeRoleKey_(currentUser.role) === 'sales' &&
      String(detail.batch.sales_id || '').trim() !== String(currentUser.user_id || '').trim()) {
    throw new Error('Akses ditolak. Batch payout ini bukan milik Anda.');
  }

  return toClientValue_(detail);
}

function markSlfPayoutBatchPaidFromDashboard(userId, formData) {
  requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(markSlfPayoutBatchPaid_(userId, formData || {}));
}

function markSlfPayoutBatchInProcessFromDashboard(userId, formData) {
  requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(markSlfPayoutBatchInProcess_(userId, formData || {}));
}

function createSuratJalanFromDashboard(userId, formData) {
  requireCurrentUserRole_(['CS/Admin'], userId);

  return createSuratJalan(formData.no_so, {
    driver: formData.driver || '',
    armada: formData.armada || '',
    catatan_kirim: formData.catatan_kirim || ''
  });
}

function getSalesOrderRevisionDataFromDashboard(userId, noSo) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  return toClientValue_(getSalesOrderForRevision_(noSo));
}

function saveSalesOrderRevisionFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);
  var payload = formData || {};

  return toClientValue_(reviseSalesOrderByCs_(payload.no_so, currentUser, payload));
}

function cancelSalesOrderFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);
  var payload = formData || {};

  return toClientValue_(cancelSalesOrderByCs_(payload.no_so, currentUser, payload.alasan_batal || ''));
}

function markOrderDeliveredFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);
  return markOrderDelivered(formData.no_so, currentUser.user_id, formData.catatan_kirim || '');
}

function getDeliveryVerificationDataFromDashboard(userId, noSo) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  return toClientValue_(getDeliveryVerificationData_(noSo));
}

function verifyDeliveredOrderFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);

  return verifyDeliveredOrder(formData.no_so, currentUser.user_id, {
    items: Array.isArray(formData.items) ? formData.items : [],
    catatan_verifikasi_cs: formData.catatan_verifikasi_cs || ''
  });
}

function generateKledoExportFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(generateKledoExportBatchFile(currentUser));
}

function markKledoExportedFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(markKledoBatchExported(currentUser, formData.catatan_export_kledo || ''));
}

function completeOrderFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);
  return completeOrder(formData.no_so, currentUser.user_id, formData.catatan_kirim || '');
}

function submitAgentOrderFromAdmin(userId, formData) {
  var currentUser = requireCurrentUserRole_(['CS/Admin'], userId);
  var catatan = formData.catatan || '';
  var catatanGabungan = '[AGEN/CS] Input oleh ' + currentUser.nama_user;

  if (catatan) {
    catatanGabungan += ' | ' + catatan;
  }

  return submitSalesOrder({
    sales_id: currentUser.user_id,
    sales_nama: currentUser.nama_user + ' (CS/Admin)',
    jenis_customer: formData.jenis_customer,
    customer_id: formData.customer_id,
    nama_customer_input: formData.nama_customer_input,
    alamat_kirim: formData.alamat_kirim,
    link_google_maps: formData.link_google_maps,
    latitude: Number(formData.latitude || 0) || '',
    longitude: Number(formData.longitude || 0) || '',
    pic_customer: formData.pic_customer,
    no_hp_customer: formData.no_hp_customer,
    items: Array.isArray(formData.items) ? formData.items : [],
    subtotal: Number(formData.subtotal || 0),
    total: Number(formData.total || 0),
    term_pembayaran: formData.term_pembayaran,
    tanggal_jatuh_tempo: formData.tanggal_jatuh_tempo,
    tanggal_kirim_rencana: formData.tanggal_kirim_rencana,
    catatan: catatanGabungan
  });
}

function getSuratJalanPrintDataFromDashboard(userId, noSo) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  return toClientValue_(getSuratJalanPrintData(noSo));
}

function getSuratJalanPreviewDataFromDashboard(userId, noSo) {
  requireCurrentUserRole_(['CS/Admin'], userId);
  return toClientValue_(getSuratJalanPreviewData(noSo));
}

function validateSalesNewCustomerFields_(payload) {
  var requiredFields;

  if (String(payload.jenis_customer || '').trim().toLowerCase() !== 'baru') {
    return;
  }

  requiredFields = [
    { key: 'nama_customer_input', label: 'Nama Customer' },
    { key: 'pic_customer', label: 'PIC Customer' },
    { key: 'alamat_kirim', label: 'Alamat Kirim' },
    { key: 'no_hp_customer', label: 'No. HP Customer' },
    { key: 'link_google_maps', label: 'Link Google Maps' }
  ];

  requiredFields.forEach(function(field) {
    if (!String(payload[field.key] || '').trim()) {
      throw new Error(field.label + ' wajib diisi untuk customer baru.');
    }
  });
}

function getProductCatalog_() {
  ensureDefaultMasterItems_();
  syncMasterItemNames_();

  return getSheetData_(APP_CONFIG.SHEETS.MASTER_ITEM).filter(function(row) {
    return normalizeText_(row.status_aktif || 'aktif') !== 'nonaktif';
  }).map(function(row) {
    var hargaDasar = Number(row.harga_dasar || 0);
    var hargaDefault = Number(row.harga_default || 0) || hargaDasar;

    return {
      kode_item: String(row.kode_item || '').trim(),
      nama_item: String(row.nama_item || '').trim(),
      satuan: String(row.satuan || '').trim(),
      harga_default: hargaDefault,
      harga_dasar: hargaDasar,
      diupdate_oleh: String(row.diupdate_oleh || '').trim(),
      tanggal_update_harga: row.tanggal_update_harga || '',
      status_aktif: row.status_aktif || 'Aktif'
    };
  });
}

function ensureDefaultMasterItems_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_ITEM, APP_CONFIG.HEADERS.MASTER_ITEM);

  if (getSheetData_(APP_CONFIG.SHEETS.MASTER_ITEM).length) {
    return;
  }

  writeRowsByHeaders_(
    APP_CONFIG.SHEETS.MASTER_ITEM,
    APP_CONFIG.HEADERS.MASTER_ITEM,
    getDefaultProductCatalogSeed_()
  );
}

function getDefaultProductCatalogSeed_() {
  return [
    { kode_item: 'PRD001', nama_item: 'AIRTIS Galon 19L', satuan: 'pcs', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD002', nama_item: 'AIRTIS Galon Refill 19L', satuan: 'pcs', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD003', nama_item: 'AIRTIS Cup 150 ml', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD004', nama_item: 'AIRTIS Cup 220ml - 48', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD005', nama_item: 'AIRTIS Botol 220 ml', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD006', nama_item: 'AIRTIS Botol 330ml', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD007', nama_item: 'AIRTIS Botol 600ml', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' },
    { kode_item: 'PRD008', nama_item: 'AIRTIS Botol 1500ml', satuan: 'dus', harga_default: 0, harga_dasar: 0, diupdate_oleh: '', tanggal_update_harga: '', status_aktif: 'Aktif' }
  ];
}

function syncMasterItemNames_() {
  var expectedNamesByCode = {
    PRD002: 'AIRTIS Galon Refill 19L',
    PRD004: 'AIRTIS Cup 220ml - 48',
    PRD006: 'AIRTIS Botol 330ml',
    PRD007: 'AIRTIS Botol 600ml',
    PRD008: 'AIRTIS Botol 1500ml'
  };
  var expectedNamesByLegacyName = {
    'AIRTIS Refill galon 19L': 'AIRTIS Galon Refill 19L',
    'AIRTIS Cup 220 ml': 'AIRTIS Cup 220ml - 48',
    'AIRTIS Botol 330 ml': 'AIRTIS Botol 330ml',
    'AIRTIS Botol 600 ml': 'AIRTIS Botol 600ml',
    'AIRTIS Botol 1500 ml': 'AIRTIS Botol 1500ml'
  };
  var rows = getSheetData_(APP_CONFIG.SHEETS.MASTER_ITEM);

  rows.forEach(function(row) {
    var kodeItem = String(row.kode_item || '').trim();
    var currentName = String(row.nama_item || '').trim();
    var expectedName = expectedNamesByCode[kodeItem] || expectedNamesByLegacyName[currentName];

    if (!expectedName || currentName === expectedName) {
      return;
    }

    updateRowByKey_(APP_CONFIG.SHEETS.MASTER_ITEM, 'kode_item', kodeItem, {
      nama_item: expectedName
    });
  });
}

function updateProductBasePrice_(kodeItem, hargaDasar, currentUser) {
  var now = getNowParts_();
  var userLabel = currentUser && currentUser.user_id ? currentUser.user_id : '';

  return updateRowByKey_(APP_CONFIG.SHEETS.MASTER_ITEM, 'kode_item', kodeItem, {
    harga_dasar: Number(hargaDasar || 0),
    harga_default: Number(hargaDasar || 0),
    diupdate_oleh: userLabel,
    tanggal_update_harga: now.tanggal + ' ' + now.jam
  });
}

function getBrandLogoDataUrl_() {
  return 'data:image/webp;base64,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';
}

// === KPI (Target Sales) ===
function getApproverKpiTargetDataFromDashboard(userId, formData) {
  requireCurrentUserRole_(['Approver'], userId);
  var bulan = String(formData && formData.bulan ? formData.bulan : '').trim();

  return toClientValue_(getApproverKpiProgressData_(bulan));
}

function upsertSalesKpiTargetFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  var bulan = String(formData && formData.bulan ? formData.bulan : '').trim();
  var salesId = String(formData && formData.sales_id ? formData.sales_id : '').trim();
  var targetQty = Number(formData && typeof formData.target_qty !== 'undefined' ? formData.target_qty : 0);
  var catatan = String(formData && formData.catatan ? formData.catatan : '').trim();
  var result;

  if (salesId === '__ALL__') {
    var results = upsertSalesKpiTargetForAllSales_(bulan, targetQty, catatan, currentUser.user_id);
    result = {
      success: true,
      mode: 'all',
      updated_count: results.length
    };
  } else {
    upsertSalesKpiTarget_(bulan, salesId, targetQty, catatan, currentUser.user_id);
    result = {
      success: true,
      mode: 'single',
      updated_count: 1
    };
  }

  result.bulan = bulan;
  result.targets = listSalesKpiTargetsByMonth_(bulan);
  result.progressData = getApproverKpiProgressData_(bulan);

  return toClientValue_(result);
}

function getSalesKpiSummaryFromDashboard(userId, formData) {
  var currentUser = requireCurrentUserRole_(['Sales'], userId);
  var bulan = String(formData && formData.bulan ? formData.bulan : '').trim();

  return toClientValue_(getSalesKpiSummary_(bulan, currentUser.user_id));
}

function testGetApproverDashboardData() {
  console.log(JSON.stringify(getApproverDashboardData('APV44')));
}

function testGetAdminDashboardData() {
  console.log(JSON.stringify(getAdminDashboardData('APV44')));
}

function testDebugSlfCommissionReadyByNoSo() {
  console.log(JSON.stringify(debugSlfCommissionReadyByNoSo_('SO-20260420123930')));
}

function testGetApproverCommissionSlfDataFromDashboard() {
  console.log(JSON.stringify(getApproverCommissionSlfDataFromDashboard('APV44')));
}

function testGetApproverCommissionReadySummary() {
  var payload = getApproverCommissionSlfDataFromDashboard('APV44') || {};
  var rows = Array.isArray(payload.ready_to_pay) ? payload.ready_to_pay : [];

  console.log(JSON.stringify({
    ready_to_pay_count: rows.length,
    ready_to_pay_rows: rows
  }));
}
