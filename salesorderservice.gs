function submitSalesOrder(payload) {
  validateSalesOrderPayload_(payload);

  var now = getNowParts_();
  var noSo = generateDocNumber_(APP_CONFIG.DOC_PREFIX.SALES_ORDER);
  var items = normalizeOrderItems_(payload);
  var commissionInfo;
  var customerCheck = buildOrderCustomerCheck_(payload);
  var priceCheck = buildOrderPriceCheck_(items);
  var approvalDecision = mergeOrderApprovalChecks_(customerCheck, priceCheck);
  var totals = calculateOrderTotals_(items);
  var prioritasKirim = resolvePrioritasKirim_(payload.tanggal_kirim_rencana, now.timestamp);

  payload._normalized_items = items;
  validateFreelanceCustomerAccess_(payload);
  validateFreelanceCommissionCoverage_(payload, customerCheck.customer_id, noSo);
  commissionInfo = calculateSalesOrderCommission_(payload, customerCheck.customer_id, noSo);

  var salesOrderRow = {
    no_so: noSo,
    tanggal_order: now.tanggal,
    jam_order: now.jam,
    sales_id: payload.sales_id,
    sales_nama: payload.sales_nama,
    jenis_customer: payload.jenis_customer,
    customer_id: customerCheck.customer_id,
    nama_customer_input: customerCheck.nama_customer_input,
    alamat_kirim: payload.alamat_kirim,
    link_google_maps: payload.link_google_maps || '',
    latitude: payload.latitude || '',
    longitude: payload.longitude || '',
    pic_customer: payload.pic_customer,
    no_hp_customer: payload.no_hp_customer,
    item: buildOrderItemsSummary_(items),
    qty: buildOrderQtyDisplay_(items),
    harga: items.length === 1 ? items[0].harga : '',
    diskon: totals.diskon_order,
    subtotal: totals.subtotal_order,
    total: totals.total_order,
    subtotal_final: totals.subtotal_order,
    diskon_final: totals.diskon_order,
    total_final: totals.total_order,
    term_pembayaran: payload.term_pembayaran,
    tanggal_jatuh_tempo: payload.tanggal_jatuh_tempo,
    status_pembayaran_customer: customerCheck.status_pembayaran_customer,
    total_tunggakan: customerCheck.total_tunggakan,
    jumlah_nota_overdue: customerCheck.jumlah_nota_overdue,
    tanggal_jatuh_tempo_terdekat: customerCheck.tanggal_jatuh_tempo_terdekat,
    catatan_piutang: customerCheck.catatan_piutang,
    status_verifikasi_cs: 'Belum Dicek',
    diverifikasi_oleh: '',
    tanggal_verifikasi_cs: '',
    catatan_verifikasi_cs: '',
    status_export_kledo: 'Belum Siap',
    tanggal_export_kledo: '',
    diekspor_oleh: '',
    catatan_export_kledo: '',
    tanggal_selesai: '',
    status_order: approvalDecision.status_order,
    prioritas_kirim: prioritasKirim,
    tanggal_kirim_rencana: payload.tanggal_kirim_rencana,
    catatan: payload.catatan || '',
    butuh_persetujuan: approvalDecision.butuh_persetujuan,
    alasan_hold: approvalDecision.alasan_hold,
    tipe_sales: String(payload.tipe_sales || '').trim() || (payload.is_freelance ? 'Freelance' : 'Internal'),
    channel_sales: String(payload.channel_sales || '').trim() || (payload.is_freelance ? 'SLF' : 'SLS'),
    customer_owner_id: customerCheck.customer_owner_id || '',
    customer_owner_nama: customerCheck.customer_owner_nama || '',
    jenis_komisi_order: commissionInfo.jenis_komisi_order || '',
    komisi_scheme_source: commissionInfo.komisi_scheme_source || '',
    komisi_id_referensi: commissionInfo.komisi_id_referensi || '',
    tarif_komisi_per_unit: commissionInfo.tarif_komisi_per_unit || '',
    qty_komisi: commissionInfo.qty_komisi || 0,
    estimasi_komisi: commissionInfo.total_estimasi_komisi || 0,
    komisi_realisasi: commissionInfo.komisi_realisasi || 0,
    status_komisi: commissionInfo.status_komisi || '',
    tanggal_status_komisi: commissionInfo.tanggal_status_komisi || '',
    tanggal_siap_cair: commissionInfo.tanggal_siap_cair || '',
    tanggal_bayar_komisi: commissionInfo.tanggal_bayar_komisi || '',
    catatan_komisi: commissionInfo.catatan_komisi || ''
  };

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  appendRowByHeaders_(APP_CONFIG.SHEETS.SALES_ORDER, salesOrderRow);
  writeSalesOrderDetails_(noSo, items);
  logStatusOrder_(noSo, '', salesOrderRow.status_order, payload.sales_id, salesOrderRow.alasan_hold);

  if (salesOrderRow.butuh_persetujuan === 'Ya') {
    createApprovalOrder_(noSo, payload.sales_id, salesOrderRow.alasan_hold);
  }

  if (normalizeText_(salesOrderRow.status_order) === 'siap kirim') {
    try {
      recordKpiLogForOrderIfEligible_(noSo, payload.sales_id, now.tanggal + ' ' + now.jam);
    } catch (error) {
      // KPI bersifat tambahan; jangan bikin submit order gagal.
      console.log('KPI log gagal: ' + (error && error.message ? error.message : error));
    }
  }

  return {
    success: true,
    no_so: noSo,
    customer_id: salesOrderRow.customer_id,
    jumlah_item: items.length,
    status_order: salesOrderRow.status_order,
    tanggal_jatuh_tempo: salesOrderRow.tanggal_jatuh_tempo || '',
    butuh_persetujuan: salesOrderRow.butuh_persetujuan,
    alasan_hold: salesOrderRow.alasan_hold
  };
}

function testSubmitSalesOrderAman() {
  var result = submitSalesOrder({
    sales_id: 'U001',
    sales_nama: 'Andi Sales',
    jenis_customer: 'Lama',
    customer_id: 'CUST001',
    alamat_kirim: 'Jl. Raya Bekasi No. 12, Bekasi',
    link_google_maps: '',
    latitude: '',
    longitude: '',
    pic_customer: 'Pak Joko',
    no_hp_customer: '081300000001',
    item: 'AIRTIS Galon 19L',
    qty: 10,
    harga: 18000,
    diskon: 0,
    subtotal: 180000,
    total: 180000,
    term_pembayaran: 'Cash',
    tanggal_jatuh_tempo: Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    tanggal_kirim_rencana: Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    catatan: 'Test order customer lancar'
  });

  console.log(JSON.stringify(result));
}

function testSubmitSalesOrderHold() {
  var result = submitSalesOrder({
    sales_id: 'U001',
    sales_nama: 'Andi Sales',
    jenis_customer: 'Lama',
    customer_id: 'CUST002',
    alamat_kirim: 'Jl. Industri No. 8, Cikarang',
    link_google_maps: '',
    latitude: '',
    longitude: '',
    pic_customer: 'Ibu Rina',
    no_hp_customer: '081300000002',
    item: 'AIRTIS Galon 19L',
    qty: 20,
    harga: 18000,
    diskon: 0,
    subtotal: 360000,
    total: 360000,
    term_pembayaran: 'Tempo 7 Hari',
    tanggal_jatuh_tempo: Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    tanggal_kirim_rencana: Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    catatan: 'Test order customer menunggak'
  });

  console.log(JSON.stringify(result));
}

function getSalesOrderForRevision_(noSo) {
  var noSoKey = String(noSo || '').trim();
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var orderDisplay;
  var statusOrder;
  var statusKirim;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  statusOrder = normalizeText_(salesOrder.status_order);
  if (['terkirim', 'selesai', 'ditolak'].indexOf(statusOrder) !== -1) {
    throw new Error('Order dengan status ' + salesOrder.status_order + ' tidak bisa direvisi CS.');
  }

  if (suratJalan) {
    statusKirim = normalizeText_(suratJalan.status_kirim);
    if (statusKirim === 'terkirim' || statusKirim === 'selesai') {
      throw new Error('Order tidak bisa direvisi karena surat jalan sudah masuk proses kirim/final.');
    }
  }

  orderDisplay = buildSalesOrderClientRow_(salesOrder);

  return {
    no_so: salesOrder.no_so || '',
    status_order: salesOrder.status_order || '',
    customer_id: salesOrder.customer_id || '',
    nama_customer_input: salesOrder.nama_customer_input || '',
    sales_nama: salesOrder.sales_nama || orderDisplay.sales_nama || '',
    tanggal_order: normalizeSheetDateToYmd_(salesOrder.tanggal_order),
    term_pembayaran: salesOrder.term_pembayaran || '',
    tanggal_jatuh_tempo: normalizeSheetDateToYmd_(salesOrder.tanggal_jatuh_tempo),
    tanggal_kirim_rencana: normalizeSheetDateToYmd_(salesOrder.tanggal_kirim_rencana),
    catatan: salesOrder.catatan || '',
    has_surat_jalan: Boolean(suratJalan),
    status_kirim: suratJalan ? (suratJalan.status_kirim || '') : '',
    details: (orderDisplay.details || []).map(function(detail) {
      return {
        detail_id: detail.detail_id || '',
        kode_item: detail.kode_item || '',
        nama_item: detail.nama_item || '',
        qty: Number(detail.qty || 0),
        satuan: detail.satuan || '',
        harga: Number(detail.harga || 0),
        diskon: Number(detail.diskon || 0),
        subtotal: Number(detail.subtotal || 0)
      };
    })
  };
}

function reviseSalesOrderByCs_(noSo, currentUser, payload) {
  var noSoKey = String(noSo || '').trim();
  var user = currentUser || {};
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var revisionPayload = payload || {};
  var existingDetails;
  var existingDetailMap = {};
  var normalizedItems;
  var totals;
  var previousSnapshot;
  var nextSnapshot;
  var updates;
  var now;
  var pendingApproval;
  var nextStatusOrder;
  var statusKirim;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  if (['terkirim', 'selesai', 'ditolak'].indexOf(normalizeText_(salesOrder.status_order)) !== -1) {
    throw new Error('Order dengan status ' + salesOrder.status_order + ' tidak bisa direvisi CS.');
  }

  if (suratJalan) {
    statusKirim = normalizeText_(suratJalan.status_kirim);
    if (statusKirim === 'terkirim' || statusKirim === 'selesai') {
      throw new Error('Order tidak bisa direvisi karena surat jalan sudah masuk proses kirim/final.');
    }
  }

  validateSalesOrderRevisionPayload_(revisionPayload);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, APP_CONFIG.HEADERS.SALES_ORDER_DETAIL);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.APPROVAL_ORDER, APP_CONFIG.HEADERS.APPROVAL_ORDER);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.LOG_REVISI_ORDER, APP_CONFIG.HEADERS.LOG_REVISI_ORDER);

  existingDetails = getSalesOrderDetailsByNoSo_(noSoKey);
  if (!existingDetails.length) {
    throw new Error('Detail item order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  existingDetails.forEach(function(detail) {
    existingDetailMap[String(detail.detail_id || '').trim()] = detail;
  });

  normalizedItems = normalizeRevisionItems_(revisionPayload.items, existingDetailMap);
  totals = calculateOrderTotals_(normalizedItems);
  previousSnapshot = buildRevisionSnapshot_(salesOrder, existingDetails);
  nextSnapshot = buildRevisionSnapshot_({
    tanggal_kirim_rencana: revisionPayload.tanggal_kirim_rencana,
    term_pembayaran: revisionPayload.term_pembayaran,
    tanggal_jatuh_tempo: revisionPayload.tanggal_jatuh_tempo,
    catatan: revisionPayload.catatan
  }, normalizedItems);
  now = getNowParts_();
  nextStatusOrder = determineRevisedOrderStatus_(salesOrder.status_order, revisionPayload.tanggal_kirim_rencana, now.tanggal);
  replaceSalesOrderDetails_(noSoKey, normalizedItems);

  updates = {
    item: buildOrderItemsSummary_(normalizedItems),
    qty: buildOrderQtyDisplay_(normalizedItems),
    harga: normalizedItems.length === 1 ? normalizedItems[0].harga : '',
    diskon: totals.diskon_order,
    subtotal: totals.subtotal_order,
    total: totals.total_order,
    subtotal_final: totals.subtotal_order,
    diskon_final: totals.diskon_order,
    total_final: totals.total_order,
    tanggal_kirim_rencana: revisionPayload.tanggal_kirim_rencana,
    term_pembayaran: revisionPayload.term_pembayaran,
    tanggal_jatuh_tempo: revisionPayload.tanggal_jatuh_tempo,
    catatan: revisionPayload.catatan,
    status_order: nextStatusOrder,
    butuh_persetujuan: 'Tidak',
    alasan_hold: ''
  };

  updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSoKey, updates);
  syncSuratJalanDraftFromSalesOrder_(noSoKey);

  pendingApproval = findPendingApprovalByNoSo_(noSoKey);
  if (pendingApproval) {
    updateRowByKey_(APP_CONFIG.SHEETS.APPROVAL_ORDER, 'approval_id', pendingApproval.approval_id, {
      status_approval: 'Ditolak',
      diputuskan_oleh: user.user_id || '',
      tanggal_keputusan: now.tanggal + ' ' + now.jam,
      catatan_approval: 'Ditutup otomatis setelah Revisi CS: keputusan manajemen tidak memerlukan approval ulang.'
    });
  }

  appendRowByHeaders_(APP_CONFIG.SHEETS.LOG_REVISI_ORDER, {
    revisi_id: generateDocNumber_('REV'),
    no_so: noSoKey,
    tanggal: now.tanggal,
    jam: now.jam,
    direvisi_oleh: user.user_id || '',
    nama_user_revisi: user.nama_user || '',
    alasan_revisi: String(revisionPayload.alasan_revisi || '').trim(),
    ringkasan_perubahan: buildRevisionSummary_(previousSnapshot, nextSnapshot),
    perubahan_json: JSON.stringify({
      before: previousSnapshot,
      after: nextSnapshot
    })
  });

  if (String(salesOrder.status_order || '').trim() !== String(nextStatusOrder || '').trim()) {
    logStatusOrder_(noSoKey, salesOrder.status_order, nextStatusOrder, user.user_id, 'Revisi CS: ' + String(revisionPayload.alasan_revisi || '').trim());
  }

  return {
    success: true,
    no_so: noSoKey,
    status_order: nextStatusOrder,
    term_pembayaran: revisionPayload.term_pembayaran,
    tanggal_jatuh_tempo: revisionPayload.tanggal_jatuh_tempo,
    total_order: totals.total_order,
    message: 'Revisi CS berhasil disimpan.'
  };
}

function cancelSalesOrderByCs_(noSo, currentUser, reason) {
  var noSoKey = String(noSo || '').trim();
  var user = currentUser || {};
  var cancelReason = String(reason || '').trim();
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var statusOrder;
  var now;
  var pendingApproval;
  var approvalNote;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  if (!cancelReason) {
    throw new Error('Alasan pembatalan wajib diisi.');
  }

  statusOrder = normalizeText_(salesOrder.status_order);
  if (['menunggu persetujuan', 'disetujui', 'siap kirim'].indexOf(statusOrder) === -1) {
    throw new Error('Order dengan status ' + (salesOrder.status_order || '-') + ' tidak bisa dibatalkan.');
  }

  if (suratJalan) {
    throw new Error('Order tidak bisa dibatalkan karena surat jalan sudah dibuat.');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.APPROVAL_ORDER, APP_CONFIG.HEADERS.APPROVAL_ORDER);
  now = getNowParts_();

  updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSoKey, {
    status_order: 'Dibatalkan',
    butuh_persetujuan: 'Tidak',
    alasan_batal: cancelReason,
    dibatalkan_oleh: user.user_id || '',
    tanggal_dibatalkan: now.tanggal + ' ' + now.jam
  });

  pendingApproval = findPendingApprovalByNoSo_(noSoKey);
  if (pendingApproval) {
    approvalNote = 'Dibatalkan oleh CS/Admin: ' + cancelReason;
    updateRowByKey_(APP_CONFIG.SHEETS.APPROVAL_ORDER, 'approval_id', pendingApproval.approval_id, {
      status_approval: 'Ditolak',
      diputuskan_oleh: user.user_id || '',
      tanggal_keputusan: now.tanggal + ' ' + now.jam,
      catatan_approval: approvalNote
    });
  }

  logStatusOrder_(noSoKey, salesOrder.status_order, 'Dibatalkan', user.user_id || '', cancelReason);

  return {
    success: true,
    no_so: noSoKey,
    status_order: 'Dibatalkan',
    alasan_batal: cancelReason,
    dibatalkan_oleh: user.user_id || '',
    tanggal_dibatalkan: now.tanggal + ' ' + now.jam,
    message: 'Sales order berhasil dibatalkan.'
  };
}

function validateSalesOrderPayload_(payload) {
  var requiredFields = [
    'sales_id',
    'sales_nama',
    'jenis_customer',
    'alamat_kirim',
    'term_pembayaran',
    'tanggal_jatuh_tempo',
    'tanggal_kirim_rencana'
  ];
  var isNewCustomer = normalizeText_(payload.jenis_customer) === 'baru';

  if (isNewCustomer) {
    requiredFields = requiredFields.concat([
      'nama_customer_input',
      'pic_customer',
      'no_hp_customer'
    ]);
  }

  requiredFields.forEach(function(field) {
    if (payload[field] === undefined || payload[field] === null || payload[field] === '') {
      throw new Error('Field wajib belum diisi: ' + field);
    }
  });

  if (normalizeText_(payload.jenis_customer) === 'lama' && !payload.customer_id) {
    throw new Error('Customer lama wajib memilih customer_id');
  }

  normalizeOrderItems_(payload);
}

function validateSalesOrderRevisionPayload_(payload) {
  var revisionPayload = payload || {};
  var items = Array.isArray(revisionPayload.items) ? revisionPayload.items : [];

  if (!String(revisionPayload.tanggal_kirim_rencana || '').trim()) {
    throw new Error('Tanggal kirim rencana wajib diisi.');
  }

  if (!String(revisionPayload.term_pembayaran || '').trim()) {
    throw new Error('Term pembayaran wajib diisi.');
  }

  if (!String(revisionPayload.tanggal_jatuh_tempo || '').trim()) {
    throw new Error('Tanggal jatuh tempo wajib diisi.');
  }

  if (!String(revisionPayload.alasan_revisi || '').trim()) {
    throw new Error('Alasan revisi wajib diisi.');
  }

  if (!items.length) {
    throw new Error('Minimal satu item revisi wajib diisi.');
  }
}

function buildOrderCustomerCheck_(payload) {
  if (normalizeText_(payload.jenis_customer) === 'baru') {
    var customer = ensureCustomerMasterForNewOrder_(payload);

    return {
      customer_id: customer.kode_customer || '',
      nama_customer_input: customer.nama_customer || payload.nama_customer_input || '',
      status_pembayaran_customer: 'Lancar',
      total_tunggakan: 0,
      jumlah_nota_overdue: 0,
      tanggal_jatuh_tempo_terdekat: '',
      catatan_piutang: '',
      customer_owner_id: customer.sales_owner_id || '',
      customer_owner_nama: customer.sales_owner_nama || '',
      status_order: 'Siap Kirim',
      butuh_persetujuan: 'Tidak',
      alasan_hold: ''
    };
  }

  var eligibility = checkCustomerEligibility(payload.customer_id);
  var customer = eligibility.customer || {};

  return {
    customer_id: payload.customer_id,
    nama_customer_input: customer.nama_customer || payload.nama_customer_input || '',
    status_pembayaran_customer: customer.status_pembayaran || '',
    total_tunggakan: Number(customer.total_tunggakan || 0),
    jumlah_nota_overdue: Number(customer.jumlah_nota_overdue || 0),
    tanggal_jatuh_tempo_terdekat: customer.tanggal_jatuh_tempo_terdekat || '',
    catatan_piutang: customer.catatan_piutang || customer.catatan || '',
    customer_owner_id: customer.sales_owner_id || '',
    customer_owner_nama: customer.sales_owner_nama || '',
    status_order: eligibility.butuh_persetujuan === 'Ya' ? 'Menunggu Persetujuan' : 'Siap Kirim',
    butuh_persetujuan: eligibility.butuh_persetujuan,
    alasan_hold: eligibility.alasan_hold
  };
}

function validateFreelanceCustomerAccess_(payload) {
  var safePayload = payload || {};
  var customer;
  var ownerId;

  if (!safePayload.is_freelance) {
    return null;
  }

  if (normalizeText_(safePayload.jenis_customer) !== 'lama') {
    return null;
  }

  customer = findCustomerByCode_(safePayload.customer_id);
  if (!customer) {
    throw new Error('Customer lama tidak ditemukan.');
  }

  ownerId = String(customer.sales_owner_id || '').trim();
  if (!ownerId || normalizeText_(ownerId) !== normalizeText_(safePayload.sales_id)) {
    throw new Error('Akses customer ditolak. Sales freelance hanya boleh membuat order untuk customer lama yang dimilikinya.');
  }

  return customer;
}

function getCustomerCommissionType_(customerId, excludeNoSo) {
  var safeCustomerId = String(customerId || '').trim();
  var excludedNoSo = String(excludeNoSo || '').trim();
  var hasCompletedOrder;

  if (!safeCustomerId) {
    return 'BARU';
  }

  hasCompletedOrder = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).some(function(row) {
    var noSo = String(row.no_so || '').trim();

    if (!noSo || (excludedNoSo && noSo === excludedNoSo)) {
      return false;
    }

    return String(row.customer_id || '').trim() === safeCustomerId &&
      normalizeText_(row.status_order) === 'selesai';
  });

  return hasCompletedOrder ? 'REPEAT' : 'BARU';
}

function getActiveCommissionRate_(jenisKomisi, kodeItem, tanggalRef) {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF);
  var safeJenisKomisi = String(jenisKomisi || '').trim().toUpperCase();
  var safeKodeItem = String(kodeItem || '').trim().toUpperCase();
  var refDate = normalizeSheetDateToYmd_(tanggalRef);

  if (!sheet || !safeJenisKomisi || !safeKodeItem) {
    return null;
  }

  return getSheetData_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF).find(function(row) {
    var statusAktif = normalizeText_(row.status_aktif || 'aktif');
    var rowJenisKomisi = resolveCommissionJenisKomisi_(row);
    var rowKodeItem = resolveCommissionKodeItem_(row);

    if (statusAktif !== 'aktif') {
      return false;
    }

    if (rowJenisKomisi !== safeJenisKomisi) {
      return false;
    }

    if (rowKodeItem !== safeKodeItem) {
      return false;
    }

    return isCommissionDateActive_(refDate, row.tanggal_mulai, row.tanggal_berakhir);
  }) || null;
}

function calculateSalesOrderCommission_(payload, customerId, excludeNoSo) {
  var safePayload = payload || {};
  var result = {
    jenis_komisi_order: '',
    total_estimasi_komisi: 0,
    qty_komisi: 0,
    lines: [],
    komisi_scheme_source: '',
    komisi_id_referensi: '',
    tarif_komisi_per_unit: '',
    komisi_realisasi: 0,
    status_komisi: '',
    tanggal_status_komisi: '',
    tanggal_siap_cair: '',
    tanggal_bayar_komisi: '',
    catatan_komisi: ''
  };
  var items;
  var jenisKomisi;
  var tanggalRef;
  var uniqueKomisiIds = [];
  var uniqueTarif = [];

  if (!safePayload.is_freelance) {
    return result;
  }

  items = Array.isArray(safePayload._normalized_items) ? safePayload._normalized_items : normalizeOrderItems_(safePayload);
  jenisKomisi = getCustomerCommissionType_(customerId, excludeNoSo);
  tanggalRef = safePayload.tanggal_order || safePayload.tanggal_kirim_rencana || new Date();
  result.jenis_komisi_order = jenisKomisi;
  result.komisi_scheme_source = 'MASTER_KOMISI_SLF';
  result.komisi_realisasi = 0;
  result.status_komisi = 'Potensial';
  result.tanggal_status_komisi = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  result.tanggal_siap_cair = '';
  result.tanggal_bayar_komisi = '';

  items.forEach(function(item) {
    var commissionRate = getActiveCommissionRate_(jenisKomisi, item.kode_item, tanggalRef);
    var tarif = Number(commissionRate && commissionRate.tarif_komisi_per_unit || 0);
    var qtyKomisi = Number(item.qty || 0);
    var estimasi = qtyKomisi * tarif;
    var line = {
      kode_item: item.kode_item || '',
      nama_item: item.nama_item || '',
      qty: qtyKomisi,
      jenis_komisi: jenisKomisi,
      komisi_id: commissionRate && String(commissionRate.komisi_id || '').trim() || '',
      tarif_komisi_per_unit: tarif,
      estimasi_komisi: estimasi
    };

    if (commissionRate) {
      result.qty_komisi += qtyKomisi;
      result.total_estimasi_komisi += estimasi;

      if (line.komisi_id && uniqueKomisiIds.indexOf(line.komisi_id) === -1) {
        uniqueKomisiIds.push(line.komisi_id);
      }

      if (String(tarif) && uniqueTarif.indexOf(String(tarif)) === -1) {
        uniqueTarif.push(String(tarif));
      }
    }

    result.lines.push(line);
  });

  result.komisi_id_referensi = uniqueKomisiIds.join(', ');
  result.tarif_komisi_per_unit = uniqueTarif.join(', ');
  result.catatan_komisi = buildSalesOrderCommissionNote_(result.lines);

  return result;
}

function validateFreelanceCommissionCoverage_(payload, customerId, excludeNoSo) {
  var safePayload = payload || {};
  var items;
  var jenisKomisi;
  var tanggalRef;
  var missingItems = [];

  if (!safePayload.is_freelance) {
    return;
  }

  items = Array.isArray(safePayload._normalized_items) ? safePayload._normalized_items : normalizeOrderItems_(safePayload);
  jenisKomisi = getCustomerCommissionType_(customerId, excludeNoSo);
  tanggalRef = safePayload.tanggal_order || safePayload.tanggal_kirim_rencana || new Date();

  items.forEach(function(item) {
    var kodeItem = String(item.kode_item || '').trim().toUpperCase();
    var namaItem = String(item.nama_item || kodeItem || '').trim();
    var commissionRate;

    if (!kodeItem) {
      missingItems.push(namaItem || 'Item tanpa kode');
      return;
    }

    commissionRate = getActiveCommissionRate_(jenisKomisi, kodeItem, tanggalRef);
    if (!commissionRate) {
      missingItems.push((namaItem || kodeItem) + ' (' + kodeItem + ')');
    }
  });

  if (missingItems.length) {
    throw new Error(
      'Master komisi SLF aktif belum tersedia untuk item berikut: ' +
      missingItems.join(', ') +
      '. Hubungi Approver atau jalankan seed master komisi terlebih dahulu.'
    );
  }
}

function resolveCommissionJenisKomisi_(row) {
  var directValue = String(row && row.jenis_komisi || '').trim().toUpperCase();
  var fallbackValue = String(row && row.nama_skema_komisi || '').trim().toUpperCase();

  if (directValue) {
    return directValue;
  }

  if (fallbackValue === 'BARU' || fallbackValue === 'REPEAT') {
    return fallbackValue;
  }

  return '';
}

function resolveCommissionKodeItem_(row) {
  return String(row && (row.kode_item || row.produk_komisi) || '').trim().toUpperCase();
}

function isCommissionDateActive_(tanggalRef, tanggalMulai, tanggalBerakhir) {
  var ref = normalizeSheetDateToYmd_(tanggalRef);
  var start = normalizeSheetDateToYmd_(tanggalMulai);
  var end = normalizeSheetDateToYmd_(tanggalBerakhir);

  if (!ref) {
    return true;
  }

  if (start && ref < start) {
    return false;
  }

  if (end && ref > end) {
    return false;
  }

  return true;
}

function buildSalesOrderCommissionNote_(lines) {
  var baseNote = 'Menunggu progres order dan pembayaran';
  var snapshots = (lines || []).filter(function(line) {
    return Number(line.tarif_komisi_per_unit || 0) > 0;
  }).map(function(line) {
    return [
      line.kode_item || '-',
      line.jenis_komisi || '-',
      'x' + String(Number(line.qty || 0)),
      '@' + String(Number(line.tarif_komisi_per_unit || 0)),
      '=' + String(Number(line.estimasi_komisi || 0))
    ].join(' ');
  });

  if (!snapshots.length) {
    return baseNote;
  }

  return baseNote + ' | Snapshot: ' + snapshots.join('; ');
}

function buildRevisionCustomerCheck_(salesOrder, termPembayaran) {
  var source = salesOrder || {};
  var normalizedTerm = normalizeText_(termPembayaran);
  var customer;
  var eligibility;

  if (normalizeText_(source.jenis_customer) !== 'lama') {
    return {
      status_pembayaran_customer: 'Lancar',
      total_tunggakan: 0,
      jumlah_nota_overdue: 0,
      tanggal_jatuh_tempo_terdekat: '',
      catatan_piutang: '',
      butuh_persetujuan: 'Tidak',
      alasan_hold: ''
    };
  }

  customer = findCustomerByCode_(source.customer_id);
  if (!customer) {
    return {
      status_pembayaran_customer: '',
      total_tunggakan: 0,
      jumlah_nota_overdue: 0,
      tanggal_jatuh_tempo_terdekat: '',
      catatan_piutang: '',
      butuh_persetujuan: 'Ya',
      alasan_hold: 'Customer tidak ditemukan'
    };
  }

  eligibility = normalizedTerm.indexOf('tempo') !== -1
    ? checkCustomerEligibility(source.customer_id)
    : { butuh_persetujuan: 'Tidak', alasan_hold: '' };

  return {
    status_pembayaran_customer: customer.status_pembayaran || '',
    total_tunggakan: Number(customer.total_tunggakan || 0),
    jumlah_nota_overdue: Number(customer.jumlah_nota_overdue || 0),
    tanggal_jatuh_tempo_terdekat: customer.tanggal_jatuh_tempo_terdekat || '',
    catatan_piutang: customer.catatan_piutang || customer.catatan || '',
    butuh_persetujuan: eligibility.butuh_persetujuan || 'Tidak',
    alasan_hold: eligibility.alasan_hold || ''
  };
}

function normalizeOrderItems_(payload) {
  var rawItems = Array.isArray(payload.items) ? payload.items : [];
  var items = rawItems.map(function(item, index) {
    return normalizeOrderItemRow_(item, index);
  }).filter(function(item) {
    return item.nama_item;
  });

  if (!items.length && payload.item) {
    items.push(normalizeOrderItemRow_({
      nama_item: payload.item,
      qty: payload.qty,
      harga: payload.harga,
      diskon: payload.diskon,
      subtotal: payload.total || payload.subtotal
    }, 0));
  }

  if (!items.length) {
    throw new Error('Minimal satu item order wajib diisi.');
  }

  items.forEach(function(item, index) {
    if (!item.nama_item) {
      throw new Error('Nama item pada baris ' + (index + 1) + ' wajib diisi.');
    }
    if (item.qty <= 0) {
      throw new Error('Qty pada baris ' + (index + 1) + ' harus lebih dari 0.');
    }
    if (item.harga < 0) {
      throw new Error('Harga pada baris ' + (index + 1) + ' tidak boleh negatif.');
    }
    if (item.diskon < 0) {
      throw new Error('Diskon pada baris ' + (index + 1) + ' tidak boleh negatif.');
    }
    if (item.subtotal < 0) {
      throw new Error('Subtotal pada baris ' + (index + 1) + ' tidak valid.');
    }
  });

  return items;
}

function buildOrderPriceCheck_(items) {
  var violations = [];

  (items || []).forEach(function(item, index) {
    var product = getProductByNameServer_(item.nama_item);
    var hargaDasar = Number(product.harga_dasar || product.harga_default || 0);

    if (!product.nama_item) {
      throw new Error('Item pada baris ' + (index + 1) + ' tidak ditemukan di MASTER_ITEM.');
    }

    if (hargaDasar <= 0) {
      throw new Error('Harga dasar untuk item ' + item.nama_item + ' belum diinput approver.');
    }

    if (Number(item.harga || 0) < hargaDasar) {
      violations.push(
        'Harga ' + item.nama_item +
        ' di bawah harga dasar Rp ' + formatNumberServer_(hargaDasar) +
        ' (harga order Rp ' + formatNumberServer_(item.harga) + ')'
      );
    }
  });

  return {
    butuh_persetujuan: violations.length ? 'Ya' : 'Tidak',
    alasan_hold: violations.join('. ')
  };
}

function mergeOrderApprovalChecks_(customerCheck, priceCheck) {
  var reasons = [
    String(customerCheck && customerCheck.alasan_hold || '').trim(),
    String(priceCheck && priceCheck.alasan_hold || '').trim()
  ].filter(Boolean);
  var needsApproval = String(customerCheck && customerCheck.butuh_persetujuan || '') === 'Ya' ||
    String(priceCheck && priceCheck.butuh_persetujuan || '') === 'Ya';

  return {
    status_order: needsApproval ? 'Menunggu Persetujuan' : 'Siap Kirim',
    butuh_persetujuan: needsApproval ? 'Ya' : 'Tidak',
    alasan_hold: reasons.join('. ')
  };
}

function normalizeOrderItemRow_(item, index) {
  var normalizedItem = item || {};
  var product = getProductByNameServer_(normalizedItem.nama_item);
  var qty = Number(normalizedItem.qty || 0);
  var harga = Number(normalizedItem.harga || 0);
  var diskon = Number(normalizedItem.diskon || 0);
  var subtotal = Number(normalizedItem.subtotal);

  if (isNaN(subtotal)) {
    subtotal = (qty * harga) - diskon;
  }

  return {
    detail_id: generateDocNumber_('DTL'),
    urutan_item: Number(index || 0) + 1,
    kode_item: product.kode_item || '',
    nama_item: String(normalizedItem.nama_item || '').trim(),
    qty: qty,
    satuan: product.satuan || '',
    harga: harga,
    diskon: diskon,
    subtotal: subtotal > 0 ? subtotal : 0
  };
}

function normalizeRevisionItems_(items, existingDetailMap) {
  var normalizedItems = (items || []).map(function(item, index) {
    var rawItem = item || {};
    var detailId = String(rawItem.detail_id || '').trim();
    var existingDetail = existingDetailMap[detailId] || null;
    var resolvedProduct = resolveRevisionProduct_(rawItem, existingDetail);
    var qty = Number(rawItem.qty || 0);
    var harga = Number(rawItem.harga || 0);
    var diskon = Number(rawItem.diskon || 0);
    var effectiveDetailId = detailId || generateDocNumber_('DTL');
    var orderIndex = Number(existingDetail && existingDetail.urutan_item || index + 1);

    if (!(qty > 0)) {
      throw new Error('Qty item ' + (resolvedProduct.nama_item || rawItem.nama_item || effectiveDetailId) + ' harus lebih dari 0.');
    }

    if (harga < 0) {
      throw new Error('Harga item ' + (resolvedProduct.nama_item || rawItem.nama_item || effectiveDetailId) + ' tidak boleh negatif.');
    }

    if (diskon < 0) {
      throw new Error('Diskon item ' + (resolvedProduct.nama_item || rawItem.nama_item || effectiveDetailId) + ' tidak boleh negatif.');
    }

    return {
      detail_id: effectiveDetailId,
      urutan_item: orderIndex,
      kode_item: resolvedProduct.kode_item || String(rawItem.kode_item || existingDetail && existingDetail.kode_item || '').trim(),
      nama_item: resolvedProduct.nama_item || String(rawItem.nama_item || existingDetail && existingDetail.nama_item || '').trim(),
      qty: qty,
      satuan: resolvedProduct.satuan || String(existingDetail && existingDetail.satuan || '').trim(),
      harga: harga,
      diskon: diskon,
      subtotal: Math.max((qty * harga) - diskon, 0)
    };
  }).filter(function(item) {
    return String(item.nama_item || '').trim();
  });

  if (!normalizedItems.length) {
    throw new Error('Minimal satu item revisi wajib aktif.');
  }

  normalizedItems.forEach(function(item, index) {
    item.urutan_item = index + 1;
  });

  return normalizedItems.sort(function(left, right) {
    return Number(left.urutan_item || 0) - Number(right.urutan_item || 0);
  });
}

function getProductByNameServer_(itemName) {
  var requestedName = String(itemName || '').trim();
  var normalizedRequested = normalizeRevisionProductName_(requestedName);
  var catalog = getProductCatalog_();
  var exactMatch = catalog.find(function(row) {
    return String(row.nama_item || '').trim() === requestedName;
  });

  if (exactMatch) {
    return exactMatch;
  }

  return catalog.find(function(row) {
    return normalizeRevisionProductName_(row.nama_item) === normalizedRequested;
  }) || {};
}

function calculateOrderTotals_(items) {
  return (items || []).reduce(function(result, item) {
    result.subtotal_order += Number(item.qty || 0) * Number(item.harga || 0);
    result.diskon_order += Number(item.diskon || 0);
    result.total_order += Number(item.subtotal || 0);
    return result;
  }, {
    subtotal_order: 0,
    diskon_order: 0,
    total_order: 0
  });
}

function writeSalesOrderDetails_(noSo, items) {
  var headers = APP_CONFIG.HEADERS.SALES_ORDER_DETAIL;

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, headers);
  (items || []).forEach(function(item) {
    appendRowByHeaders_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, {
      detail_id: item.detail_id,
      no_so: noSo,
      urutan_item: item.urutan_item,
      kode_item: item.kode_item,
      nama_item: item.nama_item,
      qty: item.qty,
      satuan: item.satuan,
      harga: item.harga,
      diskon: item.diskon,
      subtotal: item.subtotal,
      qty_terkirim: item.qty,
      harga_final: item.harga,
      diskon_final: item.diskon,
      subtotal_final: item.subtotal
    });
  });
}

function replaceSalesOrderDetails_(noSo, items) {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);
  var safeNoSo = String(noSo || '').trim();
  var rowIndex;

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, APP_CONFIG.HEADERS.SALES_ORDER_DETAIL);

  if (!sheet) {
    writeSalesOrderDetails_(safeNoSo, items);
    return;
  }

  for (rowIndex = sheet.getLastRow(); rowIndex >= 2; rowIndex -= 1) {
    if (String(sheet.getRange(rowIndex, 2).getValue() || '').trim() === safeNoSo) {
      sheet.deleteRow(rowIndex);
    }
  }

  writeSalesOrderDetails_(safeNoSo, items);
}

function getSalesOrderDetailsByNoSo_(noSo) {
  var sheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL);

  if (!sheet) {
    return [];
  }

  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL).filter(function(row) {
    return String(row.no_so || '').trim() === String(noSo || '').trim();
  }).sort(function(left, right) {
    return Number(left.urutan_item || 0) - Number(right.urutan_item || 0);
  });
}

function getSalesOrderDetailsForDisplay_(order) {
  var details = getSalesOrderDetailsByNoSo_(order && order.no_so);

  if (details.length) {
    return details.map(function(detail, index) {
      return {
        detail_id: detail.detail_id || '',
        urutan_item: Number(detail.urutan_item || index + 1),
        kode_item: detail.kode_item || '',
        nama_item: detail.nama_item || '',
        qty: Number(detail.qty || 0),
        qty_terkirim: Number(detail.qty_terkirim || detail.qty || 0),
        satuan: detail.satuan || '',
        harga: Number(detail.harga || 0),
        harga_final: Number(detail.harga_final || detail.harga || 0),
        diskon: Number(detail.diskon || 0),
        diskon_final: Number(detail.diskon_final || detail.diskon || 0),
        subtotal: Number(detail.subtotal || 0),
        subtotal_final: Number(detail.subtotal_final || detail.subtotal || 0)
      };
    });
  }

  if (!order) {
    return [];
  }

  if (!String(order.item || '').trim()) {
    return [];
  }

  return [{
    detail_id: '',
    urutan_item: 1,
    kode_item: '',
    nama_item: String(order.item || '').trim(),
    qty: Number(order.qty || 0),
    qty_terkirim: Number(order.qty || 0),
    satuan: getProductUnitByNameServer_(order.item),
    harga: Number(order.harga || 0),
    harga_final: Number(order.harga || 0),
    diskon: Number(order.diskon || 0),
    diskon_final: Number(order.diskon || 0),
    subtotal: Number(order.total || order.subtotal || 0),
    subtotal_final: Number(order.total || order.subtotal || 0)
  }];
}

function buildSalesOrderDetailsFallback_(order) {
  var source = order || {};

  if (!String(source.item || '').trim()) {
    return [];
  }

  return [{
    detail_id: '',
    urutan_item: 1,
    kode_item: '',
    nama_item: String(source.item || '').trim(),
    qty: Number(source.qty || 0),
    qty_terkirim: Number(source.qty || 0),
    satuan: getProductUnitByNameServer_(source.item),
    harga: Number(source.harga || 0),
    harga_final: Number(source.harga || 0),
    diskon: Number(source.diskon || 0),
    diskon_final: Number(source.diskon || 0),
    subtotal: Number(source.total || source.subtotal || 0),
    subtotal_final: Number(source.total || source.subtotal || 0)
  }];
}

function buildSalesOrderClientRowFromDetails_(order, rawDetails) {
  var source = order || {};
  var details = Array.isArray(rawDetails) ? rawDetails : [];

  if (!details.length) {
    details = buildSalesOrderDetailsFallback_(source);
  }

  var totals = calculateOrderTotals_(details);
  var derivedVerificationStatus = String(source.status_verifikasi_cs || '').trim();
  var derivedExportStatus = String(source.status_export_kledo || '').trim();
  var finalTotals = details.reduce(function(result, detail) {
    result.subtotal_order += Number(detail.qty_terkirim || 0) * Number(detail.harga_final || 0);
    result.diskon_order += Number(detail.diskon_final || 0);
    result.total_order += Number(detail.subtotal_final || 0);
    return result;
  }, {
    subtotal_order: 0,
    diskon_order: 0,
    total_order: 0
  });

  return Object.keys(source).reduce(function(result, key) {
    result[key] = source[key];
    return result;
  }, {
    details: details,
    item_summary: buildOrderItemsSummary_(details) || source.item || '',
    qty_summary: buildOrderQtyDisplay_(details) || source.qty || '',
    jumlah_item: details.length,
    subtotal_order: Number(source.subtotal || totals.subtotal_order || 0),
    diskon_order: Number(source.diskon || totals.diskon_order || 0),
    total_order: Number(source.total || totals.total_order || 0),
    subtotal_final: Number(source.subtotal_final || finalTotals.subtotal_order || 0),
    diskon_final: Number(source.diskon_final || finalTotals.diskon_order || 0),
    total_final: Number(source.total_final || finalTotals.total_order || 0),
    status_verifikasi_cs: derivedVerificationStatus || (normalizeText_(source.status_order) === 'selesai' ? 'Sudah Dicek' : 'Belum Dicek'),
    status_export_kledo: derivedExportStatus || (normalizeText_(source.status_order) === 'selesai' ? 'Siap Export' : 'Belum Siap'),
    tanggal_export_kledo: source.tanggal_export_kledo || '',
    diekspor_oleh: source.diekspor_oleh || '',
    catatan_export_kledo: source.catatan_export_kledo || '',
    tanggal_selesai: source.tanggal_selesai || '',
    diverifikasi_oleh: source.diverifikasi_oleh || '',
    tanggal_verifikasi_cs: source.tanggal_verifikasi_cs || '',
    catatan_verifikasi_cs: source.catatan_verifikasi_cs || ''
  });
}

function buildSalesOrderClientRow_(order) {
  var source = order || {};
  var details = getSalesOrderDetailsForDisplay_(source);
  var totals = calculateOrderTotals_(details);
  var derivedVerificationStatus = String(source.status_verifikasi_cs || '').trim();
  var derivedExportStatus = String(source.status_export_kledo || '').trim();
  var finalTotals = details.reduce(function(result, detail) {
    result.subtotal_order += Number(detail.qty_terkirim || 0) * Number(detail.harga_final || 0);
    result.diskon_order += Number(detail.diskon_final || 0);
    result.total_order += Number(detail.subtotal_final || 0);
    return result;
  }, {
    subtotal_order: 0,
    diskon_order: 0,
    total_order: 0
  });

  return Object.keys(source).reduce(function(result, key) {
    result[key] = source[key];
    return result;
  }, {
    details: details,
    item_summary: buildOrderItemsSummary_(details) || source.item || '',
    qty_summary: buildOrderQtyDisplay_(details) || source.qty || '',
    jumlah_item: details.length,
    subtotal_order: Number(source.subtotal || totals.subtotal_order || 0),
    diskon_order: Number(source.diskon || totals.diskon_order || 0),
    total_order: Number(source.total || totals.total_order || 0),
    subtotal_final: Number(source.subtotal_final || finalTotals.subtotal_order || 0),
    diskon_final: Number(source.diskon_final || finalTotals.diskon_order || 0),
    total_final: Number(source.total_final || finalTotals.total_order || 0),
    status_verifikasi_cs: derivedVerificationStatus || (normalizeText_(source.status_order) === 'selesai' ? 'Sudah Dicek' : 'Belum Dicek'),
    status_export_kledo: derivedExportStatus || (normalizeText_(source.status_order) === 'selesai' ? 'Siap Export' : 'Belum Siap'),
    tanggal_export_kledo: source.tanggal_export_kledo || '',
    diekspor_oleh: source.diekspor_oleh || '',
    catatan_export_kledo: source.catatan_export_kledo || '',
    tanggal_selesai: source.tanggal_selesai || '',
    diverifikasi_oleh: source.diverifikasi_oleh || '',
    tanggal_verifikasi_cs: source.tanggal_verifikasi_cs || '',
    catatan_verifikasi_cs: source.catatan_verifikasi_cs || ''
  });
}

function buildOrderItemsSummary_(items) {
  var safeItems = items || [];
  var names = safeItems.map(function(item) {
    return String(item.nama_item || '').trim();
  }).filter(Boolean);

  if (!names.length) {
    return '';
  }

  if (names.length === 1) {
    return names[0];
  }

  if (names.length === 2) {
    return names.join(', ');
  }

  return names.slice(0, 2).join(', ') + ' +' + (names.length - 2) + ' item';
}

function buildOrderQtyDisplay_(items) {
  var safeItems = items || [];

  if (!safeItems.length) {
    return '';
  }

  if (safeItems.length === 1) {
    return String(safeItems[0].qty || '');
  }

  return safeItems.length + ' item';
}

function buildRevisionSnapshot_(order, details) {
  var source = order || {};
  var safeDetails = Array.isArray(details) ? details : [];

  return {
    tanggal_kirim_rencana: normalizeSheetDateToYmd_(source.tanggal_kirim_rencana),
    term_pembayaran: String(source.term_pembayaran || '').trim(),
    tanggal_jatuh_tempo: normalizeSheetDateToYmd_(source.tanggal_jatuh_tempo),
    catatan: String(source.catatan || '').trim(),
    items: safeDetails.map(function(detail) {
      return {
        detail_id: String(detail.detail_id || '').trim(),
        kode_item: String(detail.kode_item || '').trim(),
        nama_item: String(detail.nama_item || '').trim(),
        qty: Number(detail.qty || 0),
        harga: Number(detail.harga || 0),
        diskon: Number(detail.diskon || 0)
      };
    })
  };
}

function buildRevisionSummary_(beforeSnapshot, afterSnapshot) {
  var beforeData = beforeSnapshot || {};
  var afterData = afterSnapshot || {};
  var lines = [];
  var beforeItemsById = {};
  var afterItemIds = {};

  if (String(beforeData.tanggal_kirim_rencana || '').trim() !== String(afterData.tanggal_kirim_rencana || '').trim()) {
    lines.push('tanggal_kirim_rencana: ' + (beforeData.tanggal_kirim_rencana || '-') + ' -> ' + (afterData.tanggal_kirim_rencana || '-'));
  }

  if (String(beforeData.term_pembayaran || '').trim() !== String(afterData.term_pembayaran || '').trim()) {
    lines.push('term_pembayaran: ' + (beforeData.term_pembayaran || '-') + ' -> ' + (afterData.term_pembayaran || '-'));
  }

  if (String(beforeData.tanggal_jatuh_tempo || '').trim() !== String(afterData.tanggal_jatuh_tempo || '').trim()) {
    lines.push('tanggal_jatuh_tempo: ' + (beforeData.tanggal_jatuh_tempo || '-') + ' -> ' + (afterData.tanggal_jatuh_tempo || '-'));
  }

  if (String(beforeData.catatan || '').trim() !== String(afterData.catatan || '').trim()) {
    lines.push('catatan: ' + (beforeData.catatan || '-') + ' -> ' + (afterData.catatan || '-'));
  }

  (beforeData.items || []).forEach(function(item) {
    beforeItemsById[String(item.detail_id || '').trim()] = item;
  });

  (afterData.items || []).forEach(function(item) {
    var detailId = String(item.detail_id || '').trim();
    var previousItem = beforeItemsById[detailId] || {};
    afterItemIds[detailId] = true;

    if (!beforeItemsById[detailId]) {
      lines.push('item baru ' + (item.nama_item || detailId) + ': qty ' + String(Number(item.qty || 0)));
      return;
    }

    if (Number(previousItem.qty || 0) !== Number(item.qty || 0)) {
      lines.push('qty ' + (item.nama_item || previousItem.nama_item || detailId) + ': ' + String(Number(previousItem.qty || 0)) + ' -> ' + String(Number(item.qty || 0)));
    }

    if (Number(previousItem.harga || 0) !== Number(item.harga || 0)) {
      lines.push('harga ' + (item.nama_item || previousItem.nama_item || detailId) + ': ' + formatNumberServer_(Number(previousItem.harga || 0)) + ' -> ' + formatNumberServer_(Number(item.harga || 0)));
    }

    if (Number(previousItem.diskon || 0) !== Number(item.diskon || 0)) {
      lines.push('diskon ' + (item.nama_item || previousItem.nama_item || detailId) + ': ' + formatNumberServer_(Number(previousItem.diskon || 0)) + ' -> ' + formatNumberServer_(Number(item.diskon || 0)));
    }
  });

  (beforeData.items || []).forEach(function(item) {
    var detailId = String(item.detail_id || '').trim();
    if (!afterItemIds[detailId]) {
      lines.push('item dihapus ' + (item.nama_item || detailId));
    }
  });

  return lines.join(' | ') || 'Tidak ada perubahan terdeteksi';
}

function resolveRevisionProduct_(item, existingDetail) {
  var rawItem = item || {};
  var existing = existingDetail || {};
  var product = getProductByCodeServer_(rawItem.kode_item || existing.kode_item);

  if (product && product.nama_item) {
    return product;
  }

  product = getProductByNameServer_(rawItem.nama_item || existing.nama_item);
  if (product && product.nama_item) {
    return product;
  }

  return {
    kode_item: String(rawItem.kode_item || existing.kode_item || '').trim(),
    nama_item: normalizeRevisionProductLabel_(rawItem.nama_item || existing.nama_item || ''),
    satuan: String(existing.satuan || '').trim()
  };
}

function getProductByCodeServer_(kodeItem) {
  var requestedCode = String(kodeItem || '').trim();

  if (!requestedCode) {
    return {};
  }

  return getProductCatalog_().find(function(row) {
    return String(row.kode_item || '').trim() === requestedCode;
  }) || {};
}

function normalizeRevisionProductLabel_(itemName) {
  var product = getProductByNameServer_(itemName);

  if (product && product.nama_item) {
    return String(product.nama_item || '').trim();
  }

  return String(itemName || '').trim();
}

function normalizeRevisionProductName_(itemName) {
  var legacyAliases = {
    'airtis refill galon 19l': 'airtis galon refill 19l',
    'airtis cup 220 ml': 'airtis cup 220ml - 48',
    'airtis botol 330 ml': 'airtis botol 330ml',
    'airtis botol 600 ml': 'airtis botol 600ml',
    'airtis botol 1500 ml': 'airtis botol 1500ml'
  };
  var normalized = normalizeText_(itemName);

  return legacyAliases[normalized] || normalized;
}

function determineRevisedOrderStatus_(previousStatus, tanggalKirimRencana, todayYmd) {
  return 'Siap Kirim';
}

function getProductUnitByNameServer_(itemName) {
  var product = getProductByNameServer_(itemName);

  return String(product.satuan || '').trim();
}

function resolvePrioritasKirim_(tanggalKirimRencana, currentDate) {
  if (!tanggalKirimRencana) {
    return 'Jadwal Biasa';
  }

  var today = new Date(currentDate);
  today.setHours(0, 0, 0, 0);

  var tanggalKirim = new Date(tanggalKirimRencana);
  tanggalKirim.setHours(0, 0, 0, 0);

  var diffDays = Math.round((tanggalKirim.getTime() - today.getTime()) / 86400000);

  if (diffDays < 0) {
    return 'Backdate/Input Susulan';
  }

  if (diffDays === 0) {
    return 'Same Day Opsional';
  }

  if (diffDays === 1) {
    return 'H-1 Wajib';
  }

  return 'Jadwal Biasa';
}

function createApprovalOrder_(noSo, diajukanOleh, alasanApproval) {
  var now = getNowParts_();

  appendRowByHeaders_(APP_CONFIG.SHEETS.APPROVAL_ORDER, {
    approval_id: generateDocNumber_(APP_CONFIG.DOC_PREFIX.APPROVAL),
    no_so: noSo,
    tanggal_pengajuan: now.tanggal + ' ' + now.jam,
    diajukan_oleh: diajukanOleh,
    alasan_approval: alasanApproval || 'Perlu persetujuan manual',
    status_approval: 'Menunggu',
    diputuskan_oleh: '',
    tanggal_keputusan: '',
    catatan_approval: ''
  });
}

function findPendingApprovalByNoSo_(noSo) {
  return getSheetData_(APP_CONFIG.SHEETS.APPROVAL_ORDER).find(function(row) {
    return String(row.no_so || '').trim() === String(noSo || '').trim() &&
      normalizeText_(row.status_approval) === 'menunggu';
  }) || null;
}

function logStatusOrder_(noSo, statusLama, statusBaru, diubahOleh, catatan) {
  var now = getNowParts_();

  appendRowByHeaders_(APP_CONFIG.SHEETS.LOG_STATUS_ORDER, {
    log_id: generateDocNumber_('LOG'),
    no_so: noSo,
    tanggal: now.tanggal,
    jam: now.jam,
    status_lama: statusLama || '',
    status_baru: statusBaru,
    diubah_oleh: diubahOleh || '',
    catatan: catatan || ''
  });
}
