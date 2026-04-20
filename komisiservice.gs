function getMasterKomisiSlf_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, APP_CONFIG.HEADERS.MASTER_KOMISI_SLF);
  return getSheetData_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF);
}

function findActiveSlfCommissionRate_(jenisKomisi, kodeItem, tanggalRef) {
  if (typeof getActiveCommissionRate_ === 'function') {
    return getActiveCommissionRate_(jenisKomisi, kodeItem, tanggalRef);
  }

  var safeJenisKomisi = String(jenisKomisi || '').trim().toUpperCase();
  var safeKodeItem = String(kodeItem || '').trim().toUpperCase();
  var refDate = normalizeSheetDateToYmd_(tanggalRef);

  if (!safeJenisKomisi || !safeKodeItem) {
    return null;
  }

  return getMasterKomisiSlf_().find(function(row) {
    var statusAktif = normalizeText_(row.status_aktif || 'aktif');
    var rowJenisKomisi = String(row.jenis_komisi || row.nama_skema_komisi || '').trim().toUpperCase();
    var rowKodeItem = String(row.kode_item || row.nama_item || row.produk_komisi || '').trim().toUpperCase();
    var tanggalMulai = normalizeSheetDateToYmd_(row.tanggal_berlaku_mulai || row.tanggal_mulai);
    var tanggalSampai = normalizeSheetDateToYmd_(row.tanggal_berlaku_sampai || row.tanggal_berakhir);

    if (statusAktif !== 'aktif') {
      return false;
    }

    if (rowJenisKomisi !== safeJenisKomisi) {
      return false;
    }

    if (rowKodeItem !== safeKodeItem) {
      return false;
    }

    if (tanggalMulai && refDate && refDate < tanggalMulai) {
      return false;
    }

    if (tanggalSampai && refDate && refDate > tanggalSampai) {
      return false;
    }

    return true;
  }) || null;
}

function validateSlfCommissionPayload_(payload) {
  var safePayload = payload || {};
  var requiredFields = [
    { key: 'jenis_komisi', label: 'Jenis komisi' },
    { key: 'kode_item', label: 'Kode item' },
    { key: 'nama_item', label: 'Nama item' },
    { key: 'satuan_komisi', label: 'Satuan komisi' },
    { key: 'nominal_komisi', label: 'Nominal komisi' },
    { key: 'tanggal_berlaku_mulai', label: 'Tanggal berlaku mulai' },
    { key: 'status_aktif', label: 'Status aktif' }
  ];

  requiredFields.forEach(function(field) {
    if (!String(safePayload[field.key] || '').trim()) {
      throw new Error(field.label + ' wajib diisi.');
    }
  });

  if (!(Number(safePayload.nominal_komisi || 0) >= 0)) {
    throw new Error('Nominal komisi tidak valid.');
  }
}

function getApproverCommissionSlfData(userId) {
  requireCurrentUserRole_(['Approver'], userId);

  return {
    products: getProductCatalog_(),
    master_komisi_slf: getMasterKomisiSlf_()
  };
}

function getApproverSlfCommissionReadyToPay(userId) {
  requireCurrentUserRole_(['Approver'], userId);

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);

  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(row) {
    return normalizeText_(row.channel_sales) === 'slf' &&
      normalizeText_(row.status_komisi) === 'siap cair';
  }).map(function(row) {
    return {
      no_so: row.no_so || '',
      tanggal_order: row.tanggal_order || '',
      sales_id: row.sales_id || '',
      sales_nama: row.sales_nama || '',
      customer_id: row.customer_id || '',
      nama_customer: row.nama_customer_input || row.nama_customer || '',
      jenis_komisi_order: row.jenis_komisi_order || '',
      estimasi_komisi: Number(row.estimasi_komisi || 0),
      komisi_realisasi: Number(row.komisi_realisasi || 0),
      status_komisi: row.status_komisi || '',
      tanggal_siap_cair: row.tanggal_siap_cair || ''
    };
  });
}

function seedDefaultSlfCommissionMaster(userId, options) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  var safeOptions = options || {};
  var defaultNominal = Number(safeOptions.default_nominal || 1000);
  var defaultJenisKomisi = String(safeOptions.jenis_komisi || 'BARU').trim().toUpperCase() || 'BARU';
  var defaultTanggalMulai = String(safeOptions.tanggal_berlaku_mulai || getNowParts_().tanggal).trim();
  var defaultStatusAktif = String(safeOptions.status_aktif || 'Aktif').trim() || 'Aktif';
  var now = getNowParts_();
  var products;
  var existingMap = {};
  var createdRows = [];

  if (!(defaultNominal >= 0)) {
    throw new Error('Default nominal komisi tidak valid.');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, APP_CONFIG.HEADERS.MASTER_KOMISI_SLF);
  products = getProductCatalog_();

  getMasterKomisiSlf_().forEach(function(row) {
    var key = [
      String(row.jenis_komisi || row.nama_skema_komisi || '').trim().toUpperCase(),
      String(row.kode_item || row.produk_komisi || '').trim().toUpperCase()
    ].join('|');

    if (key !== '|') {
      existingMap[key] = true;
    }
  });

  products.forEach(function(product) {
    var kodeItem = String(product.kode_item || '').trim().toUpperCase();
    var key = [defaultJenisKomisi, kodeItem].join('|');
    var rowPayload;

    if (!kodeItem || existingMap[key]) {
      return;
    }

    rowPayload = {
      komisi_id: generateDocNumber_('KMS'),
      nama_skema_komisi: defaultJenisKomisi,
      jenis_komisi: defaultJenisKomisi,
      tipe_sales: 'Freelance',
      channel_sales: 'SLF',
      kode_item: kodeItem,
      nama_item: String(product.nama_item || '').trim(),
      produk_komisi: kodeItem,
      satuan_komisi: String(product.satuan || '').trim(),
      qty_min: '',
      qty_max: '',
      tarif_komisi_per_unit: defaultNominal,
      nominal_komisi: defaultNominal,
      tanggal_mulai: defaultTanggalMulai,
      tanggal_berakhir: '',
      tanggal_berlaku_mulai: defaultTanggalMulai,
      tanggal_berlaku_sampai: '',
      status_aktif: defaultStatusAktif,
      diinput_oleh: currentUser.user_id || '',
      tanggal_input: now.tanggal + ' ' + now.jam,
      catatan: 'Auto-seed master komisi SLF'
    };

    appendRowByHeaders_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, rowPayload);
    existingMap[key] = true;
    createdRows.push({
      komisi_id: rowPayload.komisi_id,
      jenis_komisi: rowPayload.jenis_komisi,
      kode_item: rowPayload.kode_item,
      nama_item: rowPayload.nama_item,
      nominal_komisi: rowPayload.nominal_komisi
    });
    Utilities.sleep(50);
  });

  return {
    success: true,
    message: createdRows.length
      ? 'Master komisi SLF default berhasil dibuat.'
      : 'Tidak ada row baru. Semua item aktif sudah punya master komisi SLF.',
    created_count: createdRows.length,
    created_rows: createdRows
  };
}

function seedDefaultSlfCommissionMasterNow() {
  var currentUser = getCurrentUserProfile();
  return seedDefaultSlfCommissionMaster(currentUser.user_id, {
    default_nominal: 1000,
    jenis_komisi: 'BARU',
    status_aktif: 'Aktif'
  });
}

function seedDefaultSlfCommissionMasterUnsafeNow() {
  var safeOptions = {
    default_nominal: 1000,
    jenis_komisi: 'BARU',
    status_aktif: 'Aktif'
  };
  var defaultNominal = Number(safeOptions.default_nominal || 1000);
  var defaultJenisKomisi = String(safeOptions.jenis_komisi || 'BARU').trim().toUpperCase() || 'BARU';
  var defaultTanggalMulai = getNowParts_().tanggal;
  var defaultStatusAktif = String(safeOptions.status_aktif || 'Aktif').trim() || 'Aktif';
  var now = getNowParts_();
  var products;
  var existingMap = {};
  var createdRows = [];

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, APP_CONFIG.HEADERS.MASTER_KOMISI_SLF);
  products = getProductCatalog_();

  getMasterKomisiSlf_().forEach(function(row) {
    var key = [
      String(row.jenis_komisi || row.nama_skema_komisi || '').trim().toUpperCase(),
      String(row.kode_item || row.produk_komisi || '').trim().toUpperCase()
    ].join('|');

    if (key !== '|') {
      existingMap[key] = true;
    }
  });

  products.forEach(function(product) {
    var kodeItem = String(product.kode_item || '').trim().toUpperCase();
    var key = [defaultJenisKomisi, kodeItem].join('|');
    var rowPayload;

    if (!kodeItem || existingMap[key]) {
      return;
    }

    rowPayload = {
      komisi_id: generateDocNumber_('KMS'),
      nama_skema_komisi: defaultJenisKomisi,
      jenis_komisi: defaultJenisKomisi,
      tipe_sales: 'Freelance',
      channel_sales: 'SLF',
      kode_item: kodeItem,
      nama_item: String(product.nama_item || '').trim(),
      produk_komisi: kodeItem,
      satuan_komisi: String(product.satuan || '').trim(),
      qty_min: '',
      qty_max: '',
      tarif_komisi_per_unit: defaultNominal,
      nominal_komisi: defaultNominal,
      tanggal_mulai: defaultTanggalMulai,
      tanggal_berakhir: '',
      tanggal_berlaku_mulai: defaultTanggalMulai,
      tanggal_berlaku_sampai: '',
      status_aktif: defaultStatusAktif,
      diinput_oleh: 'SYSTEM_SETUP',
      tanggal_input: now.tanggal + ' ' + now.jam,
      catatan: 'Auto-seed master komisi SLF (unsafe runner)'
    };

    appendRowByHeaders_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, rowPayload);
    existingMap[key] = true;
    createdRows.push({
      komisi_id: rowPayload.komisi_id,
      jenis_komisi: rowPayload.jenis_komisi,
      kode_item: rowPayload.kode_item,
      nama_item: rowPayload.nama_item,
      nominal_komisi: rowPayload.nominal_komisi
    });
    Utilities.sleep(50);
  });

  return {
    success: true,
    message: createdRows.length
      ? 'Master komisi SLF default berhasil dibuat melalui runner manual.'
      : 'Tidak ada row baru. Semua item aktif sudah punya master komisi SLF.',
    created_count: createdRows.length,
    created_rows: createdRows
  };
}

function upsertSlfCommissionRateFromApprover(userId, payload) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  var safePayload = payload || {};
  var now = getNowParts_();
  var jenisKomisi = String(safePayload.jenis_komisi || '').trim().toUpperCase();
  var kodeItem = String(safePayload.kode_item || '').trim().toUpperCase();
  var key = [jenisKomisi, kodeItem, String(safePayload.tanggal_berlaku_mulai || '').trim()].join('|');
  var existing = null;
  var rowPayload;

  validateSlfCommissionPayload_(safePayload);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, APP_CONFIG.HEADERS.MASTER_KOMISI_SLF);

  existing = getMasterKomisiSlf_().find(function(row) {
    return [
      String(row.jenis_komisi || row.nama_skema_komisi || '').trim().toUpperCase(),
      String(row.kode_item || row.nama_item || row.produk_komisi || '').trim().toUpperCase(),
      String(normalizeSheetDateToYmd_(row.tanggal_berlaku_mulai || row.tanggal_mulai) || '').trim()
    ].join('|') === key;
  }) || null;

  rowPayload = {
    komisi_id: existing && existing.komisi_id ? existing.komisi_id : generateDocNumber_('KMS'),
    nama_skema_komisi: jenisKomisi,
    jenis_komisi: jenisKomisi,
    channel_sales: 'SLF',
    kode_item: kodeItem,
    nama_item: String(safePayload.nama_item || '').trim(),
    produk_komisi: kodeItem,
    satuan_komisi: String(safePayload.satuan_komisi || '').trim(),
    qty_min: safePayload.qty_min || '',
    qty_max: safePayload.qty_max || '',
    tarif_komisi_per_unit: Number(safePayload.nominal_komisi || 0),
    nominal_komisi: Number(safePayload.nominal_komisi || 0),
    tanggal_mulai: String(safePayload.tanggal_berlaku_mulai || '').trim(),
    tanggal_berakhir: String(safePayload.tanggal_berlaku_sampai || '').trim(),
    tanggal_berlaku_mulai: String(safePayload.tanggal_berlaku_mulai || '').trim(),
    tanggal_berlaku_sampai: String(safePayload.tanggal_berlaku_sampai || '').trim(),
    status_aktif: String(safePayload.status_aktif || '').trim(),
    diinput_oleh: currentUser.user_id || '',
    tanggal_input: now.tanggal + ' ' + now.jam,
    catatan: String(safePayload.catatan || '').trim()
  };

  if (existing) {
    updateRowByKey_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, 'komisi_id', existing.komisi_id, rowPayload);
  } else {
    appendRowByHeaders_(APP_CONFIG.SHEETS.MASTER_KOMISI_SLF, rowPayload);
  }

  return {
    success: true,
    komisi_id: rowPayload.komisi_id,
    action: existing ? 'updated' : 'created'
  };
}

function updateSalesOrderCommissionStatus_(noSo, nextStatus, options) {
  var noSoKey = String(noSo || '').trim();
  var statusBaru = String(nextStatus || '').trim();
  var safeOptions = options || {};
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var now = getNowParts_();
  var updates;
  var mutasiSheet;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  if (normalizeText_(salesOrder.channel_sales) !== 'slf') {
    return salesOrder;
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  updates = {
    status_komisi: statusBaru,
    tanggal_status_komisi: now.tanggal,
    catatan_komisi: String(safeOptions.catatan || '').trim() || salesOrder.catatan_komisi || ''
  };

  if (statusBaru === 'Siap Cair') {
    updates.komisi_realisasi = Number(salesOrder.estimasi_komisi || 0);
    updates.tanggal_siap_cair = now.tanggal;
  }

  if (statusBaru === 'Sudah Dibayar') {
    updates.tanggal_bayar_komisi = now.tanggal;
  }

  if (Object.prototype.hasOwnProperty.call(safeOptions, 'komisi_realisasi')) {
    updates.komisi_realisasi = Number(safeOptions.komisi_realisasi || 0);
  }

  salesOrder = updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSoKey, updates);
  mutasiSheet = getSheetByNameOrNull_(APP_CONFIG.SHEETS.KOMISI_SLF_MUTASI);

  if (mutasiSheet) {
    ensureSheetHeadersContain_(APP_CONFIG.SHEETS.KOMISI_SLF_MUTASI, APP_CONFIG.HEADERS.KOMISI_SLF_MUTASI);
    appendRowByHeaders_(APP_CONFIG.SHEETS.KOMISI_SLF_MUTASI, {
      mutasi_komisi_id: generateDocNumber_('KMT'),
      tanggal_mutasi: now.tanggal + ' ' + now.jam,
      no_so: salesOrder.no_so || noSoKey,
      sales_id: salesOrder.sales_id || '',
      sales_nama: salesOrder.sales_nama || '',
      customer_id: salesOrder.customer_id || '',
      nama_customer: salesOrder.nama_customer_input || '',
      tipe_sales: salesOrder.tipe_sales || '',
      channel_sales: salesOrder.channel_sales || '',
      customer_owner_id: salesOrder.customer_owner_id || '',
      customer_owner_nama: salesOrder.customer_owner_nama || '',
      jenis_komisi_order: salesOrder.jenis_komisi_order || '',
      komisi_scheme_source: salesOrder.komisi_scheme_source || '',
      komisi_id_referensi: salesOrder.komisi_id_referensi || '',
      tarif_komisi_per_unit: salesOrder.tarif_komisi_per_unit || '',
      qty_komisi: salesOrder.qty_komisi || 0,
      estimasi_komisi: salesOrder.estimasi_komisi || 0,
      komisi_realisasi: updates.komisi_realisasi !== undefined ? updates.komisi_realisasi : Number(salesOrder.komisi_realisasi || 0),
      status_komisi: statusBaru,
      tanggal_status_komisi: updates.tanggal_status_komisi || '',
      tanggal_siap_cair: updates.tanggal_siap_cair || salesOrder.tanggal_siap_cair || '',
      tanggal_bayar_komisi: updates.tanggal_bayar_komisi || salesOrder.tanggal_bayar_komisi || '',
      catatan_komisi: updates.catatan_komisi || ''
    });
  }

  return salesOrder;
}

function markSlfCommissionPaidFromApprover(userId, payload) {
  requireCurrentUserRole_(['Approver'], userId);
  var safePayload = payload || {};
  var noSo = String(safePayload.no_so || '').trim();

  if (!noSo) {
    throw new Error('No. SO wajib diisi.');
  }

  return updateSalesOrderCommissionStatus_(noSo, 'Sudah Dibayar', {
    catatan: String(safePayload.catatan || '').trim() || 'Komisi dibayar oleh approver'
  });
}
