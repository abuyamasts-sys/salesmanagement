function createSuratJalan(noSo, options) {
  var noSoKey = String(noSo || '').trim();
  var lock = LockService.getScriptLock();
  var salesOrder;
  var orderDisplay;
  var now;
  var payload;
  var noSuratJalan;
  var tanggalKirim;

  lock.waitLock(30000);

  try {
    salesOrder = findSalesOrderByNoSo_(noSoKey);

    if (!salesOrder) {
      throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
    }

    assertSuratJalanApprovalGateOpen_(salesOrder);

    if (normalizeText_(salesOrder.status_order) !== 'siap kirim') {
      throw new Error('Surat jalan hanya bisa dibuat untuk order dengan status Siap Kirim');
    }

    if (findSuratJalanByNoSo_(noSoKey)) {
      throw new Error('Surat jalan sudah pernah dibuat untuk no_so: ' + noSoKey);
    }

    orderDisplay = buildSalesOrderClientRow_(salesOrder);
    now = getNowParts_();
    payload = options || {};
    noSuratJalan = generateDocNumber_(APP_CONFIG.DOC_PREFIX.SURAT_JALAN);
    tanggalKirim = normalizeSheetDateToYmd_(payload.tanggal_kirim || salesOrder.tanggal_kirim_rencana || now.tanggal);

    appendRowByHeaders_(APP_CONFIG.SHEETS.SURAT_JALAN, {
      no_surat_jalan: noSuratJalan,
      no_so: noSoKey,
      tanggal_cetak: now.tanggal + ' ' + now.jam,
      tanggal_kirim: tanggalKirim,
      customer_id: salesOrder.customer_id,
      nama_customer: salesOrder.nama_customer_input,
      alamat_kirim: salesOrder.alamat_kirim,
      item: orderDisplay.item_summary || salesOrder.item,
      qty: orderDisplay.qty_summary || salesOrder.qty,
      driver: payload.driver || '',
      armada: payload.armada || '',
      status_kirim: payload.status_kirim || 'Siap Kirim',
      catatan_kirim: payload.catatan_kirim || '',
      dibatalkan_oleh: '',
      tanggal_batal_kirim: '',
      alasan_batal_kirim: ''
    });
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    no_so: noSoKey,
    no_surat_jalan: noSuratJalan,
    status_kirim: payload.status_kirim || 'Siap Kirim'
  };
}

function getSuratJalanPrintData(noSo) {
  var suratJalan = findSuratJalanByNoSo_(noSo);

  if (!suratJalan) {
    throw new Error('Surat jalan tidak ditemukan untuk no_so: ' + noSo);
  }

  var salesOrder = findSalesOrderByNoSo_(noSo) || {};
  assertSuratJalanApprovalGateOpen_(salesOrder);

  var orderDisplay = buildSalesOrderClientRow_(salesOrder);
  var tanggalKirimEfektif = resolveEffectiveSuratJalanTanggalKirim_(suratJalan, salesOrder);
  var customerContact = resolveSuratJalanCustomerContact_(salesOrder);

  return {
    no_surat_jalan: suratJalan.no_surat_jalan || '',
    no_so: suratJalan.no_so || '',
    tanggal_cetak: suratJalan.tanggal_cetak || '',
    tanggal_kirim: tanggalKirimEfektif,
    customer_id: suratJalan.customer_id || '',
    nama_customer: suratJalan.nama_customer || salesOrder.nama_customer_input || '',
    alamat_kirim: suratJalan.alamat_kirim || salesOrder.alamat_kirim || '',
    item: suratJalan.item || orderDisplay.item_summary || salesOrder.item || '',
    qty: suratJalan.qty || orderDisplay.qty_summary || salesOrder.qty || '',
    items: (orderDisplay.details || []).map(function(detail) {
      return {
        nama_item: detail.nama_item || '',
        qty: Number(detail.qty_terkirim || detail.qty || 0),
        satuan: detail.satuan || '',
        harga: Number(detail.harga_final || detail.harga || 0),
        diskon: Number(detail.diskon_final || detail.diskon || 0),
        subtotal: Number(detail.subtotal_final || detail.subtotal || 0)
      };
    }),
    driver: suratJalan.driver || '',
    armada: suratJalan.armada || '',
    status_kirim: suratJalan.status_kirim || '',
    catatan_kirim: suratJalan.catatan_kirim || '',
    sales_nama: salesOrder.sales_nama || '',
    pic_customer: customerContact.pic_customer,
    no_hp_customer: customerContact.no_hp_customer,
    term_pembayaran: salesOrder.term_pembayaran || '',
    subtotal: orderDisplay.subtotal_final || salesOrder.subtotal_final || salesOrder.subtotal || '',
    diskon: orderDisplay.diskon_final || salesOrder.diskon_final || salesOrder.diskon || '',
    total: orderDisplay.total_final || salesOrder.total_final || salesOrder.total || '',
    catatan_order: salesOrder.catatan || ''
  };
}

function getSuratJalanPreviewData(noSo) {
  var noSoKey = String(noSo || '').trim();
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var orderDisplay;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSoKey);
  }

  assertSuratJalanApprovalGateOpen_(salesOrder);

  if (normalizeText_(salesOrder.status_order) !== 'siap kirim') {
    throw new Error('Preview SJ hanya tersedia untuk order berstatus Siap Kirim.');
  }

  orderDisplay = buildSalesOrderClientRow_(salesOrder);
  var customerContact = resolveSuratJalanCustomerContact_(salesOrder);

  return {
    no_surat_jalan: '',
    no_so: salesOrder.no_so || '',
    tanggal_cetak: '',
    tanggal_kirim: normalizeSheetDateToYmd_(salesOrder.tanggal_kirim_rencana || ''),
    customer_id: salesOrder.customer_id || '',
    nama_customer: salesOrder.nama_customer_input || '',
    alamat_kirim: salesOrder.alamat_kirim || '',
    item: orderDisplay.item_summary || salesOrder.item || '',
    qty: orderDisplay.qty_summary || salesOrder.qty || '',
    items: (orderDisplay.details || []).map(function(detail) {
      return {
        nama_item: detail.nama_item || '',
        qty: Number(detail.qty || 0),
        satuan: detail.satuan || '',
        harga: Number(detail.harga || 0),
        diskon: Number(detail.diskon || 0),
        subtotal: Number(detail.subtotal || 0)
      };
    }),
    driver: '',
    armada: '',
    status_kirim: 'Preview',
    catatan_kirim: '',
    sales_nama: salesOrder.sales_nama || '',
    pic_customer: customerContact.pic_customer,
    no_hp_customer: customerContact.no_hp_customer,
    term_pembayaran: salesOrder.term_pembayaran || '',
    subtotal: orderDisplay.subtotal_order || salesOrder.subtotal || '',
    diskon: orderDisplay.diskon_order || salesOrder.diskon || '',
    total: orderDisplay.total_order || salesOrder.total || '',
    catatan_order: salesOrder.catatan || ''
  };
}

function resolveSuratJalanCustomerContact_(salesOrder) {
  var order = salesOrder || {};
  var customer = {};
  var pic = String(order.pic_customer || '').trim();
  var noHp = String(order.no_hp_customer || '').trim();

  if ((!pic || !noHp) && order.customer_id && typeof findCustomerByCode_ === 'function') {
    customer = findCustomerByCode_(order.customer_id) || {};
  }

  return {
    pic_customer: pic || String(customer.pic || '').trim(),
    no_hp_customer: noHp || String(customer.no_hp || '').trim()
  };
}

function assertSuratJalanApprovalGateOpen_(salesOrder) {
  var order = salesOrder || {};
  var statusOrder = normalizeText_(order.status_order);
  var butuhPersetujuan = normalizeText_(order.butuh_persetujuan) === 'ya';
  var hasCustomerDebtRisk = hasCustomerDebtRiskForSuratJalan_(order);
  var approval = null;
  var approvalStatus = '';
  var isApproved = false;

  if (!hasCustomerDebtRisk) {
    return;
  }

  if (typeof findApprovalByNoSo_ === 'function') {
    approval = findApprovalByNoSo_(order.no_so);
  }

  approvalStatus = normalizeText_(approval && approval.status_approval);
  isApproved = approvalStatus === 'disetujui' || (
    statusOrder === 'siap kirim' &&
    !butuhPersetujuan &&
    approvalStatus !== 'ditolak'
  );

  if (!isApproved || statusOrder === 'ditolak') {
    throw new Error(buildSuratJalanApprovalGateMessage_(order, approval));
  }
}

function hasCustomerDebtRiskForSuratJalan_(salesOrder) {
  var order = salesOrder || {};
  var statusPembayaran = normalizeText_(order.status_pembayaran_customer);
  var alasanHold = normalizeText_(order.alasan_hold);

  return Number(order.total_tunggakan || 0) > 0 ||
    statusPembayaran === 'menunggak' ||
    statusPembayaran === 'ditahan' ||
    alasanHold.indexOf('customer menunggak') !== -1 ||
    alasanHold.indexOf('customer ditahan') !== -1 ||
    alasanHold.indexOf('total tunggakan') !== -1;
}

function buildSuratJalanApprovalGateMessage_(salesOrder, approval) {
  var order = salesOrder || {};
  var totalTunggakan = Number(order.total_tunggakan || 0);
  var approvalStatus = normalizeText_(approval && approval.status_approval);
  var parts = [
    'Customer masih memiliki tunggakan.',
    'Nominal tunggakan: Rp ' + formatNumberServer_(totalTunggakan) + '.'
  ];

  if (approvalStatus === 'ditolak' || normalizeText_(order.status_order) === 'ditolak') {
    parts.push('Approval approver ditolak. Surat Jalan tidak boleh dibuat atau dicetak.');
  } else {
    parts.push('Order harus disetujui approver terlebih dahulu sebelum Surat Jalan dibuat atau dicetak.');
  }

  return parts.join(' ');
}

function markOrderDelivered(noSo, userId, catatanKirim) {
  return updateDeliveryOrderStatus_(noSo, userId, 'Terkirim', 'Terkirim', catatanKirim);
}

function cancelDeliveryOrder_(noSo, userId, alasanBatalKirim) {
  var noSoKey = String(noSo || '').trim();
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var statusKirim;
  var now = getNowParts_();
  var reason = String(alasanBatalKirim || '').trim();
  var note;

  if (!suratJalan) {
    throw new Error('Surat jalan aktif tidak ditemukan untuk no_so: ' + noSoKey);
  }

  statusKirim = normalizeText_(suratJalan.status_kirim);
  if (statusKirim !== 'siap kirim' && statusKirim !== 'terkirim') {
    throw new Error('Batal kirim hanya bisa untuk surat jalan berstatus Siap Kirim atau Terkirim.');
  }

  if (salesOrder && String(salesOrder.status_verifikasi_cs || '').trim() === 'Sudah Dicek') {
    throw new Error('Batal kirim tidak bisa dilakukan setelah verifikasi CS tersimpan.');
  }

  if (!reason) {
    throw new Error('Alasan batal kirim wajib diisi.');
  }

  note = 'Batal kirim: ' + reason + ' | Barang kembali ke kantor | ' + now.tanggal + ' ' + now.jam;

  cancelActiveSuratJalanRowsByNoSo_(noSoKey, {
    status_kirim: 'Batal Kirim',
    catatan_kirim: note,
    dibatalkan_oleh: userId,
    tanggal_batal_kirim: now.tanggal + ' ' + now.jam,
    alasan_batal_kirim: reason
  });

  if (salesOrder) {
    updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSoKey, {
      status_order: 'Siap Kirim'
    });

    logStatusOrder_(noSoKey, salesOrder.status_order, 'Siap Kirim', userId, note);
  }

  return {
    success: true,
    no_so: noSoKey,
    no_surat_jalan: suratJalan.no_surat_jalan || '',
    status_kirim: 'Batal Kirim',
    status_order: salesOrder ? 'Siap Kirim' : '',
    sales_order_found: !!salesOrder,
    catatan_kirim: note
  };
}

function cancelActiveSuratJalanRowsByNoSo_(noSo, updates) {
  var sheet = getSheetByName_(APP_CONFIG.SHEETS.SURAT_JALAN);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers;
  var values;
  var noSoIndex;
  var statusIndex;
  var updatedAny = false;
  var updatesObj = updates || {};
  var noSoKey = String(noSo || '').trim();

  if (!lastRow || lastRow < 2 || !lastCol) {
    return 0;
  }

  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  noSoIndex = headers.indexOf('no_so');
  statusIndex = headers.indexOf('status_kirim');

  if (noSoIndex === -1 || statusIndex === -1) {
    throw new Error('Kolom no_so/status_kirim tidak ditemukan di sheet Surat Jalan.');
  }

  values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach(function(row) {
    var isSameNoSo = String(row[noSoIndex] || '').trim() === noSoKey;
    var isActive = normalizeText_(row[statusIndex]) !== 'batal kirim';

    if (!isSameNoSo || !isActive) {
      return;
    }

    headers.forEach(function(header, columnIndex) {
      if (Object.prototype.hasOwnProperty.call(updatesObj, header)) {
        row[columnIndex] = updatesObj[header];
      }
    });
    updatedAny = true;
  });

  if (updatedAny) {
    sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  }

  return values.filter(function(row) {
    return String(row[noSoIndex] || '').trim() === noSoKey &&
      normalizeText_(row[statusIndex]) === 'batal kirim';
  }).length;
}

function cancelDuplicateSuratJalanRowsByNoSo_(noSo, keepNoSuratJalan, userId, reason) {
  var sheet = getSheetByName_(APP_CONFIG.SHEETS.SURAT_JALAN);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers;
  var values;
  var noSoIndex;
  var noSjIndex;
  var statusIndex;
  var catatanIndex;
  var dibatalkanIndex;
  var tanggalBatalIndex;
  var alasanIndex;
  var noSoKey = String(noSo || '').trim();
  var keepNoSjKey = String(keepNoSuratJalan || '').trim();
  var cancelReason = String(reason || '').trim();
  var now = getNowParts_();
  var updated = [];

  if (!noSoKey) {
    throw new Error('Nomor SO wajib diisi.');
  }

  if (!keepNoSjKey) {
    throw new Error('Nomor SJ yang dipertahankan wajib diisi.');
  }

  if (!cancelReason) {
    throw new Error('Alasan koreksi duplicate SJ wajib diisi.');
  }

  if (!lastRow || lastRow < 2 || !lastCol) {
    throw new Error('Sheet Surat Jalan belum memiliki data.');
  }

  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  noSoIndex = headers.indexOf('no_so');
  noSjIndex = headers.indexOf('no_surat_jalan');
  statusIndex = headers.indexOf('status_kirim');
  catatanIndex = headers.indexOf('catatan_kirim');
  dibatalkanIndex = headers.indexOf('dibatalkan_oleh');
  tanggalBatalIndex = headers.indexOf('tanggal_batal_kirim');
  alasanIndex = headers.indexOf('alasan_batal_kirim');

  if (noSoIndex === -1 || noSjIndex === -1 || statusIndex === -1) {
    throw new Error('Kolom no_so/no_surat_jalan/status_kirim tidak ditemukan di sheet Surat Jalan.');
  }

  values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach(function(row) {
    var rowNoSo = String(row[noSoIndex] || '').trim();
    var rowNoSj = String(row[noSjIndex] || '').trim();
    var isActive = normalizeText_(row[statusIndex]) !== 'batal kirim';
    var note;

    if (rowNoSo !== noSoKey || rowNoSj === keepNoSjKey || !isActive) {
      return;
    }

    note = 'Koreksi duplicate SJ: ' + cancelReason + ' | SJ valid: ' + keepNoSjKey + ' | ' + now.tanggal + ' ' + now.jam;
    row[statusIndex] = 'Batal Kirim';
    if (catatanIndex !== -1) row[catatanIndex] = note;
    if (dibatalkanIndex !== -1) row[dibatalkanIndex] = userId || '';
    if (tanggalBatalIndex !== -1) row[tanggalBatalIndex] = now.tanggal + ' ' + now.jam;
    if (alasanIndex !== -1) row[alasanIndex] = cancelReason;
    updated.push(rowNoSj);
  });

  if (!updated.length) {
    return {
      success: true,
      no_so: noSoKey,
      keep_no_surat_jalan: keepNoSjKey,
      canceled_count: 0,
      canceled_no_surat_jalan: [],
      message: 'Tidak ada SJ duplicate aktif yang perlu dibatalkan.'
    };
  }

  sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  logStatusOrder_(noSoKey, '', 'Koreksi Duplicate SJ', userId || '', 'SJ duplicate dibatalkan: ' + updated.join(', '));

  return {
    success: true,
    no_so: noSoKey,
    keep_no_surat_jalan: keepNoSjKey,
    canceled_count: updated.length,
    canceled_no_surat_jalan: updated,
    message: 'SJ duplicate berhasil dibatalkan: ' + updated.join(', ')
  };
}

function cancelDuplicateSuratJalanKeepFirstByNoSo_(noSo, userId, reason) {
  var sheet = getSheetByName_(APP_CONFIG.SHEETS.SURAT_JALAN);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers;
  var values;
  var noSoIndex;
  var noSjIndex;
  var statusIndex;
  var noSoKey = String(noSo || '').trim();
  var activeRows = [];

  if (!noSoKey) {
    throw new Error('Nomor SO wajib diisi.');
  }

  if (!lastRow || lastRow < 2 || !lastCol) {
    throw new Error('Sheet Surat Jalan belum memiliki data.');
  }

  headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  noSoIndex = headers.indexOf('no_so');
  noSjIndex = headers.indexOf('no_surat_jalan');
  statusIndex = headers.indexOf('status_kirim');

  if (noSoIndex === -1 || noSjIndex === -1 || statusIndex === -1) {
    throw new Error('Kolom no_so/no_surat_jalan/status_kirim tidak ditemukan di sheet Surat Jalan.');
  }

  values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach(function(row) {
    if (String(row[noSoIndex] || '').trim() === noSoKey &&
        normalizeText_(row[statusIndex]) !== 'batal kirim') {
      activeRows.push({
        no_surat_jalan: String(row[noSjIndex] || '').trim(),
        status_kirim: row[statusIndex] || ''
      });
    }
  });

  if (activeRows.length <= 1) {
    return {
      success: true,
      no_so: noSoKey,
      active_count: activeRows.length,
      canceled_count: 0,
      canceled_no_surat_jalan: [],
      message: 'Tidak ada duplicate SJ aktif untuk SO ini.'
    };
  }

  return cancelDuplicateSuratJalanRowsByNoSo_(
    noSoKey,
    activeRows[0].no_surat_jalan,
    userId,
    reason || 'Koreksi duplicate SJ, mempertahankan SJ aktif pertama'
  );
}

function repairDuplicateSjSo20260426190035Cs05() {
  return cancelDuplicateSuratJalanKeepFirstByNoSo_(
    'SO-20260426190035',
    'CS05',
    'Koreksi laporan team: input double, nomor SO sama, nomor SJ berbeda'
  );
}

function verifyDeliveredOrder(noSo, userId, payload) {
  var suratJalan = findSuratJalanByNoSo_(noSo);
  var salesOrder = findSalesOrderByNoSo_(noSo);
  var verificationPayload = payload || {};
  var submittedItems = Array.isArray(verificationPayload.items) ? verificationPayload.items : [];
  var orderDetails;
  var detailMap = {};
  var totals;
  var now;
  var nominalTransfer;
  var selisihPembayaran;
  var catatanPembayaran;
  var statusPersetujuanPembayaran;
  var paymentApprovalReason;
  var pendingApproval;

  if (!suratJalan) {
    throw new Error('Surat jalan tidak ditemukan untuk no_so: ' + noSo);
  }

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSo);
  }

  if (normalizeText_(suratJalan.status_kirim) !== 'terkirim') {
    throw new Error('Verifikasi CS hanya bisa dilakukan untuk surat jalan berstatus Terkirim.');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER, APP_CONFIG.HEADERS.SALES_ORDER);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, APP_CONFIG.HEADERS.SALES_ORDER_DETAIL);

  orderDetails = getSalesOrderDetailsByNoSo_(noSo);

  if (!orderDetails.length) {
    throw new Error('Detail order tidak ditemukan. Verifikasi CS membutuhkan detail item.');
  }

  orderDetails.forEach(function(detail) {
    detailMap[String(detail.detail_id || '').trim()] = detail;
  });

  totals = submittedItems.reduce(function(result, item, index) {
    var detailId = String(item.detail_id || '').trim();
    var sourceDetail = detailMap[detailId];
    var qtyTerkirim;
    var hargaFinal;
    var diskonFinal;
    var subtotalFinal;

    if (!detailId || !sourceDetail) {
      throw new Error('Baris verifikasi ke-' + (index + 1) + ' tidak cocok dengan detail order.');
    }

    qtyTerkirim = Number(item.qty_terkirim || 0);
    hargaFinal = Number(item.harga_final || 0);
    diskonFinal = Number(item.diskon_final || 0);
    subtotalFinal = Number(item.subtotal_final);

    if (qtyTerkirim < 0) {
      throw new Error('Qty terkirim untuk item ' + sourceDetail.nama_item + ' tidak boleh negatif.');
    }

    if (hargaFinal < 0) {
      throw new Error('Harga final untuk item ' + sourceDetail.nama_item + ' tidak boleh negatif.');
    }

    if (diskonFinal < 0) {
      throw new Error('Diskon final untuk item ' + sourceDetail.nama_item + ' tidak boleh negatif.');
    }

    if (isNaN(subtotalFinal)) {
      subtotalFinal = (qtyTerkirim * hargaFinal) - diskonFinal;
    }

    if (subtotalFinal < 0) {
      throw new Error('Subtotal final untuk item ' + sourceDetail.nama_item + ' tidak valid.');
    }

    updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER_DETAIL, 'detail_id', detailId, {
      qty_terkirim: qtyTerkirim,
      harga_final: hargaFinal,
      diskon_final: diskonFinal,
      subtotal_final: subtotalFinal
    });

    result.subtotal += qtyTerkirim * hargaFinal;
    result.diskon += diskonFinal;
    result.total += subtotalFinal;
    return result;
  }, {
    subtotal: 0,
    diskon: 0,
    total: 0
  });

  now = getNowParts_();
  nominalTransfer = Number(verificationPayload.nominal_transfer_diterima);
  if (isNaN(nominalTransfer)) {
    nominalTransfer = totals.total;
  }
  if (nominalTransfer < 0) {
    throw new Error('Nominal transfer diterima tidak boleh negatif.');
  }

  selisihPembayaran = nominalTransfer - totals.total;
  catatanPembayaran = String(verificationPayload.catatan_pembayaran_cs || verificationPayload.catatan_verifikasi_cs || '').trim();

  if (selisihPembayaran !== 0 && !catatanPembayaran) {
    throw new Error('Catatan pembayaran wajib diisi jika nominal transfer berbeda dari total final.');
  }

  statusPersetujuanPembayaran = determinePaymentApprovalStatus_(selisihPembayaran);

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.APPROVAL_ORDER, APP_CONFIG.HEADERS.APPROVAL_ORDER);
  pendingApproval = findPendingApprovalByNoSo_(noSo);

  updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, {
    subtotal_final: totals.subtotal,
    diskon_final: totals.diskon,
    total_final: totals.total,
    status_verifikasi_cs: 'Sudah Dicek',
    diverifikasi_oleh: userId,
    tanggal_verifikasi_cs: now.tanggal + ' ' + now.jam,
    catatan_verifikasi_cs: verificationPayload.catatan_verifikasi_cs || '',
    nominal_transfer_diterima: nominalTransfer,
    selisih_pembayaran: selisihPembayaran,
    status_persetujuan_pembayaran: statusPersetujuanPembayaran,
    catatan_pembayaran_cs: catatanPembayaran,
    status_order: statusPersetujuanPembayaran === 'Menunggu Persetujuan' ? 'Menunggu Persetujuan' : resolvePostVerificationOrderStatus_(salesOrder, pendingApproval),
    butuh_persetujuan: statusPersetujuanPembayaran === 'Menunggu Persetujuan' ? 'Ya' : resolvePostVerificationApprovalFlag_(salesOrder, pendingApproval),
    alasan_hold: statusPersetujuanPembayaran === 'Menunggu Persetujuan' ? 'Selisih pembayaran lebih dari Rp 2.000' : resolvePostVerificationHoldReason_(salesOrder, pendingApproval)
  });

  if (statusPersetujuanPembayaran === 'Menunggu Persetujuan') {
    if (!pendingApproval) {
      paymentApprovalReason = buildPaymentApprovalReason_(totals.total, nominalTransfer, selisihPembayaran, catatanPembayaran);
      createApprovalOrder_(noSo, userId, paymentApprovalReason);
      logStatusOrder_(noSo, salesOrder.status_order, 'Menunggu Persetujuan', userId, paymentApprovalReason);
    }
  } else if (pendingApproval && isPaymentApprovalReason_(pendingApproval.alasan_approval)) {
    updateRowByKey_(APP_CONFIG.SHEETS.APPROVAL_ORDER, 'approval_id', pendingApproval.approval_id, {
      status_approval: 'Ditolak',
      diputuskan_oleh: userId,
      tanggal_keputusan: now.tanggal + ' ' + now.jam,
      catatan_approval: 'Ditutup otomatis karena CS mengubah nominal transfer sehingga tidak perlu approval selisih pembayaran.'
    });
  }

  createOrSyncTagihanFromSalesOrderTempo_(Object.keys(salesOrder || {}).reduce(function(result, key) {
    result[key] = salesOrder[key];
    return result;
  }, {}), suratJalan, totals, { user_id: userId });

  if (normalizeText_(salesOrder.channel_sales) === 'slf') {
    updateSalesOrderCommissionStatus_(noSo, 'Menunggu Pembayaran', {
      catatan: 'Order sudah diverifikasi CS, menunggu pembayaran diterima perusahaan'
    });
  }

  return {
    success: true,
    no_so: noSo,
    status_verifikasi_cs: 'Sudah Dicek',
    subtotal_final: totals.subtotal,
    diskon_final: totals.diskon,
    total_final: totals.total,
    nominal_transfer_diterima: nominalTransfer,
    selisih_pembayaran: selisihPembayaran,
    status_persetujuan_pembayaran: statusPersetujuanPembayaran,
    requires_payment_approval: statusPersetujuanPembayaran === 'Menunggu Persetujuan'
  };
}

function completeOrder(noSo, userId, catatanKirim) {
  var salesOrder = findSalesOrderByNoSo_(noSo);
  var result;
  var now = getNowParts_();
  var isSlfCashOrder;

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSo);
  }

  if (String(salesOrder.status_verifikasi_cs || '').trim() !== 'Sudah Dicek') {
    throw new Error('Order belum bisa selesai. CS wajib simpan verifikasi qty dan nominal final terlebih dahulu.');
  }

  if (String(salesOrder.status_persetujuan_pembayaran || '').trim() === 'Menunggu Persetujuan') {
    throw new Error('Transaksi belum bisa diselesaikan. Silahkan meminta approval approver untuk menyelesaikan transaksi.');
  }

  if (String(salesOrder.status_persetujuan_pembayaran || '').trim() === 'Ditolak') {
    throw new Error('Transaksi belum bisa diselesaikan karena approval selisih pembayaran ditolak.');
  }

  if (String(salesOrder.status_persetujuan_pembayaran || '').trim() === 'Kurang Bayar') {
    throw new Error('Transaksi belum bisa diselesaikan karena nominal transfer lebih kecil dari total final.');
  }

  result = updateDeliveryOrderStatus_(noSo, userId, 'Selesai', 'Selesai', catatanKirim);

  updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, {
    status_export_kledo: 'Siap Export',
    tanggal_selesai: now.tanggal
  });

  isSlfCashOrder = normalizeText_(salesOrder.channel_sales) === 'slf' &&
    normalizeText_(salesOrder.term_pembayaran) === 'cash';

  if (isSlfCashOrder) {
    updateSalesOrderCommissionStatus_(noSo, 'Siap Cair', {
      catatan: 'Order cash selesai dan sudah diverifikasi CS, komisi siap cair'
    });
    result.status_komisi = 'Siap Cair';
    result.tanggal_siap_cair = now.tanggal;
  }

  result.status_export_kledo = 'Siap Export';
  result.tanggal_selesai = now.tanggal;
  return result;
}

function determinePaymentApprovalStatus_(selisihPembayaran) {
  var diff = Number(selisihPembayaran || 0);

  if (diff < 0) {
    return 'Kurang Bayar';
  }

  if (diff > 2000) {
    return 'Menunggu Persetujuan';
  }

  if (diff > 0) {
    return 'Lebih Bayar Disetujui Otomatis';
  }

  return 'Sesuai';
}

function buildPaymentApprovalReason_(totalFinal, nominalTransfer, selisihPembayaran, catatanPembayaran) {
  return [
    'Approval selisih pembayaran',
    'Total final Rp ' + formatNumberServer_(totalFinal),
    'Transfer diterima Rp ' + formatNumberServer_(nominalTransfer),
    'Lebih bayar Rp ' + formatNumberServer_(selisihPembayaran),
    'Catatan CS: ' + String(catatanPembayaran || '-').trim()
  ].join(' | ');
}

function isPaymentApprovalReason_(reason) {
  return normalizeText_(reason || '').indexOf('approval selisih pembayaran') !== -1;
}

function resolvePostVerificationOrderStatus_(salesOrder, pendingApproval) {
  if (pendingApproval && isPaymentApprovalReason_(pendingApproval.alasan_approval) &&
    normalizeText_(salesOrder.status_order) === 'menunggu persetujuan') {
    return 'Terkirim';
  }

  return salesOrder.status_order;
}

function resolvePostVerificationApprovalFlag_(salesOrder, pendingApproval) {
  if (pendingApproval && isPaymentApprovalReason_(pendingApproval.alasan_approval)) {
    return 'Tidak';
  }

  return salesOrder.butuh_persetujuan;
}

function resolvePostVerificationHoldReason_(salesOrder, pendingApproval) {
  if (pendingApproval && isPaymentApprovalReason_(pendingApproval.alasan_approval)) {
    return '';
  }

  return salesOrder.alasan_hold;
}

function generateKledoExportBatchFile(currentUser, options) {
  var readyOrders = getReadyKledoExportOrders_(options || {});
  var exportRows = [];
  var workbookXml;
  var fileName;
  var itemCount = 0;

  if (!readyOrders.length) {
    throw new Error('Belum ada order Siap Export untuk dibuatkan batch Kledo.');
  }

  readyOrders.forEach(function(order) {
    var suratJalan = findSuratJalanByNoSo_(order.no_so) || {};
    var rows = buildKledoExportRows_(order, suratJalan);

    itemCount += rows.length;
    exportRows = exportRows.concat(rows);
  });

  workbookXml = buildKledoExportWorkbookXml_([APP_CONFIG.KLEDO_EXPORT.HEADERS].concat(exportRows));
  fileName = [
    APP_CONFIG.KLEDO_EXPORT.FILE_PREFIX + '-batch',
    Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss')
  ].join('-') + '.xls';

  return {
    success: true,
    order_count: readyOrders.length,
    item_count: itemCount,
    no_so_list: readyOrders.map(function(order) {
      return order.no_so;
    }),
    file_name: fileName,
    mime_type: 'application/vnd.ms-excel;charset=utf-8;',
    file_content: workbookXml
  };
}

function markKledoBatchExported(currentUser, catatanExport, options) {
  var readyOrders = getReadyKledoExportOrders_(options || {});
  var now = getNowParts_();
  var noSoList;

  if (!readyOrders.length) {
    throw new Error('Belum ada order Siap Export untuk ditandai Sudah Export.');
  }

  noSoList = readyOrders.map(function(order) {
    return order.no_so;
  });

  noSoList.forEach(function(noSo) {
    updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, {
      status_export_kledo: 'Sudah Export',
      tanggal_export_kledo: now.tanggal + ' ' + now.jam,
      diekspor_oleh: currentUser.user_id || '',
      catatan_export_kledo: String(catatanExport || '').trim()
    });
  });

  return {
    success: true,
    status_export_kledo: 'Sudah Export',
    order_count: noSoList.length,
    no_so_list: noSoList
  };
}

function buildKledoExportRows_(salesOrder, suratJalan) {
  var details = Array.isArray(salesOrder.details) ? salesOrder.details : [];
  var customerName = salesOrder.customer || salesOrder.nama_customer_input || '';
  var orderNote = salesOrder.catatan || salesOrder.catatan_order || '';
  var tanggalKirimEfektif = resolveEffectiveSuratJalanTanggalKirim_(suratJalan || {}, salesOrder || {});

  if (!details.length) {
    throw new Error('Detail item order tidak ditemukan untuk export Kledo.');
  }

  return details.map(function(detail) {
    var qty = Number(detail.qty_terkirim || detail.qty || 0);
    var harga = Number(detail.harga_final || detail.harga || 0);
    var diskon = Number(detail.diskon_final || detail.diskon || 0);

    return [
      customerName,
      '',
      salesOrder.alamat_kirim || '',
      salesOrder.no_hp_customer || '',
      '',
      salesOrder.no_so || '',
      '',
      formatKledoDate_(salesOrder.tanggal_order || ''),
      formatKledoDate_(salesOrder.tanggal_jatuh_tempo || ''),
      APP_CONFIG.KLEDO_EXPORT.WAREHOUSE_NAME,
      buildKledoOrderNote_({
        catatan: orderNote,
        catatan_verifikasi_cs: salesOrder.catatan_verifikasi_cs || ''
      }),
      formatKledoDate_(tanggalKirimEfektif || salesOrder.tanggal_kirim_rencana || ''),
      suratJalan.armada || '',
      '',
      APP_CONFIG.KLEDO_EXPORT.INCLUDE_TAX,
      detail.nama_item || '',
      detail.kode_item || '',
      '',
      qty > 0 ? String(qty) : '0',
      resolveKledoUnitName_(detail),
      String(diskon || 0),
      '',
      String(harga || 0),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Webapp ATS'
    ];
  });
}

function resolveKledoUnitName_(detail) {
  var configuredUnit = String(APP_CONFIG.KLEDO_EXPORT.UNIT_NAME || '').trim();

  if (configuredUnit) {
    return configuredUnit;
  }

  return '';
}

function buildKledoOrderNote_(salesOrder) {
  var notes = [];

  if (salesOrder.catatan) {
    notes.push('Order: ' + salesOrder.catatan);
  }

  if (salesOrder.catatan_verifikasi_cs) {
    notes.push('Verifikasi CS: ' + salesOrder.catatan_verifikasi_cs);
  }

  return notes.join(' | ');
}

function formatKledoDate_(value) {
  var normalized = normalizeSheetDateToYmd_(value);
  var match = String(normalized || '').match(/^(\d{4})-(\d{2})-(\d{2})/);

  if (!match) {
    return '';
  }

  return [match[3], match[2], match[1]].join('/');
}

function buildKledoExportWorkbookXml_(rows) {
  var safeRows = Array.isArray(rows) ? rows : [];
  var rowXml = safeRows.map(function(row) {
    return '<Row>' + (Array.isArray(row) ? row : []).map(function(cell) {
      return '<Cell><Data ss:Type="String">' + escapeXml_(cell) + '</Data></Cell>';
    }).join('') + '</Row>';
  }).join('');

  return [
    '<?xml version="1.0"?>',
    '<?mso-application progid="Excel.Sheet"?>',
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:o="urn:schemas-microsoft-com:office:office"',
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"',
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:html="http://www.w3.org/TR/REC-html40">',
    '<Worksheet ss:Name="Kledo Import">',
    '<Table>',
    rowXml,
    '</Table>',
    '</Worksheet>',
    '</Workbook>'
  ].join('');
}

function escapeXml_(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function testCreateSuratJalanFromLatestReadyOrder() {
  var salesOrder = getLatestSalesOrderByStatus_('Siap Kirim');

  if (!salesOrder) {
    throw new Error('Tidak ada order Siap Kirim untuk dites');
  }

  console.log(JSON.stringify(createSuratJalan(salesOrder.no_so, {
    driver: 'Dedi',
    armada: 'B 9123 KXA',
    catatan_kirim: 'Surat jalan test'
  })));
}

function testCreateSuratJalanByNoSo() {
  console.log(JSON.stringify(createSuratJalan('ISI_NO_SO_DI_SINI', {
    driver: 'Dedi',
    armada: 'B 9123 KXA',
    catatan_kirim: 'Surat jalan test manual'
  })));
}

function testCreateSuratJalanManual() {
  console.log(JSON.stringify(createSuratJalan('SO-20260406232617', {
    driver: '',
    armada: '',
    catatan_kirim: 'Surat jalan dari order webapp'
  })));
}

function testMarkLatestDelivered() {
  var suratJalan = getLatestSuratJalanByStatus_('Siap Kirim');

  if (!suratJalan) {
    throw new Error('Tidak ada surat jalan dengan status Siap Kirim untuk dites');
  }

  console.log(JSON.stringify(markOrderDelivered(suratJalan.no_so, 'U002', 'Barang sudah dikirim')));
}

function testCompleteLatestOrder() {
  var suratJalan = getLatestSuratJalanByStatus_('Terkirim');

  if (!suratJalan) {
    throw new Error('Tidak ada surat jalan dengan status Terkirim untuk dites');
  }

  console.log(JSON.stringify(completeOrder(suratJalan.no_so, 'U002', 'Order selesai')));
}

function findSuratJalanByNoSo_(noSo) {
  var rows = getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN).filter(function(row) {
    return String(row.no_so).trim() === String(noSo).trim() &&
      normalizeText_(row.status_kirim) !== 'batal kirim';
  });

  return rows.length ? rows[rows.length - 1] : null;
}

function resolveEffectiveSuratJalanTanggalKirim_(suratJalan, salesOrder) {
  var suratJalanRow = suratJalan || {};
  var salesOrderRow = salesOrder || {};
  var statusKirim = normalizeText_(suratJalanRow.status_kirim);
  var tanggalSo = normalizeSheetDateToYmd_(salesOrderRow.tanggal_kirim_rencana);
  var tanggalSj = normalizeSheetDateToYmd_(suratJalanRow.tanggal_kirim || suratJalanRow.tanggal_cetak || '');

  if (statusKirim !== 'terkirim' && statusKirim !== 'selesai' && tanggalSo) {
    return tanggalSo;
  }

  return tanggalSj || tanggalSo || '';
}

function syncSuratJalanDraftFromSalesOrder_(noSo) {
  var noSoKey = String(noSo || '').trim();
  var suratJalan = findSuratJalanByNoSo_(noSoKey);
  var salesOrder = findSalesOrderByNoSo_(noSoKey);
  var orderDisplay;
  var statusKirim;

  if (!suratJalan || !salesOrder) {
    return null;
  }

  statusKirim = normalizeText_(suratJalan.status_kirim);
  if (statusKirim === 'terkirim' || statusKirim === 'selesai') {
    return suratJalan;
  }

  orderDisplay = buildSalesOrderClientRow_(salesOrder);

  return updateRowByKey_(APP_CONFIG.SHEETS.SURAT_JALAN, 'no_surat_jalan', suratJalan.no_surat_jalan, {
    tanggal_kirim: normalizeSheetDateToYmd_(salesOrder.tanggal_kirim_rencana || suratJalan.tanggal_kirim || ''),
    customer_id: salesOrder.customer_id || suratJalan.customer_id || '',
    nama_customer: salesOrder.nama_customer_input || suratJalan.nama_customer || '',
    alamat_kirim: salesOrder.alamat_kirim || suratJalan.alamat_kirim || '',
    item: orderDisplay.item_summary || salesOrder.item || suratJalan.item || '',
    qty: orderDisplay.qty_summary || salesOrder.qty || suratJalan.qty || ''
  });
}

function getLatestSalesOrderByStatus_(statusOrder) {
  var rows = getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(row) {
    return normalizeText_(row.status_order) === normalizeText_(statusOrder);
  });

  if (!rows.length) {
    return null;
  }

  return rows[rows.length - 1];
}

function getLatestSuratJalanByStatus_(statusKirim) {
  var rows = getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN).filter(function(row) {
    return normalizeText_(row.status_kirim) === normalizeText_(statusKirim);
  });

  if (!rows.length) {
    return null;
  }

  return rows[rows.length - 1];
}

function updateDeliveryOrderStatus_(noSo, userId, statusKirimBaru, statusOrderBaru, catatanKirim) {
  var suratJalan = findSuratJalanByNoSo_(noSo);

  if (!suratJalan) {
    throw new Error('Surat jalan tidak ditemukan untuk no_so: ' + noSo);
  }

  var salesOrder = findSalesOrderByNoSo_(noSo);

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSo);
  }

  updateRowByKey_(APP_CONFIG.SHEETS.SURAT_JALAN, 'no_surat_jalan', suratJalan.no_surat_jalan, {
    status_kirim: statusKirimBaru,
    catatan_kirim: catatanKirim || suratJalan.catatan_kirim || ''
  });

  updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, {
    status_order: statusOrderBaru
  });

  logStatusOrder_(noSo, salesOrder.status_order, statusOrderBaru, userId, catatanKirim || statusKirimBaru);

  return {
    success: true,
    no_so: noSo,
    status_kirim: statusKirimBaru,
    status_order: statusOrderBaru
  };
}
