function submitSalesOrder(payload) {
  validateSalesOrderPayload_(payload);

  var now = getNowParts_();
  var noSo = generateDocNumber_(APP_CONFIG.DOC_PREFIX.SALES_ORDER);
  var items = normalizeOrderItems_(payload);
  var customerCheck = buildOrderCustomerCheck_(payload);
  var priceCheck = buildOrderPriceCheck_(items);
  var approvalDecision = mergeOrderApprovalChecks_(customerCheck, priceCheck);
  var totals = calculateOrderTotals_(items);
  var prioritasKirim = resolvePrioritasKirim_(payload.tanggal_kirim_rencana, now.timestamp);
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
    alasan_hold: approvalDecision.alasan_hold
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

function validateSalesOrderPayload_(payload) {
  var requiredFields = [
    'sales_id',
    'sales_nama',
    'jenis_customer',
    'alamat_kirim',
    'pic_customer',
    'no_hp_customer',
    'term_pembayaran',
    'tanggal_jatuh_tempo',
    'tanggal_kirim_rencana'
  ];

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
    status_order: eligibility.butuh_persetujuan === 'Ya' ? 'Menunggu Persetujuan' : 'Siap Kirim',
    butuh_persetujuan: eligibility.butuh_persetujuan,
    alasan_hold: eligibility.alasan_hold
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

function getProductByNameServer_(itemName) {
  return getProductCatalog_().find(function(row) {
    return String(row.nama_item || '').trim() === String(itemName || '').trim();
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
