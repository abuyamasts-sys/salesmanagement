function getCustomers() {
  return getSheetData_(APP_CONFIG.SHEETS.MASTER_CUSTOMER);
}

function getActiveCustomers() {
  return getCustomers().filter(function(customer) {
    return normalizeText_(customer.status_customer) !== 'ditahan';
  });
}

function getCustomerByCode(kodeCustomer) {
  var customer = findCustomerByCode_(kodeCustomer);

  if (!customer) {
    throw new Error('Customer tidak ditemukan: ' + kodeCustomer);
  }

  return buildCustomerStatusResult_(customer);
}

function checkCustomerEligibility(kodeCustomer) {
  var customer = findCustomerByCode_(kodeCustomer);

  if (!customer) {
    return {
      found: false,
      eligible: false,
      butuh_persetujuan: 'Ya',
      alasan_hold: 'Customer tidak ditemukan'
    };
  }

  return buildCustomerStatusResult_(customer);
}

function testGetCustomers() {
  Logger.log(getCustomers());
}

function testCheckCustomerEligibility() {
  Logger.log(checkCustomerEligibility('CUST001'));
  Logger.log(checkCustomerEligibility('CUST002'));
  Logger.log(checkCustomerEligibility('CUST003'));
}

function getCustomersOwnedBySales_(salesId) {
  var normalizedSalesId = normalizeText_(salesId);

  if (!normalizedSalesId) {
    return [];
  }

  return getCustomers().filter(function(customer) {
    var ownerId = normalizeText_(customer.sales_owner_id);
    return !ownerId || ownerId === normalizedSalesId;
  });
}

function isCustomerOwnedBySales_(kodeCustomer, salesId) {
  var customer = findCustomerByCode_(kodeCustomer);
  var normalizedSalesId = normalizeText_(salesId);
  var ownerId;

  if (!customer) {
    return false;
  }

  if (!normalizedSalesId) {
    return false;
  }

  ownerId = normalizeText_(customer.sales_owner_id);
  return !ownerId || ownerId === normalizedSalesId;
}

function assignCustomerOwnership_(kodeCustomer, salesId, salesName, updatedBy, note) {
  var customer = findCustomerByCode_(kodeCustomer);
  var normalizedSalesId = String(salesId || '').trim();
  var normalizedUpdatedBy = String(updatedBy || '').trim();
  var profile;
  var resolvedSalesName;
  var now;
  var ownershipNote;
  var updates;

  if (!customer) {
    throw new Error('Customer tidak ditemukan: ' + kodeCustomer);
  }

  if (!normalizedSalesId) {
    throw new Error('salesId wajib diisi untuk assign ownership customer.');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, APP_CONFIG.HEADERS.MASTER_CUSTOMER);
  profile = getCurrentUserProfile(normalizedSalesId);
  resolvedSalesName = String(salesName || '').trim() || profile.nama_user || '';
  now = getNowParts_();
  ownershipNote = String(note || '').trim() || 'Ownership customer di-assign manual';

  if (normalizedUpdatedBy) {
    ownershipNote += ' oleh ' + normalizedUpdatedBy;
  }

  updates = {
    sales_owner_id: normalizedSalesId,
    sales_owner_nama: resolvedSalesName,
    channel_akuisisi: String(customer.channel_akuisisi || '').trim() || profile.channel_sales_default || '',
    status_kepemilikan: 'Aktif',
    tanggal_update_owner: now.tanggal,
    catatan_kepemilikan: ownershipNote
  };

  return updateRowByKey_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, 'kode_customer', customer.kode_customer, updates);
}

function findCustomerByCode_(kodeCustomer) {
  var targetCode = normalizeText_(kodeCustomer);

  return getCustomers().find(function(customer) {
    return normalizeText_(customer.kode_customer) === targetCode;
  }) || null;
}

function buildCustomerStatusResult_(customer) {
  var statusCustomer = normalizeText_(customer.status_customer);
  var statusPembayaran = normalizeText_(customer.status_pembayaran);
  var isHold = statusCustomer === 'menunggak' ||
    statusCustomer === 'ditahan' ||
    statusPembayaran === 'menunggak' ||
    statusPembayaran === 'ditahan';
  var alasanHold = buildCustomerHoldReason_(customer);

  if (!isHold) {
    alasanHold = '';
  }

  return {
    found: true,
    eligible: !isHold,
    butuh_persetujuan: isHold ? 'Ya' : 'Tidak',
    alasan_hold: alasanHold,
    customer: customer
  };
}

function buildCustomerHoldReason_(customer) {
  var statusCustomer = normalizeText_(customer.status_customer);
  var statusPembayaran = normalizeText_(customer.status_pembayaran);
  var totalTunggakan = Number(customer.total_tunggakan || 0);
  var jumlahNotaOverdue = Number(customer.jumlah_nota_overdue || 0);
  var tanggalJatuhTempo = formatCustomerDate_(customer.tanggal_jatuh_tempo_terdekat);
  var catatanPiutang = String(customer.catatan_piutang || customer.catatan || '').trim();
  var parts = [];

  if (statusCustomer === 'ditahan' || statusPembayaran === 'ditahan') {
    parts.push('Customer ditahan');
  } else if (statusCustomer === 'menunggak' || statusPembayaran === 'menunggak') {
    parts.push('Customer menunggak');
  }

  if (totalTunggakan > 0) {
    parts.push('Total tunggakan: ' + formatNumberServer_(totalTunggakan));
  }

  if (jumlahNotaOverdue > 0) {
    parts.push('Nota overdue: ' + jumlahNotaOverdue);
  }

  if (tanggalJatuhTempo) {
    parts.push('JT terdekat: ' + tanggalJatuhTempo);
  }

  if (catatanPiutang) {
    parts.push('Catatan: ' + catatanPiutang);
  }

  return parts.join('. ') || '';
}

function formatCustomerDate_(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(value, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }

  return String(value).trim();
}

function formatNumberServer_(value) {
  return Number(value || 0).toLocaleString('id-ID');
}

function ensureCustomerMasterForNewOrder_(payload) {
  var existingCustomer = findExistingCustomerForNewOrder_(payload);
  var ownershipMeta;

  if (existingCustomer) {
    return existingCustomer;
  }

  ownershipMeta = buildCustomerOwnershipMetaForNewOrder_(payload);

  var customerRow = {
    kode_customer: generateNextCustomerCode_(),
    nama_customer: String(payload.nama_customer_input || '').trim(),
    tipe_customer: 'Retail',
    kategori_customer: 'Customer Baru',
    alamat: String(payload.alamat_kirim || '').trim(),
    link_google_maps: String(payload.link_google_maps || '').trim(),
    pic: String(payload.pic_customer || '').trim(),
    no_hp: String(payload.no_hp_customer || '').trim(),
    latitude: String(payload.latitude || '').trim(),
    longitude: String(payload.longitude || '').trim(),
    status_customer: 'Baru',
    status_pembayaran: 'Lancar',
    total_tunggakan: 0,
    jumlah_nota_overdue: 0,
    tanggal_jatuh_tempo_terdekat: '',
    catatan_piutang: '',
    limit_tunggakan: 0,
    catatan: 'Auto dibuat dari sales order customer baru',
    dibuat_oleh_user_id: ownershipMeta.dibuat_oleh_user_id,
    dibuat_oleh_nama: ownershipMeta.dibuat_oleh_nama,
    sales_owner_id: ownershipMeta.sales_owner_id,
    sales_owner_nama: ownershipMeta.sales_owner_nama,
    channel_akuisisi: ownershipMeta.channel_akuisisi,
    tanggal_akuisisi: ownershipMeta.tanggal_akuisisi,
    status_kepemilikan: ownershipMeta.status_kepemilikan,
    tanggal_update_owner: ownershipMeta.tanggal_update_owner,
    catatan_kepemilikan: ownershipMeta.catatan_kepemilikan
  };

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, APP_CONFIG.HEADERS.MASTER_CUSTOMER);
  appendRowByHeaders_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, customerRow);
  return customerRow;
}

function buildCustomerOwnershipMetaForNewOrder_(payload) {
  var salesId = String(payload && payload.sales_id || '').trim();
  var salesName = String(payload && payload.sales_nama || '').trim();
  var profile = salesId ? getCurrentUserProfile(salesId) : null;
  var now = getNowParts_();

  return {
    dibuat_oleh_user_id: salesId,
    dibuat_oleh_nama: salesName || profile && profile.nama_user || '',
    sales_owner_id: salesId,
    sales_owner_nama: salesName || profile && profile.nama_user || '',
    channel_akuisisi: String(payload && payload.channel_sales || '').trim() || profile && profile.channel_sales_default || '',
    tanggal_akuisisi: now.tanggal,
    status_kepemilikan: 'Aktif',
    tanggal_update_owner: now.tanggal,
    catatan_kepemilikan: 'Auto assign saat customer baru dibuat dari sales order'
  };
}

function findExistingCustomerForNewOrder_(payload) {
  var targetName = normalizeText_(payload.nama_customer_input);
  var targetPhone = normalizeText_(payload.no_hp_customer);

  if (!targetName) {
    return null;
  }

  return getCustomers().find(function(customer) {
    var customerName = normalizeText_(customer.nama_customer);
    var customerPhone = normalizeText_(customer.no_hp);

    if (customerName !== targetName) {
      return false;
    }

    if (!targetPhone) {
      return true;
    }

    return customerPhone === targetPhone;
  }) || null;
}

function generateNextCustomerCode_() {
  var maxNumber = 0;

  getCustomers().forEach(function(customer) {
    var match = String(customer.kode_customer || '').trim().toUpperCase().match(/^CUST(\d+)$/);
    var numberValue;

    if (!match) {
      return;
    }

    numberValue = Number(match[1] || 0);
    if (numberValue > maxNumber) {
      maxNumber = numberValue;
    }
  });

  return 'CUST' + String(maxNumber + 1).padStart(3, '0');
}
