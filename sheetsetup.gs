function setupAppSheets() {
  try {
    validateSheetSetupConfig_();
    return ensureConfiguredSheetsSafely_();
  } catch (error) {
    Logger.log(error && error.stack ? error.stack : error);
    throw new Error('Setup sheet gagal: ' + (error && error.message ? error.message : error));
  }
}

function setupMissingSheetsNow() {
  try {
    validateSheetSetupConfig_();
    return ensureConfiguredSheetsSafely_([
      'MASTER_KOMISI_SLF',
      'KOMISI_SLF_MUTASI'
    ]);
  } catch (error) {
    Logger.log(error && error.stack ? error.stack : error);
    throw new Error('Setup missing sheets gagal: ' + (error && error.message ? error.message : error));
  }
}

function validateSheetSetupConfig_() {
  if (!APP_CONFIG || !APP_CONFIG.SPREADSHEET_ID) {
    throw new Error('APP_CONFIG.SPREADSHEET_ID belum diisi.');
  }
}

function ensureConfiguredSheetsSafely_(sheetKeys) {
  var keys = Array.isArray(sheetKeys) && sheetKeys.length
    ? sheetKeys
    : Object.keys(APP_CONFIG.HEADERS);
  var results = [];

  keys.forEach(function(sheetKey) {
    var sheetName = APP_CONFIG.SHEETS[sheetKey];
    var headers = APP_CONFIG.HEADERS[sheetKey];
    var existingSheet;
    var existingHeaders;
    var missingHeaders;
    var ensuredSheet;

    if (!sheetName) {
      throw new Error('Nama sheet tidak ditemukan untuk key: ' + sheetKey);
    }

    existingSheet = getSheetByNameOrNull_(sheetName);
    existingHeaders = existingSheet && existingSheet.getLastColumn() > 0
      ? existingSheet.getRange(1, 1, 1, existingSheet.getLastColumn()).getValues()[0].map(function(header) {
        return String(header || '').trim();
      })
      : [];
    missingHeaders = (headers || []).filter(function(header) {
      return existingHeaders.indexOf(header) === -1;
    });

    ensuredSheet = ensureSheetHeadersContain_(sheetName, headers);
    if (ensuredSheet) {
      ensuredSheet.setFrozenRows(1);
    }

    results.push({
      sheet_key: sheetKey,
      sheet_name: sheetName,
      created: !existingSheet,
      appended_headers: missingHeaders
    });
  });

  return {
    success: true,
    message: 'Setup sheet aman selesai.',
    sheets: results
  };
}

function ensureSheetWithHeaders_(sheetName, headers) {
  return ensureSheetHeadersContain_(sheetName, headers);
}

function seedDummyDataAsSales() {
  return seedDummyData_('Sales');
}

function seedDummyDataAsApprover() {
  return seedDummyData_('Approver');
}

function seedDummyDataAsAdmin() {
  return seedDummyData_('CS/Admin');
}

function clearDummyTransactionsOnly() {
  setupAppSheets();

  [
    APP_CONFIG.SHEETS.SALES_ORDER,
    APP_CONFIG.SHEETS.SALES_ORDER_DETAIL,
    APP_CONFIG.SHEETS.APPROVAL_ORDER,
    APP_CONFIG.SHEETS.SURAT_JALAN,
    APP_CONFIG.SHEETS.LOG_STATUS_ORDER,
    APP_CONFIG.SHEETS.LOG_REVISI_ORDER,
    APP_CONFIG.SHEETS.TAGIHAN,
    APP_CONFIG.SHEETS.KPI_TARGET_SALES,
    APP_CONFIG.SHEETS.KPI_LOG
  ].forEach(function(sheetName) {
    clearSheetRowsPreserveHeader_(sheetName);
  });

  return {
    success: true,
    message: 'Transaksi dummy berhasil dihapus. Data master customer dan user tetap disimpan.'
  };
}

function clearAllDummyData() {
  setupAppSheets();
  clearDataRows_();

  return {
    success: true,
    message: 'Semua data dummy berhasil dihapus dari master dan transaksi.'
  };
}

function seedDummyData_(currentRole) {
  setupAppSheets();
  clearDataRows_();

  var currentEmail = getCurrentUserEmail_() || 'ganti-email@login.com';
  var role = currentRole || 'Sales';
  var userSeed = buildDummyUsers_(currentEmail, role);
  var customerSeeds = buildDummyCustomers_();
  var productSeeds = getDefaultProductCatalogSeed_().map(function(item, index) {
    item.harga_dasar = [19000, 8000, 32000, 35000, 38000, 42000, 47000, 52000][index] || 0;
    item.harga_default = item.harga_dasar;
    item.diupdate_oleh = 'U030';
    item.tanggal_update_harga = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    return item;
  });

  writeRowsByHeaders_(APP_CONFIG.SHEETS.MASTER_USER, APP_CONFIG.HEADERS.MASTER_USER, userSeed);
  writeRowsByHeaders_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, APP_CONFIG.HEADERS.MASTER_CUSTOMER, customerSeeds);
  writeRowsByHeaders_(APP_CONFIG.SHEETS.MASTER_ITEM, APP_CONFIG.HEADERS.MASTER_ITEM, productSeeds);

  createDummyTransactions_();

  return {
    success: true,
    current_email: currentEmail,
    active_role: role,
    users: userSeed.length,
    customers: customerSeeds.length,
    sales_orders: getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).length,
    approvals: getSheetData_(APP_CONFIG.SHEETS.APPROVAL_ORDER).length,
    surat_jalan: getSheetData_(APP_CONFIG.SHEETS.SURAT_JALAN).length
  };
}

function clearDataRows_() {
  [
    APP_CONFIG.SHEETS.MASTER_CUSTOMER,
    APP_CONFIG.SHEETS.MASTER_USER,
    APP_CONFIG.SHEETS.MASTER_ITEM,
    APP_CONFIG.SHEETS.SALES_ORDER,
    APP_CONFIG.SHEETS.SALES_ORDER_DETAIL,
    APP_CONFIG.SHEETS.APPROVAL_ORDER,
    APP_CONFIG.SHEETS.SURAT_JALAN,
    APP_CONFIG.SHEETS.LOG_STATUS_ORDER,
    APP_CONFIG.SHEETS.LOG_REVISI_ORDER,
    APP_CONFIG.SHEETS.TAGIHAN,
    APP_CONFIG.SHEETS.KPI_TARGET_SALES,
    APP_CONFIG.SHEETS.KPI_LOG
  ].forEach(function(sheetName) {
    clearSheetRowsPreserveHeader_(sheetName);
  });
}

function clearSheetRowsPreserveHeader_(sheetName) {
  var sheet = getSheetByName_(sheetName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow > 1 && lastColumn > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}

function writeRowsByHeaders_(sheetName, headers, rows) {
  var sheet = getSheetByName_(sheetName);
  var safeRows = rows || [];

  if (!safeRows.length) {
    return;
  }

  var matrix = safeRows.map(function(row) {
    return headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
    });
  });

  sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
}

function buildDummyUsers_(currentEmail, currentRole) {
  var defaultUsers = {
    'Sales': { user_id: 'U001', nama_user: 'Andi Sales', role: 'Sales' },
    'CS/Admin': { user_id: 'U002', nama_user: 'Citra Admin', role: 'CS/Admin' },
    'Approver': { user_id: 'U003', nama_user: 'Bima Approver', role: 'Approver' }
  };
  var activeUser = defaultUsers[currentRole] || defaultUsers.Sales;

  return [
    {
      user_id: activeUser.user_id,
      nama_user: activeUser.nama_user,
      role: activeUser.role,
      no_hp: '081300000000',
      email: currentEmail,
      password: activeUser.user_id,
      status_aktif: 'Aktif'
    },
    {
      user_id: 'U010',
      nama_user: 'Andi Sales Dummy',
      role: 'Sales',
      no_hp: '081300000010',
      email: 'sales.dummy@airtis.local',
      password: 'U010',
      status_aktif: 'Aktif'
    },
    {
      user_id: 'U020',
      nama_user: 'Citra Admin Dummy',
      role: 'CS/Admin',
      no_hp: '081300000020',
      email: 'admin.dummy@airtis.local',
      password: 'U020',
      status_aktif: 'Aktif'
    },
    {
      user_id: 'U030',
      nama_user: 'Bima Approver Dummy',
      role: 'Approver',
      no_hp: '081300000030',
      email: 'approver.dummy@airtis.local',
      password: 'U030',
      status_aktif: 'Aktif'
    }
  ];
}

function buildDummyCustomers_() {
  return [
    {
      kode_customer: 'CUST001',
      nama_customer: 'Toko Maju Jaya',
      tipe_customer: 'Retail',
      kategori_customer: 'Toko',
      alamat: 'Jl. Raya Bekasi No. 12',
      link_google_maps: '',
      pic: 'Pak Joko',
      no_hp: '081300000101',
      latitude: '-6.200000',
      longitude: '106.816666',
      status_customer: 'Aktif',
      status_pembayaran: 'Lancar',
      total_tunggakan: 0,
      jumlah_nota_overdue: 0,
      tanggal_jatuh_tempo_terdekat: '',
      catatan_piutang: '',
      limit_tunggakan: 0,
      catatan: 'Customer lancar'
    },
    {
      kode_customer: 'CUST002',
      nama_customer: 'Agen Sumber Tirta',
      tipe_customer: 'Agen',
      kategori_customer: 'Agen',
      alamat: 'Jl. Industri No. 8',
      link_google_maps: '',
      pic: 'Ibu Rina',
      no_hp: '081300000102',
      latitude: '-6.210000',
      longitude: '106.826666',
      status_customer: 'Menunggak',
      status_pembayaran: 'Menunggak',
      total_tunggakan: 3250000,
      jumlah_nota_overdue: 2,
      tanggal_jatuh_tempo_terdekat: '2026-04-03',
      catatan_piutang: 'Janji bayar hari Jumat',
      limit_tunggakan: 2500000,
      catatan: 'Perlu approval sebelum kirim'
    },
    {
      kode_customer: 'CUST003',
      nama_customer: 'Distributor Amanah',
      tipe_customer: 'Distributor',
      kategori_customer: 'Distributor',
      alamat: 'Jl. Serang Baru No. 21',
      link_google_maps: '',
      pic: 'Pak Hadi',
      no_hp: '081300000103',
      latitude: '-6.220000',
      longitude: '106.836666',
      status_customer: 'Ditahan',
      status_pembayaran: 'Ditahan',
      total_tunggakan: 5100000,
      jumlah_nota_overdue: 3,
      tanggal_jatuh_tempo_terdekat: '2026-03-28',
      catatan_piutang: 'Hold sampai ada pembayaran minimal 50%',
      limit_tunggakan: 3000000,
      catatan: 'Risiko tinggi'
    }
  ];
}

function createDummyTransactions_() {
  var today = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');

  var orderAman = submitSalesOrder({
    sales_id: 'U010',
    sales_nama: 'Andi Sales Dummy',
    jenis_customer: 'Lama',
    customer_id: 'CUST001',
    alamat_kirim: 'Jl. Raya Bekasi No. 12',
    link_google_maps: '',
    latitude: '-6.200000',
    longitude: '106.816666',
    pic_customer: 'Pak Joko',
    no_hp_customer: '081300000101',
    items: [
      {
        nama_item: 'AIRTIS Galon 19L',
        qty: 10,
        harga: 18000,
        diskon: 0,
        subtotal: 180000
      },
      {
        nama_item: 'AIRTIS Cup 220ml - 48',
        qty: 5,
        harga: 15000,
        diskon: 5000,
        subtotal: 70000
      }
    ],
    subtotal: 255000,
    total: 250000,
    term_pembayaran: 'Cash',
    tanggal_jatuh_tempo: today,
    tanggal_kirim_rencana: today,
    catatan: 'Dummy order lancar'
  });

  pauseForDocNumber_();

  var orderTunggakan = submitSalesOrder({
    sales_id: 'U010',
    sales_nama: 'Andi Sales Dummy',
    jenis_customer: 'Lama',
    customer_id: 'CUST002',
    alamat_kirim: 'Jl. Industri No. 8',
    link_google_maps: '',
    latitude: '-6.210000',
    longitude: '106.826666',
    pic_customer: 'Ibu Rina',
    no_hp_customer: '081300000102',
    item: 'AIRTIS Cup 220ml - 48',
    qty: 40,
    harga: 15000,
    diskon: 25000,
    subtotal: 600000,
    total: 575000,
    term_pembayaran: 'Tempo 14 Hari',
    tanggal_jatuh_tempo: today,
    tanggal_kirim_rencana: today,
    catatan: 'Dummy order menunggak'
  });

  pauseForDocNumber_();

  var orderDitahan = submitSalesOrder({
    sales_id: 'U020',
    sales_nama: 'Citra Admin Dummy (CS/Admin)',
    jenis_customer: 'Lama',
    customer_id: 'CUST003',
    alamat_kirim: 'Jl. Serang Baru No. 21',
    link_google_maps: '',
    latitude: '-6.220000',
    longitude: '106.836666',
    pic_customer: 'Pak Hadi',
    no_hp_customer: '081300000103',
    item: 'AIRTIS Botol 600ml',
    qty: 25,
    harga: 22000,
    diskon: 0,
    subtotal: 550000,
    total: 550000,
    term_pembayaran: 'Tempo 30 Hari',
    tanggal_jatuh_tempo: today,
    tanggal_kirim_rencana: today,
    catatan: '[AGEN/CS] Dummy order customer ditahan'
  });

  pauseForDocNumber_();

  var orderKirim = submitSalesOrder({
    sales_id: 'U010',
    sales_nama: 'Andi Sales Dummy',
    jenis_customer: 'Lama',
    customer_id: 'CUST001',
    alamat_kirim: 'Jl. Raya Bekasi No. 12',
    link_google_maps: '',
    latitude: '-6.200000',
    longitude: '106.816666',
    pic_customer: 'Pak Joko',
    no_hp_customer: '081300000101',
    items: [
      {
        nama_item: 'AIRTIS Botol 330ml',
        qty: 15,
        harga: 21000,
        diskon: 5000,
        subtotal: 310000
      },
      {
        nama_item: 'AIRTIS Botol 600ml',
        qty: 8,
        harga: 22000,
        diskon: 0,
        subtotal: 176000
      }
    ],
    subtotal: 491000,
    total: 486000,
    term_pembayaran: 'Cash',
    tanggal_jatuh_tempo: today,
    tanggal_kirim_rencana: today,
    catatan: 'Dummy order siap dibuatkan SJ'
  });

  pauseForDocNumber_();

  createSuratJalan(orderAman.no_so, {
    driver: 'Dedi',
    armada: 'B 9123 KXA',
    catatan_kirim: 'Dummy surat jalan siap kirim'
  });

  pauseForDocNumber_();

  createSuratJalan(orderKirim.no_so, {
    driver: 'Roni',
    armada: 'B 8899 TIR',
    catatan_kirim: 'Dummy surat jalan untuk testing status'
  });

  markOrderDelivered(orderKirim.no_so, 'U020', 'Dummy barang terkirim');
  verifyDeliveredOrder(orderKirim.no_so, 'U020', {
    items: getSalesOrderDetailsByNoSo_(orderKirim.no_so).map(function(detail) {
      return {
        detail_id: detail.detail_id,
        qty_terkirim: Number(detail.qty || 0),
        harga_final: Number(detail.harga || 0),
        diskon_final: Number(detail.diskon || 0),
        subtotal_final: Number(detail.subtotal || 0)
      };
    }),
    catatan_verifikasi_cs: 'Dummy verifikasi CS selesai'
  });
  completeOrder(orderKirim.no_so, 'U020', 'Dummy order selesai');

  return {
    order_aman: orderAman.no_so,
    order_tunggakan: orderTunggakan.no_so,
    order_ditahan: orderDitahan.no_so,
    order_kirim: orderKirim.no_so
  };
}

function pauseForDocNumber_() {
  Utilities.sleep(1100);
}
