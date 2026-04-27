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
      'KOMISI_SLF_MUTASI',
      'KOMISI_PAYOUT',
      'KOMISI_PAYOUT_DETAIL',
      'CATATAN_MONITORING'
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

function seedDummyYasserMapData24() {
  setupAppSheets();
  ensureFieldActivitySheets_();

  var dateKey = '2026-04-24';
  var salesId = 'SLF101';
  var salesName = 'Yasser';
  var existingRows = getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI).filter(function(row) {
    return normalizeSheetDateToYmd_(row.tanggal) === dateKey &&
      String(row.user_id || '').trim() === salesId &&
      String(row.aktivitas_id || '').indexOf('ACT-YASSER-2404-') === 0;
  });
  var points;
  var previousPoint = null;

  if (existingRows.length) {
    return {
      success: true,
      message: 'Dummy Yasser tanggal 24 sudah ada.',
      rows: existingRows.length
    };
  }

  points = [
    {
      suffix: '0800',
      time: '08:05:00',
      target: '08:00',
      type: 'Mulai Kerja',
      statusTime: 'Tepat Waktu',
      lat: -6.175392,
      lng: 106.827153,
      customerId: '',
      customerName: '',
      activity: '',
      result: 'Mulai dari area Monas/Gambir',
      locationStatus: 'Normal',
      deviceStatus: 'Perangkat Utama'
    },
    {
      suffix: '1000',
      time: '10:18:00',
      target: '10:00',
      type: 'Update Lokasi',
      statusTime: 'Terlambat',
      lat: -6.196876,
      lng: 106.823213,
      customerId: 'CUST001',
      customerName: 'Toko Maju Jaya',
      activity: 'Kunjungan Customer',
      result: 'Follow up order galon',
      locationStatus: 'Normal',
      deviceStatus: 'Perangkat Utama'
    },
    {
      suffix: '1200',
      time: '12:04:00',
      target: '12:00',
      type: 'Catat Kegiatan',
      statusTime: 'Tepat Waktu',
      lat: -6.214622,
      lng: 106.845130,
      customerId: 'CUST002',
      customerName: 'Agen Sumber Tirta',
      activity: 'Penagihan',
      result: 'Customer janji bayar sore',
      locationStatus: 'Normal',
      deviceStatus: 'Perangkat Utama'
    },
    {
      suffix: '1400',
      time: '14:28:00',
      target: '14:00',
      type: 'Update Lokasi',
      statusTime: 'Terlambat',
      lat: -6.238270,
      lng: 106.975573,
      customerId: '',
      customerName: 'Prospek Bekasi Barat',
      activity: 'Survey Customer Baru',
      result: 'Survey toko prospek',
      locationStatus: 'Perlu Cek',
      deviceStatus: 'Perangkat Utama'
    },
    {
      suffix: '1600',
      time: '16:03:00',
      target: '16:00',
      type: 'Selesai Kerja',
      statusTime: 'Tepat Waktu',
      lat: -6.242311,
      lng: 106.992416,
      customerId: '',
      customerName: '',
      activity: '',
      result: 'Selesai area Bekasi Barat',
      locationStatus: 'Normal',
      deviceStatus: 'Perangkat Utama'
    }
  ];

  points.forEach(function(point) {
    var movement = previousPoint
      ? buildDummyMovementFromPrevious_(previousPoint, point)
      : { distanceKm: '', minutes: '', speedKmh: '' };

    appendRowByHeaders_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI, {
      aktivitas_id: 'ACT-YASSER-2404-' + point.suffix,
      tanggal: dateKey,
      user_id: salesId,
      nama_user: salesName,
      tipe_sales: 'Freelance',
      tipe_aktivitas: point.type,
      jenis_kegiatan: point.activity,
      customer_id: point.customerId,
      nama_customer: point.customerName,
      hasil_kegiatan: point.result,
      target_jam: point.target,
      jam_server: point.time,
      status_waktu: point.statusTime,
      latitude: point.lat,
      longitude: point.lng,
      akurasi_gps: 18,
      link_maps: buildGoogleMapsLink_(point.lat, point.lng),
      jarak_km_dari_sebelumnya: movement.distanceKm,
      menit_dari_sebelumnya: movement.minutes,
      estimasi_kecepatan_kmh: movement.speedKmh,
      status_lokasi: point.locationStatus,
      device_id: 'dummy-yasser-phone',
      user_agent: 'Dummy seed Apps Script',
      status_perangkat: point.deviceStatus,
      catatan: 'Dummy histori maps untuk MTR',
      created_at: dateKey + 'T' + point.time
    });

    previousPoint = point;
  });

  return {
    success: true,
    message: 'Dummy histori maps Yasser tanggal 24 berhasil dibuat.',
    rows: points.length
  };
}

function seedDummyAllSalesBogorMapData23() {
  setupAppSheets();
  ensureFieldActivitySheets_();

  var dateKey = '2026-04-23';
  var activeSales = getActiveSalesUsers_();
  var existingRows = getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI);
  var routeTemplates = getDummyBogorRouteTemplates23_();
  var rowsCreated = 0;
  var skippedSales = 0;

  activeSales.forEach(function(sales, salesIndex) {
    var salesId = String(sales.user_id || '').trim();
    var salesName = String(sales.nama_user || sales.nama || salesId || '').trim();
    var salesKey = sanitizeDummyIdPart_(salesId || salesName || String(salesIndex + 1));
    var activityPrefix = 'ACT-BOGOR-2304-' + salesKey + '-';
    var route = routeTemplates[salesIndex % routeTemplates.length];
    var existingForSales = existingRows.filter(function(row) {
      return normalizeSheetDateToYmd_(row.tanggal) === dateKey &&
        String(row.user_id || '').trim() === salesId &&
        String(row.aktivitas_id || '').indexOf(activityPrefix) === 0;
    });
    var previousPoint = null;

    if (!salesId || existingForSales.length) {
      skippedSales += 1;
      return;
    }

    route.points.forEach(function(point, pointIndex) {
      var movement = previousPoint
        ? buildDummyMovementFromPrevious_(previousPoint, point)
        : { distanceKm: '', minutes: '', speedKmh: '' };
      var isCheckPoint = pointIndex === 3 && salesIndex % 3 === 0;
      var isLatePoint = (pointIndex === 1 && salesIndex % 2 === 0) || (pointIndex === 3 && salesIndex % 4 === 0);

      appendRowByHeaders_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI, {
        aktivitas_id: activityPrefix + point.suffix,
        tanggal: dateKey,
        user_id: salesId,
        nama_user: salesName,
        tipe_sales: sales.tipe_sales || '',
        tipe_aktivitas: point.type,
        jenis_kegiatan: point.activity,
        customer_id: '',
        nama_customer: point.customerName,
        hasil_kegiatan: point.result + ' - ' + route.area,
        target_jam: point.target,
        jam_server: point.time,
        status_waktu: isLatePoint ? 'Terlambat' : 'Tepat Waktu',
        latitude: point.lat,
        longitude: point.lng,
        akurasi_gps: 16 + (pointIndex % 3),
        link_maps: buildGoogleMapsLink_(point.lat, point.lng),
        jarak_km_dari_sebelumnya: movement.distanceKm,
        menit_dari_sebelumnya: movement.minutes,
        estimasi_kecepatan_kmh: movement.speedKmh,
        status_lokasi: isCheckPoint ? 'Perlu Cek' : 'Normal',
        device_id: 'dummy-bogor-' + salesKey,
        user_agent: 'Dummy seed Apps Script',
        status_perangkat: salesIndex % 5 === 0 ? 'Perangkat Baru' : 'Perangkat Utama',
        catatan: 'Dummy histori maps semua sales area Bogor tanggal 23',
        created_at: dateKey + 'T' + point.time
      });

      previousPoint = point;
      rowsCreated += 1;
    });
  });

  return {
    success: true,
    message: 'Dummy histori maps semua sales tanggal 23 area Kota/Kabupaten Bogor selesai dibuat.',
    tanggal: dateKey,
    sales: activeSales.length,
    sales_dilewati: skippedSales,
    rows: rowsCreated
  };
}

function getDummyBogorRouteTemplates23_() {
  return [
    {
      area: 'Kota Bogor - Bogor Tengah, Baranangsiang, Tanah Sareal',
      points: [
        buildDummyBogorPoint23_('0800', '08:03:00', '08:00', 'Mulai Kerja', '', 'Kantor/area awal Bogor Tengah', 'Mulai dari area Kebun Raya Bogor', -6.595038, 106.816635),
        buildDummyBogorPoint23_('1000', '10:16:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Suryakencana', 'Follow up kebutuhan galon', -6.601650, 106.806390),
        buildDummyBogorPoint23_('1200', '12:05:00', '12:00', 'Catat Kegiatan', 'Kunjungan Customer', 'Outlet Baranangsiang', 'Cek stok dan catat order', -6.613850, 106.810120),
        buildDummyBogorPoint23_('1400', '14:21:00', '14:00', 'Update Lokasi', 'Survey Customer Baru', 'Prospek Tanah Sareal', 'Survey titik prospek warung', -6.570950, 106.793570),
        buildDummyBogorPoint23_('1600', '16:02:00', '16:00', 'Selesai Kerja', '', 'Area Kedung Badak', 'Selesai rute Kota Bogor', -6.556420, 106.806260)
      ]
    },
    {
      area: 'Kabupaten Bogor - Cibinong, Bojonggede',
      points: [
        buildDummyBogorPoint23_('0800', '08:01:00', '08:00', 'Mulai Kerja', '', 'Cibinong City Mall area', 'Mulai dari Cibinong', -6.485210, 106.842430),
        buildDummyBogorPoint23_('1000', '10:08:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Pakansari', 'Kunjungan outlet reguler', -6.469760, 106.854520),
        buildDummyBogorPoint23_('1200', '12:03:00', '12:00', 'Catat Kegiatan', 'Penagihan', 'Customer Cibinong Barat', 'Follow up tagihan dan order', -6.498620, 106.831930),
        buildDummyBogorPoint23_('1400', '14:06:00', '14:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Bojonggede', 'Cek stok display', -6.527180, 106.806990),
        buildDummyBogorPoint23_('1600', '16:04:00', '16:00', 'Selesai Kerja', '', 'Area Tajur Halang', 'Selesai rute Cibinong', -6.543210, 106.789820)
      ]
    },
    {
      area: 'Kabupaten Bogor - Dramaga, Ciampea',
      points: [
        buildDummyBogorPoint23_('0800', '08:04:00', '08:00', 'Mulai Kerja', '', 'Dramaga', 'Mulai dari area Dramaga', -6.559720, 106.725410),
        buildDummyBogorPoint23_('1000', '10:14:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Laladon', 'Kunjungan outlet sekitar Laladon', -6.571680, 106.746300),
        buildDummyBogorPoint23_('1200', '12:06:00', '12:00', 'Catat Kegiatan', 'Survey Customer Baru', 'Prospek Bubulak', 'Catat prospek toko baru', -6.587140, 106.755660),
        buildDummyBogorPoint23_('1400', '14:10:00', '14:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Ciampea', 'Follow up order depot', -6.605370, 106.698620),
        buildDummyBogorPoint23_('1600', '16:01:00', '16:00', 'Selesai Kerja', '', 'Area Cibungbulang', 'Selesai rute Dramaga', -6.621900, 106.704510)
      ]
    },
    {
      area: 'Kabupaten Bogor - Ciawi, Megamendung',
      points: [
        buildDummyBogorPoint23_('0800', '08:02:00', '08:00', 'Mulai Kerja', '', 'Ciawi', 'Mulai dari area Ciawi', -6.650520, 106.850240),
        buildDummyBogorPoint23_('1000', '10:11:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Cisarua bawah', 'Kunjungan outlet rute puncak bawah', -6.667320, 106.865500),
        buildDummyBogorPoint23_('1200', '12:07:00', '12:00', 'Catat Kegiatan', 'Kunjungan Customer', 'Outlet Megamendung', 'Cek stok customer', -6.681980, 106.890640),
        buildDummyBogorPoint23_('1400', '14:25:00', '14:00', 'Update Lokasi', 'Survey Customer Baru', 'Prospek Gadog', 'Survey titik pelanggan baru', -6.698870, 106.914760),
        buildDummyBogorPoint23_('1600', '16:05:00', '16:00', 'Selesai Kerja', '', 'Area Megamendung atas', 'Selesai rute Ciawi', -6.713450, 106.932300)
      ]
    },
    {
      area: 'Kabupaten Bogor - Parung, Kemang',
      points: [
        buildDummyBogorPoint23_('0800', '08:03:00', '08:00', 'Mulai Kerja', '', 'Parung', 'Mulai dari area Parung', -6.421880, 106.733620),
        buildDummyBogorPoint23_('1000', '10:09:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Kemang', 'Kunjungan customer Kemang', -6.452410, 106.748930),
        buildDummyBogorPoint23_('1200', '12:04:00', '12:00', 'Catat Kegiatan', 'Penagihan', 'Customer Jampang', 'Follow up pembayaran', -6.481220, 106.759870),
        buildDummyBogorPoint23_('1400', '14:12:00', '14:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Rancabungur', 'Cek kebutuhan stok', -6.502100, 106.739520),
        buildDummyBogorPoint23_('1600', '16:03:00', '16:00', 'Selesai Kerja', '', 'Area Atang Senjaya', 'Selesai rute Parung', -6.526900, 106.726300)
      ]
    },
    {
      area: 'Kabupaten Bogor - Jonggol, Cileungsi',
      points: [
        buildDummyBogorPoint23_('0800', '08:02:00', '08:00', 'Mulai Kerja', '', 'Jonggol', 'Mulai dari area Jonggol', -6.466280, 107.006720),
        buildDummyBogorPoint23_('1000', '10:13:00', '10:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Cileungsi timur', 'Kunjungan outlet reguler', -6.435820, 106.963400),
        buildDummyBogorPoint23_('1200', '12:08:00', '12:00', 'Catat Kegiatan', 'Survey Customer Baru', 'Prospek Cileungsi', 'Survey customer baru', -6.398990, 106.956830),
        buildDummyBogorPoint23_('1400', '14:18:00', '14:00', 'Update Lokasi', 'Kunjungan Customer', 'Outlet Klapanunggal', 'Follow up order', -6.427110, 106.910230),
        buildDummyBogorPoint23_('1600', '16:06:00', '16:00', 'Selesai Kerja', '', 'Area Gunung Putri', 'Selesai rute Cileungsi', -6.471300, 106.903950)
      ]
    }
  ];
}

function buildDummyBogorPoint23_(suffix, time, target, type, activity, customerName, result, lat, lng) {
  return {
    suffix: suffix,
    time: time,
    target: target,
    type: type,
    activity: activity,
    customerName: customerName,
    result: result,
    lat: lat,
    lng: lng
  };
}

function sanitizeDummyIdPart_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'SALES';
}

function buildDummyMovementFromPrevious_(previousPoint, point) {
  var minutes = Math.max(1, timeToMinutes_(point.time) - timeToMinutes_(previousPoint.time));
  var distanceKm = calculateDistanceKm_(previousPoint.lat, previousPoint.lng, point.lat, point.lng);
  var speed = (distanceKm / minutes) * 60;

  return {
    distanceKm: roundNumber_(distanceKm, 2),
    minutes: minutes,
    speedKmh: roundNumber_(speed, 1)
  };
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
    APP_CONFIG.SHEETS.KOMISI_PAYOUT,
    APP_CONFIG.SHEETS.KOMISI_PAYOUT_DETAIL,
    APP_CONFIG.SHEETS.TAGIHAN,
    APP_CONFIG.SHEETS.KPI_TARGET_SALES,
    APP_CONFIG.SHEETS.KPI_LOG,
    APP_CONFIG.SHEETS.AKTIVITAS_LOKASI,
    APP_CONFIG.SHEETS.PERANGKAT_USER,
    APP_CONFIG.SHEETS.CATATAN_MONITORING
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
    APP_CONFIG.SHEETS.KOMISI_PAYOUT,
    APP_CONFIG.SHEETS.KOMISI_PAYOUT_DETAIL,
    APP_CONFIG.SHEETS.TAGIHAN,
    APP_CONFIG.SHEETS.KPI_TARGET_SALES,
    APP_CONFIG.SHEETS.KPI_LOG,
    APP_CONFIG.SHEETS.AKTIVITAS_LOKASI,
    APP_CONFIG.SHEETS.PERANGKAT_USER,
    APP_CONFIG.SHEETS.CATATAN_MONITORING
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
    'Approver': { user_id: 'U003', nama_user: 'Bima Approver', role: 'Approver' },
    'MTR': { user_id: 'MTR1', nama_user: 'Monitoring', role: 'MTR' },
    'Controller': { user_id: 'CTR1', nama_user: 'Controller', role: 'Controller' }
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
      status_aktif: 'Aktif',
      bank_nama: 'BCA',
      bank_no_rekening: '1234567890',
      bank_nama_pemilik: activeUser.nama_user
    },
    {
      user_id: 'U010',
      nama_user: 'Andi Sales Dummy',
      role: 'Sales',
      no_hp: '081300000010',
      email: 'sales.dummy@airtis.local',
      password: 'U010',
      status_aktif: 'Aktif',
      bank_nama: 'Mandiri',
      bank_no_rekening: '9876543210',
      bank_nama_pemilik: 'Andi Sales Dummy'
    },
    {
      user_id: 'U020',
      nama_user: 'Citra Admin Dummy',
      role: 'CS/Admin',
      no_hp: '081300000020',
      email: 'admin.dummy@airtis.local',
      password: 'U020',
      status_aktif: 'Aktif',
      bank_nama: '',
      bank_no_rekening: '',
      bank_nama_pemilik: ''
    },
    {
      user_id: 'U030',
      nama_user: 'Bima Approver Dummy',
      role: 'Approver',
      no_hp: '081300000030',
      email: 'approver.dummy@airtis.local',
      password: 'U030',
      status_aktif: 'Aktif',
      bank_nama: '',
      bank_no_rekening: '',
      bank_nama_pemilik: ''
    },
    {
      user_id: 'MTR1',
      nama_user: 'Monitoring',
      role: 'MTR',
      no_hp: '081300000032',
      email: 'monitoring.dummy@airtis.local',
      password: 'MTR1',
      status_aktif: 'Aktif',
      bank_nama: '',
      bank_no_rekening: '',
      bank_nama_pemilik: ''
    },
    {
      user_id: 'CTR1',
      nama_user: 'Controller',
      role: 'Controller',
      no_hp: '081300000031',
      email: 'controller.dummy@airtis.local',
      password: 'CTR1',
      status_aktif: 'Aktif',
      bank_nama: '',
      bank_no_rekening: '',
      bank_nama_pemilik: ''
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
