var FIELD_ACTIVITY_CONFIG = {
  START_TIME: '08:00',
  UPDATE_TIMES: ['10:00', '12:00', '14:00'],
  END_TIME: '16:00',
  SATURDAY_END_TIME: '14:00',
  TOLERANCE_MINUTES: 30,
  SPEED_CHECK_KMH: 80,
  SPEED_SUSPICIOUS_KMH: 120
};

function getFieldActivityTodayForSales(userId) {
  var user = requireCurrentUserRole_(['Sales'], userId);
  ensureFieldActivitySheets_();

  return toClientValue_({
    config: getFieldActivityClientConfig_(),
    today: buildFieldActivityDailyRow_(user, getNowParts_().tanggal),
    recentLogs: getFieldActivityLogsForUserDate_(user.user_id, getNowParts_().tanggal).slice(-8).reverse()
  });
}

function recordFieldActivityFromDashboard(userId, payload) {
  var user = requireCurrentUserRole_(['Sales'], userId);
  var source = payload || {};
  var location = source.location || {};
  var activityType = normalizeFieldActivityType_(source.tipe_aktivitas || source.activity_type || '');
  var now = getNowParts_();
  var deviceId = String(source.device_id || '').trim();
  var userAgent = String(source.user_agent || '').trim();
  var latitude = Number(location.latitude);
  var longitude = Number(location.longitude);
  var accuracy = Number(location.accuracy || 0);
  var targetTime;
  var timeStatus;
  var previousLog;
  var movement;
  var deviceStatus;
  var activityRow;

  ensureFieldActivitySheets_();

  if (!activityType) {
    throw new Error('Tipe aktivitas tidak dikenali.');
  }

  if (isNaN(latitude) || isNaN(longitude)) {
    throw new Error('Lokasi GPS belum terbaca. Aktifkan izin lokasi lalu coba lagi.');
  }

  targetTime = getFieldActivityTargetTime_(activityType, now.jam, now.tanggal);
  validateFieldActivitySingleEntry_(user.user_id, now.tanggal, activityType, targetTime);
  timeStatus = getFieldActivityTimeStatus_(targetTime, now.jam);
  previousLog = getPreviousFieldActivityLog_(user.user_id, now.tanggal);
  movement = buildFieldActivityMovementStatus_(previousLog, latitude, longitude, now);
  deviceStatus = registerFieldActivityDevice_(user, deviceId, userAgent, now);

  activityRow = {
    aktivitas_id: generateDocNumber_('ACT'),
    tanggal: now.tanggal,
    user_id: user.user_id || '',
    nama_user: user.nama_user || '',
    tipe_sales: user.tipe_sales || '',
    tipe_aktivitas: activityType,
    jenis_kegiatan: String(source.jenis_kegiatan || '').trim(),
    customer_id: String(source.customer_id || '').trim(),
    nama_customer: String(source.nama_customer || '').trim(),
    hasil_kegiatan: String(source.hasil_kegiatan || '').trim(),
    target_jam: targetTime,
    jam_server: now.jam,
    status_waktu: timeStatus,
    latitude: latitude,
    longitude: longitude,
    akurasi_gps: accuracy || '',
    link_maps: buildGoogleMapsLink_(latitude, longitude),
    jarak_km_dari_sebelumnya: movement.distanceKm,
    menit_dari_sebelumnya: movement.minutes,
    estimasi_kecepatan_kmh: movement.speedKmh,
    status_lokasi: movement.status,
    device_id: deviceId,
    user_agent: userAgent,
    status_perangkat: deviceStatus,
    catatan: String(source.catatan || '').trim(),
    created_at: nowIso_()
  };

  appendRowByHeaders_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI, activityRow);

  return toClientValue_({
    success: true,
    message: activityType + ' tersimpan pukul ' + now.jam + '.',
    row: activityRow,
    today: buildFieldActivityDailyRow_(user, now.tanggal),
    recentLogs: getFieldActivityLogsForUserDate_(user.user_id, now.tanggal).slice(-8).reverse()
  });
}

function validateFieldActivitySingleEntry_(userId, dateKey, activityType, targetTime) {
  var isSingleEntry = activityType === 'Mulai Kerja' || activityType === 'Selesai Kerja';
  var logs;
  var exists;

  if (!isSingleEntry) {
    return;
  }

  logs = getFieldActivityLogsForUserDate_(userId, dateKey);
  exists = logs.some(function(log) {
    return normalizeText_(log.tipe_aktivitas) === normalizeText_(activityType) ||
      (targetTime && normalizeFieldActivityTargetTime_(log.target_jam) === targetTime);
  });

  if (exists) {
    throw new Error(activityType + ' hari ini sudah tercatat.');
  }
}

function getFieldMonitoringDashboardData(userId, filters) {
  requireCurrentUserRole_(['Controller', 'MTR', 'CS/Admin', 'Approver'], userId);
  ensureFieldActivitySheets_();

  var dateKey = normalizeSheetDateToYmd_((filters || {}).tanggal || '') || getNowParts_().tanggal;
  var notesBySales = getMonitoringNotesBySalesForDate_(dateKey);
  var rows = getActiveSalesUsers_().map(function(user) {
    var row = buildFieldActivityDailyRow_(user, dateKey);
    var salesNotes = notesBySales[String(user.user_id || '').trim()] || [];
    var latestNote = salesNotes.length ? salesNotes[salesNotes.length - 1] : null;

    row.total_catatan_mtr = salesNotes.length;
    row.catatan_mtr_terakhir = latestNote ? String(latestNote.catatan_mtr || '').trim() : '';
    row.status_monitoring_mtr = latestNote ? String(latestNote.status_monitoring || '').trim() : '';
    row.kategori_temuan_mtr = latestNote ? String(latestNote.kategori_temuan || '').trim() : '';

    return row;
  });
  var weekly = buildFieldActivityWeeklyAnomalies_(dateKey);

  return toClientValue_({
    tanggal: dateKey,
    summary: buildFieldMonitoringSummary_(rows),
    rows: rows,
    weekly: weekly
  });
}

function ensureFieldActivitySheets_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI, APP_CONFIG.HEADERS.AKTIVITAS_LOKASI);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.PERANGKAT_USER, APP_CONFIG.HEADERS.PERANGKAT_USER);
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.CATATAN_MONITORING, APP_CONFIG.HEADERS.CATATAN_MONITORING);
}

function createMonitoringNoteFromDashboard(userId, payload) {
  var user = requireCurrentUserRole_(['MTR', 'Controller'], userId);
  var source = payload || {};
  var salesId = String(source.sales_id || '').trim();
  var dateKey = normalizeSheetDateToYmd_(source.tanggal || '') || getNowParts_().tanggal;
  var note = String(source.catatan_mtr || source.catatan || '').trim();
  var salesUser;
  var row;

  ensureFieldActivitySheets_();

  if (!salesId) {
    throw new Error('Sales wajib dipilih untuk catatan monitoring.');
  }

  if (!note) {
    throw new Error('Catatan monitoring wajib diisi.');
  }

  salesUser = getActiveSalesUsers_().find(function(item) {
    return String(item.user_id || '').trim() === salesId;
  });

  if (!salesUser) {
    throw new Error('Sales tidak ditemukan atau tidak aktif.');
  }

  row = {
    catatan_id: generateDocNumber_('MTR'),
    tanggal: dateKey,
    sales_id: salesId,
    nama_sales: salesUser.nama_user || '',
    status_monitoring: String(source.status_monitoring || 'Perlu Follow Up').trim(),
    kategori_temuan: String(source.kategori_temuan || 'Catatan Manual').trim(),
    catatan_mtr: note,
    tindak_lanjut: String(source.tindak_lanjut || '').trim(),
    dibuat_oleh: user.user_id || '',
    nama_mtr: user.nama_user || '',
    created_at: nowIso_()
  };

  appendRowByHeaders_(APP_CONFIG.SHEETS.CATATAN_MONITORING, row);

  return toClientValue_({
    success: true,
    message: 'Catatan monitoring tersimpan untuk ' + (salesUser.nama_user || salesId) + '.',
    row: row
  });
}

function getMtrSalesMovementMapData(userId, filters) {
  requireCurrentUserRole_(['MTR', 'Controller'], userId);
  ensureFieldActivitySheets_();

  var payload = filters || {};
  var dateKey = normalizeSheetDateToYmd_(payload.tanggal || '') || getNowParts_().tanggal;
  var requestedDateKey = dateKey;
  var salesFilter = String(payload.sales_id || '').trim();
  var activeSales = getActiveSalesUsers_();
  var activeSalesById = {};
  var pointsBySales = {};
  var rows;

  activeSales.forEach(function(user) {
    var userIdKey = String(user.user_id || '').trim();
    if (userIdKey) {
      activeSalesById[userIdKey] = user;
    }
  });

  rows = getMtrSalesMovementRowsForDate_(dateKey, salesFilter, activeSalesById);

  if (!rows.length && String(payload.fallback_latest || '').trim() === 'Ya') {
    dateKey = findLatestMtrSalesMovementDate_(salesFilter, activeSalesById) || dateKey;
    if (dateKey !== requestedDateKey) {
      rows = getMtrSalesMovementRowsForDate_(dateKey, salesFilter, activeSalesById);
    }
  }

  rows.forEach(function(row) {
    var rowSalesId = String(row.user_id || '').trim();
    var point = buildMtrSalesMovementPoint_(row);

    if (!pointsBySales[rowSalesId]) {
      pointsBySales[rowSalesId] = {
        sales_id: rowSalesId,
        nama_sales: String(row.nama_user || activeSalesById[rowSalesId].nama_user || '').trim(),
        tipe_sales: String(row.tipe_sales || activeSalesById[rowSalesId].tipe_sales || '').trim(),
        channel_sales: String(activeSalesById[rowSalesId].channel_sales_default || '').trim(),
        points: []
      };
    }

    pointsBySales[rowSalesId].points.push(point);
  });

  return toClientValue_({
    tanggal: dateKey,
    requested_tanggal: requestedDateKey,
    fallback_used: dateKey !== requestedDateKey,
    sales_id: salesFilter,
    salesOptions: activeSales.map(function(user) {
      return {
        sales_id: String(user.user_id || '').trim(),
        nama_sales: String(user.nama_user || '').trim(),
        tipe_sales: String(user.tipe_sales || '').trim(),
        channel_sales: String(user.channel_sales_default || '').trim()
      };
    }),
    routes: Object.keys(pointsBySales).map(function(key) {
      var route = pointsBySales[key];
      route.total_points = route.points.length;
      route.total_distance_km = calculateMtrRouteDistanceKm_(route.points);
      route.total_anomali = route.points.filter(function(point) {
        return point.severity === 'alert' || point.severity === 'warning';
      }).length;
      return route;
    })
  });
}

function getMtrSalesMovementRowsForDate_(dateKey, salesFilter, activeSalesById) {
  return getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI).filter(function(row) {
    var rowSalesId = String(row.user_id || '').trim();
    var latitude = Number(row.latitude);
    var longitude = Number(row.longitude);

    return normalizeSheetDateToYmd_(row.tanggal) === dateKey &&
      !!activeSalesById[rowSalesId] &&
      (!salesFilter || rowSalesId === salesFilter) &&
      !isNaN(latitude) &&
      !isNaN(longitude);
  }).sort(function(left, right) {
    var leftSales = String(left.user_id || '').trim();
    var rightSales = String(right.user_id || '').trim();
    var salesCompare = leftSales.localeCompare(rightSales, 'id-ID');

    if (salesCompare !== 0) {
      return salesCompare;
    }

    return normalizeFieldActivityTime_(left.jam_server).localeCompare(normalizeFieldActivityTime_(right.jam_server));
  });
}

function findLatestMtrSalesMovementDate_(salesFilter, activeSalesById) {
  var latest = '';

  getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI).forEach(function(row) {
    var rowSalesId = String(row.user_id || '').trim();
    var latitude = Number(row.latitude);
    var longitude = Number(row.longitude);
    var rowDate = normalizeSheetDateToYmd_(row.tanggal);

    if (!rowDate || !activeSalesById[rowSalesId] || (salesFilter && rowSalesId !== salesFilter) || isNaN(latitude) || isNaN(longitude)) {
      return;
    }

    if (!latest || rowDate > latest) {
      latest = rowDate;
    }
  });

  return latest;
}

function buildMtrSalesMovementPoint_(row) {
  var latitude = Number(row.latitude);
  var longitude = Number(row.longitude);
  var statusWaktu = getFieldActivityEffectiveTimeStatus_(row);
  var statusLokasi = String(row.status_lokasi || '').trim();
  var statusPerangkat = String(row.status_perangkat || '').trim();
  var activityType = String(row.tipe_aktivitas || '').trim();

  return {
    aktivitas_id: String(row.aktivitas_id || '').trim(),
    tanggal: normalizeSheetDateToYmd_(row.tanggal),
    jam: normalizeFieldActivityTime_(row.jam_server),
    tipe_aktivitas: activityType,
    jenis_kegiatan: String(row.jenis_kegiatan || '').trim(),
    customer_id: String(row.customer_id || '').trim(),
    nama_customer: String(row.nama_customer || '').trim(),
    hasil_kegiatan: String(row.hasil_kegiatan || '').trim(),
    latitude: latitude,
    longitude: longitude,
    link_maps: String(row.link_maps || buildGoogleMapsLink_(latitude, longitude)).trim(),
    status_waktu: statusWaktu,
    status_lokasi: statusLokasi,
    status_perangkat: statusPerangkat,
    jarak_km_dari_sebelumnya: row.jarak_km_dari_sebelumnya || '',
    estimasi_kecepatan_kmh: row.estimasi_kecepatan_kmh || '',
    severity: getMtrMovementSeverity_(statusWaktu, statusLokasi, statusPerangkat, activityType)
  };
}

function getFieldActivityEffectiveTimeStatus_(row) {
  var targetTime = normalizeFieldActivityTargetTime_(row && row.target_jam);
  var actualTime = normalizeFieldActivityTime_(row && row.jam_server);

  if (!targetTime || !actualTime) {
    return String(row && row.status_waktu || '').trim();
  }

  return getFieldActivityTimeStatus_(targetTime, actualTime);
}

function getMtrMovementSeverity_(statusWaktu, statusLokasi, statusPerangkat, activityType) {
  var locationKey = normalizeText_(statusLokasi);
  var deviceKey = normalizeText_(statusPerangkat);
  var timeKey = normalizeText_(statusWaktu);

  if (locationKey === 'mencurigakan' || deviceKey === 'sering ganti perangkat') {
    return 'alert';
  }

  if (locationKey === 'perlu cek' || deviceKey === 'perangkat baru' || timeKey === 'terlambat' || timeKey === 'lebih awal') {
    return 'warning';
  }

  if (normalizeText_(activityType) === 'catat kegiatan') {
    return 'work';
  }

  return 'normal';
}

function calculateMtrRouteDistanceKm_(points) {
  var total = 0;
  var index;
  var previous;
  var current;

  for (index = 1; index < (points || []).length; index += 1) {
    previous = points[index - 1];
    current = points[index];
    total += calculateDistanceKm_(previous.latitude, previous.longitude, current.latitude, current.longitude);
  }

  return roundNumber_(total, 2);
}

function getMonitoringNotesBySalesForDate_(dateKey) {
  var result = {};

  getSheetData_(APP_CONFIG.SHEETS.CATATAN_MONITORING).forEach(function(row) {
    var rowDate = normalizeSheetDateToYmd_(row.tanggal);
    var salesId = String(row.sales_id || '').trim();

    if (rowDate !== dateKey || !salesId) {
      return;
    }

    if (!result[salesId]) {
      result[salesId] = [];
    }

    result[salesId].push(row);
  });

  Object.keys(result).forEach(function(salesId) {
    result[salesId].sort(function(left, right) {
      return String(left.created_at || '').localeCompare(String(right.created_at || ''));
    });
  });

  return result;
}

function getFieldActivityClientConfig_() {
  var today = getNowParts_().tanggal;
  var schedule = getFieldActivityScheduleForDate_(today);

  return {
    startTime: schedule.startTime,
    updateTimes: schedule.updateTimes,
    endTime: schedule.endTime,
    toleranceMinutes: FIELD_ACTIVITY_CONFIG.TOLERANCE_MINUTES,
    scheduleText: schedule.text,
    isWorkday: schedule.isWorkday
  };
}

function normalizeFieldActivityType_(value) {
  var key = normalizeText_(value);

  if (key === 'mulai kerja' || key === 'mulai') {
    return 'Mulai Kerja';
  }

  if (key === 'update lokasi' || key === 'update') {
    return 'Update Lokasi';
  }

  if (key === 'selesai kerja' || key === 'selesai') {
    return 'Selesai Kerja';
  }

  if (key === 'catat kegiatan' || key === 'kegiatan' || key === 'kunjungan customer') {
    return 'Catat Kegiatan';
  }

  return '';
}

function getFieldActivityTargetTime_(activityType, currentTime, dateKey) {
  var schedule = getFieldActivityScheduleForDate_(dateKey);

  if (activityType === 'Mulai Kerja') {
    return schedule.startTime;
  }

  if (activityType === 'Selesai Kerja') {
    return schedule.endTime;
  }

  if (activityType === 'Catat Kegiatan') {
    return '';
  }

  return getNearestFieldUpdateTime_(currentTime, schedule.updateTimes);
}

function getNearestFieldUpdateTime_(currentTime, updateTimes) {
  var currentMinutes = timeToMinutes_(currentTime);
  var times = Array.isArray(updateTimes) && updateTimes.length ? updateTimes : FIELD_ACTIVITY_CONFIG.UPDATE_TIMES;
  var bestTime = times[0];
  var bestDiff = 99999;

  times.forEach(function(time) {
    var diff = Math.abs(timeToMinutes_(time) - currentMinutes);
    if (diff < bestDiff) {
      bestDiff = diff;
      bestTime = time;
    }
  });

  return bestTime;
}

function getFieldActivityScheduleForDate_(dateKey) {
  var day = getDayOfWeekFromDateKey_(dateKey);
  var isSaturday = day === 6;
  var isSunday = day === 0;
  var endTime = isSaturday ? FIELD_ACTIVITY_CONFIG.SATURDAY_END_TIME : FIELD_ACTIVITY_CONFIG.END_TIME;
  var updateTimes = FIELD_ACTIVITY_CONFIG.UPDATE_TIMES.filter(function(time) {
    return timeToMinutes_(time) < timeToMinutes_(endTime);
  });
  var labels = [FIELD_ACTIVITY_CONFIG.START_TIME].concat(updateTimes).concat([endTime]);

  return {
    startTime: FIELD_ACTIVITY_CONFIG.START_TIME,
    updateTimes: updateTimes,
    endTime: endTime,
    labels: labels,
    isWorkday: !isSunday,
    text: isSaturday
      ? 'Jadwal Sabtu 08:00, 10:00, 12:00, 14:00. Toleransi 30 menit.'
      : 'Jadwal Senin-Jumat 08:00, 10:00, 12:00, 14:00, 16:00. Toleransi 30 menit.'
  };
}

function getDayOfWeekFromDateKey_(dateKey) {
  var safeDate = normalizeSheetDateToYmd_(dateKey) || getNowParts_().tanggal;
  var parts = String(safeDate || '').split('-');
  var date;

  if (parts.length < 3) {
    return 1;
  }

  date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return date.getDay();
}

function getFieldActivityTimeStatus_(targetTime, actualTime) {
  if (!String(targetTime || '').trim()) {
    return 'Tercatat';
  }

  var diff = timeToMinutes_(actualTime) - timeToMinutes_(targetTime);
  var tolerance = FIELD_ACTIVITY_CONFIG.TOLERANCE_MINUTES;

  if (Math.abs(diff) <= tolerance) {
    return 'Tepat Waktu';
  }

  return diff > tolerance ? 'Terlambat' : 'Lebih Awal';
}

function timeToMinutes_(timeValue) {
  var parts = String(timeValue || '').split(':');
  return (Number(parts[0] || 0) * 60) + Number(parts[1] || 0);
}

function getActiveSalesUsers_() {
  return getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).filter(function(user) {
    return normalizeRoleKey_(user.role) === 'sales' && normalizeText_(user.status_aktif) === 'aktif';
  });
}

function getFieldActivityLogsForUserDate_(userId, dateKey) {
  var safeUserId = String(userId || '').trim();
  var safeDate = String(dateKey || '').trim();

  return getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI).filter(function(row) {
    return String(row.user_id || '').trim() === safeUserId &&
      normalizeSheetDateToYmd_(row.tanggal) === safeDate;
  }).sort(function(left, right) {
    return normalizeFieldActivityTime_(left.jam_server).localeCompare(normalizeFieldActivityTime_(right.jam_server));
  });
}

function getPreviousFieldActivityLog_(userId, dateKey) {
  var logs = getFieldActivityLogsForUserDate_(userId, dateKey);
  return logs.length ? logs[logs.length - 1] : null;
}

function buildFieldActivityMovementStatus_(previousLog, latitude, longitude, nowParts) {
  var result = {
    distanceKm: '',
    minutes: '',
    speedKmh: '',
    status: 'Normal'
  };
  var previousLat;
  var previousLng;
  var minutes;
  var distanceKm;
  var speed;

  if (!previousLog) {
    return result;
  }

  previousLat = Number(previousLog.latitude);
  previousLng = Number(previousLog.longitude);

  if (isNaN(previousLat) || isNaN(previousLng)) {
    return result;
  }

  minutes = Math.max(1, timeToMinutes_(nowParts.jam) - timeToMinutes_(normalizeFieldActivityTime_(previousLog.jam_server)));
  distanceKm = calculateDistanceKm_(previousLat, previousLng, latitude, longitude);
  speed = (distanceKm / minutes) * 60;

  result.distanceKm = roundNumber_(distanceKm, 2);
  result.minutes = minutes;
  result.speedKmh = roundNumber_(speed, 1);

  if (speed > FIELD_ACTIVITY_CONFIG.SPEED_SUSPICIOUS_KMH) {
    result.status = 'Mencurigakan';
  } else if (speed > FIELD_ACTIVITY_CONFIG.SPEED_CHECK_KMH) {
    result.status = 'Perlu Cek';
  }

  return result;
}

function calculateDistanceKm_(lat1, lon1, lat2, lon2) {
  var radiusKm = 6371;
  var dLat = degreesToRadians_(lat2 - lat1);
  var dLon = degreesToRadians_(lon2 - lon1);
  var a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(degreesToRadians_(lat1)) * Math.cos(degreesToRadians_(lat2)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return radiusKm * c;
}

function degreesToRadians_(value) {
  return Number(value || 0) * Math.PI / 180;
}

function roundNumber_(value, precision) {
  var power = Math.pow(10, precision || 0);
  return Math.round(Number(value || 0) * power) / power;
}

function buildGoogleMapsLink_(latitude, longitude) {
  return 'https://www.google.com/maps?q=' + latitude + ',' + longitude;
}

function registerFieldActivityDevice_(user, deviceId, userAgent, nowParts) {
  var safeDeviceId = String(deviceId || '').trim() || 'unknown';
  var safeUserId = String(user.user_id || '').trim();
  var key = safeUserId + '|' + safeDeviceId;
  var sheet = getSheetByName_(APP_CONFIG.SHEETS.PERANGKAT_USER);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var rows = getSheetData_(APP_CONFIG.SHEETS.PERANGKAT_USER);
  var existing = rows.find(function(row) {
    return String(row.perangkat_key || '').trim() === key;
  });
  var weekDeviceCount;
  var status;

  if (existing) {
    status = String(existing.status_perangkat || 'Perangkat Utama').trim() || 'Perangkat Utama';
    updateFieldActivityDeviceRow_(sheet, headers, key, {
      terakhir_dipakai: nowIso_(),
      jumlah_pemakaian: Number(existing.jumlah_pemakaian || 0) + 1,
      user_agent: userAgent || existing.user_agent || ''
    });
    return status;
  }

  weekDeviceCount = countFieldActivityDevicesThisWeek_(safeUserId, nowParts.tanggal);
  status = weekDeviceCount >= 2 ? 'Sering Ganti Perangkat' : (weekDeviceCount >= 1 ? 'Perangkat Baru' : 'Perangkat Utama');

  appendRowByHeaders_(APP_CONFIG.SHEETS.PERANGKAT_USER, {
    perangkat_key: key,
    device_id: safeDeviceId,
    user_id: safeUserId,
    nama_user: user.nama_user || '',
    user_agent: userAgent || '',
    pertama_dipakai: nowIso_(),
    terakhir_dipakai: nowIso_(),
    jumlah_pemakaian: 1,
    status_perangkat: status,
    catatan: ''
  });

  return status;
}

function updateFieldActivityDeviceRow_(sheet, headers, key, updates) {
  var keyIndex = headers.indexOf('perangkat_key');
  var lastRow = sheet.getLastRow();
  var values;
  var rowIndex;

  if (keyIndex === -1 || lastRow < 2) {
    return;
  }

  values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][keyIndex] || '').trim() === key) {
      headers.forEach(function(header, columnIndex) {
        if (Object.prototype.hasOwnProperty.call(updates, header)) {
          values[rowIndex][columnIndex] = updates[header];
        }
      });
      sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([values[rowIndex]]);
      return;
    }
  }
}

function countFieldActivityDevicesThisWeek_(userId, dateKey) {
  var weekStart = getMondayDateKey_(dateKey);
  var devices = {};

  getSheetData_(APP_CONFIG.SHEETS.PERANGKAT_USER).forEach(function(row) {
    var firstDate = normalizeSheetDateToYmd_(row.pertama_dipakai);
    if (String(row.user_id || '').trim() === userId && firstDate >= weekStart) {
      devices[String(row.device_id || '').trim()] = true;
    }
  });

  return Object.keys(devices).length;
}

function getMondayDateKey_(dateKey) {
  var parts = String(dateKey || getNowParts_().tanggal).split('-');
  var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var day = date.getDay();
  var diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return Utilities.formatDate(date, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function buildFieldActivityDailyRow_(user, dateKey) {
  var logs = getFieldActivityLogsForUserDate_(user.user_id, dateKey);
  var schedule = getFieldActivityScheduleForDate_(dateKey);
  var slots = buildFieldActivitySlots_(logs, schedule);
  var activityNotes = buildFieldActivitySalesNotes_(logs);
  var totalTelat = 0;
  var totalTidakUpdate = 0;
  var anomaliLokasi = 0;
  var anomaliPerangkat = 0;
  var lastLog = logs.length ? logs[logs.length - 1] : null;
  var statusHarian;

  slots.forEach(function(slot) {
    if (slot.status === 'Terlambat' || slot.status === 'Lebih Awal') {
      totalTelat += 1;
    }
    if (slot.status === 'Tidak Update') {
      totalTidakUpdate += 1;
    }
  });

  logs.forEach(function(log) {
    if (normalizeText_(log.status_lokasi) === 'perlu cek' || normalizeText_(log.status_lokasi) === 'mencurigakan') {
      anomaliLokasi += 1;
    }
    if (normalizeText_(log.status_perangkat) === 'perangkat baru' || normalizeText_(log.status_perangkat) === 'sering ganti perangkat') {
      anomaliPerangkat += 1;
    }
  });

  statusHarian = 'Hijau';
  if (!logs.length) {
    statusHarian = 'Belum Mulai';
  } else if (totalTidakUpdate >= 2 || anomaliLokasi > 0 || anomaliPerangkat > 0) {
    statusHarian = 'Merah';
  } else if (totalTidakUpdate === 1 || totalTelat > 0) {
    statusHarian = 'Kuning';
  }

  return {
    tanggal: dateKey,
    user_id: user.user_id || '',
    nama_user: user.nama_user || '',
    tipe_sales: user.tipe_sales || '',
    channel_sales_default: user.channel_sales_default || '',
    slots: slots,
    jadwal_kerja: schedule.text,
    jam_mulai: getSlotLogTime_(slots, schedule.startTime),
    jam_update_terakhir: lastLog ? normalizeFieldActivityTime_(lastLog.jam_server) : '',
    jam_selesai: getSlotLogTime_(slots, schedule.endTime),
    total_update: logs.length,
    total_catatan_kegiatan_sales: activityNotes.length,
    catatan_kegiatan_sales: activityNotes,
    total_telat: totalTelat,
    total_tidak_update: totalTidakUpdate,
    anomali_lokasi: anomaliLokasi,
    anomali_perangkat: anomaliPerangkat,
    status_harian: statusHarian,
    lokasi_terakhir: lastLog ? String(lastLog.link_maps || '') : '',
    lokasi_terakhir_jam: lastLog ? normalizeFieldActivityTime_(lastLog.jam_server) : '',
    next_action: buildFieldActivityNextAction_(slots)
  };
}

function buildFieldActivitySlots_(logs, schedule) {
  var safeSchedule = schedule || getFieldActivityScheduleForDate_(getNowParts_().tanggal);
  var definitions = [{ target: safeSchedule.startTime, label: safeSchedule.startTime, type: 'Mulai Kerja' }]
    .concat(safeSchedule.updateTimes.map(function(time) {
      return { target: time, label: time, type: 'Update Lokasi' };
    }))
    .concat([{ target: safeSchedule.endTime, label: safeSchedule.endTime, type: 'Selesai Kerja' }]);

  return definitions.map(function(definition) {
    var matching = (logs || []).filter(function(log) {
      return normalizeFieldActivityTargetTime_(log.target_jam) === definition.target;
    }).sort(function(left, right) {
      return getFieldSlotStatusRank_(left.status_waktu) - getFieldSlotStatusRank_(right.status_waktu);
    })[0];

    return {
      target: definition.target,
      label: definition.label,
      type: definition.type,
      status: matching ? getFieldActivityEffectiveTimeStatus_(matching) : 'Tidak Update',
      jam: matching ? normalizeFieldActivityTime_(matching.jam_server) : '',
      lokasi: matching ? String(matching.link_maps || '').trim() : '',
      status_lokasi: matching ? String(matching.status_lokasi || '').trim() : '',
      status_perangkat: matching ? String(matching.status_perangkat || '').trim() : ''
    };
  });
}

function buildFieldActivitySalesNotes_(logs) {
  return (logs || []).filter(function(log) {
    return normalizeText_(log.tipe_aktivitas) === 'catat kegiatan' ||
      String(log.jenis_kegiatan || '').trim() ||
      String(log.hasil_kegiatan || '').trim();
  }).map(function(log) {
    return {
      jam: normalizeFieldActivityTime_(log.jam_server),
      jenis_kegiatan: String(log.jenis_kegiatan || '').trim(),
      nama_customer: String(log.nama_customer || '').trim(),
      hasil_kegiatan: String(log.hasil_kegiatan || log.catatan || '').trim(),
      lokasi: String(log.link_maps || '').trim()
    };
  });
}

function getFieldSlotStatusRank_(status) {
  var key = normalizeText_(status);
  if (key === 'tepat waktu') return 1;
  if (key === 'terlambat' || key === 'lebih awal') return 2;
  return 3;
}

function normalizeFieldActivityTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIMEZONE, 'HH:mm:ss');
  }

  var text = String(value || '').trim();
  var match = text.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);

  if (!match) {
    return text;
  }

  return [
    String(match[1]).padStart(2, '0'),
    match[2],
    match[3] || '00'
  ].join(':');
}

function normalizeFieldActivityTargetTime_(value) {
  var timeText = normalizeFieldActivityTime_(value);
  var match = String(timeText || '').match(/^(\d{2}):(\d{2})/);

  if (!match) {
    return String(timeText || '').trim();
  }

  return match[1] + ':' + match[2];
}

function getSlotLogTime_(slots, targetTime) {
  var slot = (slots || []).find(function(item) {
    return item.target === targetTime;
  });
  return slot && slot.jam ? slot.jam : '';
}

function buildFieldActivityNextAction_(slots) {
  var nextSlot = (slots || []).find(function(slot) {
    return slot.status === 'Tidak Update';
  });

  if (!nextSlot) {
    return {
      label: 'Aktivitas hari ini lengkap',
      type: '',
      target: ''
    };
  }

  return {
    label: nextSlot.type + ' ' + nextSlot.label,
    type: nextSlot.type,
    target: nextSlot.target
  };
}

function buildFieldMonitoringSummary_(rows) {
  var total = rows.length;
  var hijau = 0;
  var kuning = 0;
  var merah = 0;

  rows.forEach(function(row) {
    if (row.status_harian === 'Hijau') hijau += 1;
    if (row.status_harian === 'Kuning') kuning += 1;
    if (row.status_harian === 'Merah' || row.status_harian === 'Belum Mulai') merah += 1;
  });

  return {
    total_sales: total,
    hijau: hijau,
    kuning: kuning,
    merah: merah
  };
}

function buildFieldActivityWeeklyAnomalies_(dateKey) {
  var weekStart = getMondayDateKey_(dateKey);
  var logs = getSheetData_(APP_CONFIG.SHEETS.AKTIVITAS_LOKASI).filter(function(log) {
    var logDate = normalizeSheetDateToYmd_(log.tanggal);
    return logDate >= weekStart && logDate <= dateKey;
  });
  var byUser = {};

  logs.forEach(function(log) {
    var userId = String(log.user_id || '').trim();
    if (!userId) return;
    if (!byUser[userId]) {
      byUser[userId] = {
        user_id: userId,
        nama_user: String(log.nama_user || '').trim(),
        telat: 0,
        anomali_lokasi: 0,
        anomali_perangkat: 0
      };
    }
    if (normalizeText_(getFieldActivityEffectiveTimeStatus_(log)) === 'terlambat') byUser[userId].telat += 1;
    if (normalizeText_(log.status_lokasi) === 'perlu cek' || normalizeText_(log.status_lokasi) === 'mencurigakan') byUser[userId].anomali_lokasi += 1;
    if (normalizeText_(log.status_perangkat) === 'perangkat baru' || normalizeText_(log.status_perangkat) === 'sering ganti perangkat') byUser[userId].anomali_perangkat += 1;
  });

  return Object.keys(byUser).map(function(key) {
    return byUser[key];
  }).filter(function(row) {
    return row.telat || row.anomali_lokasi || row.anomali_perangkat;
  }).sort(function(left, right) {
    return (right.telat + right.anomali_lokasi + right.anomali_perangkat) -
      (left.telat + left.anomali_lokasi + left.anomali_perangkat);
  });
}
