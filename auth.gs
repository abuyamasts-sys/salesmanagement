function getCurrentUserProfile(userId) {
  var user = getCurrentUserRecord_(userId);
  var tipeSales;
  var isFreelance;
  var channelSalesDefault;

  if (!user) {
    return {
      email: '',
      user_id: '',
      nama_user: '',
      role: '',
      role_key: '',
      status_aktif: '',
      tipe_sales: '',
      kode_sales: '',
      channel_sales_default: '',
      bank_nama: '',
      bank_no_rekening: '',
      bank_nama_pemilik: '',
      bank_account_complete: false,
      is_freelance: false,
      authorized: false
    };
  }

  tipeSales = String(user.tipe_sales || '').trim() || 'Internal';
  isFreelance = normalizeText_(tipeSales) === 'freelance';
  channelSalesDefault = String(user.channel_sales_default || '').trim() || (isFreelance ? 'SLF' : 'SLS');

  return {
    email: user.email || '',
    user_id: user.user_id || '',
    nama_user: user.nama_user || '',
    role: user.role || '',
    role_key: normalizeRoleKey_(user.role),
    status_aktif: user.status_aktif || '',
    tipe_sales: tipeSales,
    kode_sales: String(user.kode_sales || '').trim(),
    channel_sales_default: channelSalesDefault,
    bank_nama: String(user.bank_nama || '').trim(),
    bank_no_rekening: String(user.bank_no_rekening || '').trim(),
    bank_nama_pemilik: String(user.bank_nama_pemilik || '').trim(),
    bank_account_complete: !!(
      String(user.bank_nama || '').trim() &&
      String(user.bank_no_rekening || '').trim() &&
      String(user.bank_nama_pemilik || '').trim()
    ),
    is_freelance: isFreelance,
    authorized: true
  };
}

function getUserProfileByUserId(userId) {
  return getCurrentUserProfile(userId);
}

function ensureMasterUserPasswords_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.MASTER_USER, APP_CONFIG.HEADERS.MASTER_USER);
  ensureDefaultControllerUser_();

  getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).forEach(function(user) {
    var userId = String(user.user_id || '').trim();
    var password = String(user.password || '');

    if (!userId || password) {
      return;
    }

    updateRowByKey_(APP_CONFIG.SHEETS.MASTER_USER, 'user_id', userId, {
      password: userId
    });
  });
}

function loginWithPassword(userId, password) {
  var normalizedUserId = String(userId || '').trim().toUpperCase();
  var rawPassword = String(password || '');
  var user = getCurrentUserRecord_(normalizedUserId);
  var storedPassword;

  if (!normalizedUserId) {
    throw new Error('User ID wajib diisi.');
  }

  if (!rawPassword) {
    throw new Error('Password wajib diisi.');
  }

  if (!user) {
    throw new Error('User ID tidak ditemukan atau status user tidak aktif.');
  }

  storedPassword = String(user.password || '');
  if (!storedPassword) {
    throw new Error('Password untuk user ini belum diatur di MASTER_USER.');
  }

  if (storedPassword !== rawPassword) {
    throw new Error('Password salah.');
  }

  return getCurrentUserProfile(normalizedUserId);
}

function requireAuthorizedUser_(userId) {
  var user = getCurrentUserRecord_(userId);

  if (!user) {
    throw new Error('Akses ditolak. User ID belum terdaftar atau status user tidak aktif di MASTER_USER.');
  }

  return user;
}

function requireCurrentUserRole_(allowedRoles, userId) {
  var user = requireAuthorizedUser_(userId);

  if (!userHasAllowedRole_(user, allowedRoles)) {
    throw new Error('Akses ditolak. Role Anda (' + (user.role || '-') + ') tidak memiliki otorisasi untuk proses ini.');
  }

  return user;
}

function getCurrentUserRecord_(userId) {
  var normalizedUserId = normalizeText_(userId);

  if (!normalizedUserId) {
    return null;
  }

  return getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).find(function(user) {
    return normalizeText_(user.user_id) === normalizedUserId &&
      normalizeText_(user.status_aktif) === 'aktif';
  }) || null;
}

function normalizeRoleKey_(role) {
  var value = normalizeText_(role);

  if (value === 'csadmin' || value === 'cs/admin') {
    return 'cs_admin';
  }

  return value.replace(/[^a-z0-9]+/g, '_');
}

function userHasAllowedRole_(user, allowedRoles) {
  var roleKeys = (allowedRoles || []).map(function(role) {
    return normalizeRoleKey_(role);
  });
  var userRoleKey = normalizeRoleKey_(user && user.role);

  return roleKeys.indexOf(userRoleKey) !== -1;
}

function ensureDefaultControllerUser_() {
  var controllerUserId = 'CTR1';
  var existingController = getSheetData_(APP_CONFIG.SHEETS.MASTER_USER).find(function(user) {
    return normalizeText_(user.user_id) === normalizeText_(controllerUserId);
  });

  if (existingController) {
    return;
  }

  appendRowByHeaders_(APP_CONFIG.SHEETS.MASTER_USER, {
    user_id: controllerUserId,
    nama_user: 'Controller',
    role: 'Controller',
    no_hp: '',
    email: '',
    password: controllerUserId,
    status_aktif: 'Aktif',
    tipe_sales: '',
    kode_sales: '',
    channel_sales_default: '',
    bank_nama: '',
    bank_no_rekening: '',
    bank_nama_pemilik: '',
    aktif_komisi: '',
    catatan_user: 'Auto-seeded default controller user'
  });
}
