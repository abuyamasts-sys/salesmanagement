function getAdminBillingData_(currentUser, options) {
  var safeUser = currentUser || {};
  var syncResult = syncTagihanFromCustomerMaster_(safeUser);
  var tagihan = syncResult.tagihan || [];

  tagihan.sort(function(left, right) {
    var leftStatus = normalizeText_(left.status_tagihan);
    var rightStatus = normalizeText_(right.status_tagihan);
    if (leftStatus !== rightStatus) {
      if (leftStatus === 'belum lunas') return -1;
      if (rightStatus === 'belum lunas') return 1;
    }

    var leftJt = normalizeSheetDateToYmd_(left.jt_terdekat || '');
    var rightJt = normalizeSheetDateToYmd_(right.jt_terdekat || '');
    if (leftJt !== rightJt) {
      return String(leftJt || '9999-12-31').localeCompare(String(rightJt || '9999-12-31'), 'id-ID');
    }

    return String(left.nama_customer || '').localeCompare(String(right.nama_customer || ''), 'id-ID');
  });

  var summary = tagihan.reduce(function(result, row) {
    var status = normalizeText_(row.status_tagihan);
    var total = Number(row.total_tagihan || 0);

    if (status === 'belum lunas') {
      result.belum_lunas += 1;
      result.total_belum_lunas += total;
    } else if (status === 'sebagian') {
      result.sebagian += 1;
      result.total_sebagian += total;
    } else if (status === 'lunas') {
      result.lunas += 1;
      result.total_lunas += total;
    }

    result.total_semua += total;
    result.total_row += 1;
    return result;
  }, {
    total_row: 0,
    belum_lunas: 0,
    sebagian: 0,
    lunas: 0,
    total_semua: 0,
    total_belum_lunas: 0,
    total_sebagian: 0,
    total_lunas: 0
  });

  return {
    tagihan: tagihan,
    billingSummary: summary,
    sync: {
      created: syncResult.created || 0,
      updated: syncResult.updated || 0,
      autoClosed: syncResult.autoClosed || 0
    }
  };
}

function syncTagihanFromCustomerMaster_(currentUser) {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.TAGIHAN, APP_CONFIG.HEADERS.TAGIHAN);

  var now = getNowParts_();
  var customers = getSheetData_(APP_CONFIG.SHEETS.MASTER_CUSTOMER);
  var existing = getSheetData_(APP_CONFIG.SHEETS.TAGIHAN);
  var openByCustomer = {};

  existing.forEach(function(row) {
    var status = normalizeText_(row.status_tagihan);
    var sumberTagihan = normalizeText_(row.sumber_tagihan);
    if (sumberTagihan === 'sales_order') {
      return;
    }

    if (status !== 'belum lunas' && status !== 'sebagian') {
      return;
    }

    var kode = String(row.kode_customer || '').trim();
    if (!kode) return;
    openByCustomer[kode] = row;
  });

  var created = 0;
  var updated = 0;
  var autoClosed = 0;

  customers.forEach(function(customer) {
    var kodeCustomer = String(customer.kode_customer || '').trim();
    if (!kodeCustomer) {
      return;
    }

    var statusPembayaran = normalizeText_(customer.status_pembayaran);
    var statusCustomer = normalizeText_(customer.status_customer);
    var totalTunggakan = Number(customer.total_tunggakan || 0);
    var jumlahNotaOverdue = Number(customer.jumlah_nota_overdue || 0);

    var isMenunggak = statusPembayaran === 'menunggak' ||
      statusCustomer === 'menunggak' ||
      statusCustomer === 'ditahan' ||
      statusPembayaran === 'ditahan' ||
      totalTunggakan > 0 ||
      jumlahNotaOverdue > 0;

    var openTagihan = openByCustomer[kodeCustomer] || null;

    if (!isMenunggak) {
      if (openTagihan && normalizeText_(customer.status_pembayaran) === 'lancar' && totalTunggakan <= 0) {
        updateRowByKey_(APP_CONFIG.SHEETS.TAGIHAN, 'tagihan_id', openTagihan.tagihan_id, {
          status_tagihan: 'Lunas',
          tanggal_lunas: now.tanggal + ' ' + now.jam,
          dilunasi_oleh: (currentUser && currentUser.user_id) ? currentUser.user_id : '',
          catatan_pelunasan: 'Auto ditutup: status customer sudah lancar'
        });
        autoClosed += 1;
      }
      return;
    }

    var snapshot = {
      kode_customer: kodeCustomer,
      nama_customer: String(customer.nama_customer || '').trim(),
      status_customer: String(customer.status_customer || '').trim(),
      status_pembayaran: String(customer.status_pembayaran || '').trim(),
      total_tagihan: totalTunggakan,
      jumlah_nota_overdue: jumlahNotaOverdue,
      jt_terdekat: normalizeSheetDateToYmd_(customer.tanggal_jatuh_tempo_terdekat || ''),
      catatan_piutang: String(customer.catatan_piutang || customer.catatan || '').trim()
    };

    if (!openTagihan) {
      var totalAwal = totalTunggakan;
      appendRowByHeaders_(APP_CONFIG.SHEETS.TAGIHAN, {
        tagihan_id: generateDocNumber_('TGH'),
        sumber_tagihan: 'MASTER_CUSTOMER',
        no_so: '',
        no_surat_jalan: '',
        kode_customer: snapshot.kode_customer,
        nama_customer: snapshot.nama_customer,
        sales_id: '',
        sales_nama: '',
        term_pembayaran: '',
        tanggal_order: '',
        tanggal_kirim: '',
        tanggal_jatuh_tempo: '',
        status_customer: snapshot.status_customer,
        status_pembayaran: snapshot.status_pembayaran,
        total_tagihan: snapshot.total_tagihan,
        total_awal_tagihan: totalAwal,
        total_bayar: 0,
        sisa_tagihan: totalAwal,
        jumlah_nota_overdue: snapshot.jumlah_nota_overdue,
        jt_terdekat: snapshot.jt_terdekat,
        catatan_piutang: snapshot.catatan_piutang,
        status_tagihan: 'Belum Lunas',
        tanggal_dibuat: now.tanggal + ' ' + now.jam,
        dibuat_oleh: (currentUser && currentUser.user_id) ? currentUser.user_id : '',
        tanggal_bayar_terakhir: '',
        dibayar_oleh_terakhir: '',
        catatan_bayar_terakhir: '',
        tanggal_lunas: '',
        dilunasi_oleh: '',
        catatan_pelunasan: ''
      });
      created += 1;
      return;
    }

    var totalBayar = Number(openTagihan.total_bayar || 0);
    var totalAwalTagihan = Number(openTagihan.total_awal_tagihan || openTagihan.total_tagihan || 0);
    if (totalAwalTagihan <= 0) {
      totalAwalTagihan = Number(snapshot.total_tagihan || 0);
    }

    var sisaTagihan = Math.max(0, Number(snapshot.total_tagihan || 0) - totalBayar);
    var statusTagihan = sisaTagihan <= 0 ? 'Lunas' : (totalBayar > 0 ? 'Sebagian' : 'Belum Lunas');

    var shouldUpdate = false;
    ['nama_customer', 'status_customer', 'status_pembayaran', 'total_tagihan', 'jumlah_nota_overdue', 'jt_terdekat', 'catatan_piutang'].forEach(function(key) {
      var currentVal = openTagihan[key];
      var nextVal = snapshot[key];
      if (String(currentVal || '') !== String(nextVal || '')) {
        shouldUpdate = true;
      }
    });

    if (!shouldUpdate) {
      return;
    }

    updateRowByKey_(APP_CONFIG.SHEETS.TAGIHAN, 'tagihan_id', openTagihan.tagihan_id, Object.keys(snapshot).reduce(function(result, key) {
      result[key] = snapshot[key];
      return result;
    }, {
      sumber_tagihan: String(openTagihan.sumber_tagihan || '').trim() || 'MASTER_CUSTOMER',
      no_so: openTagihan.no_so || '',
      no_surat_jalan: openTagihan.no_surat_jalan || '',
      sales_id: openTagihan.sales_id || '',
      sales_nama: openTagihan.sales_nama || '',
      term_pembayaran: openTagihan.term_pembayaran || '',
      tanggal_order: openTagihan.tanggal_order || '',
      tanggal_kirim: openTagihan.tanggal_kirim || '',
      tanggal_jatuh_tempo: openTagihan.tanggal_jatuh_tempo || '',
      total_awal_tagihan: totalAwalTagihan,
      total_bayar: totalBayar,
      sisa_tagihan: sisaTagihan,
      status_tagihan: statusTagihan
    }));
    updated += 1;
  });

  return {
    tagihan: getSheetData_(APP_CONFIG.SHEETS.TAGIHAN),
    created: created,
    updated: updated,
    autoClosed: autoClosed
  };
}

function createOrSyncTagihanFromSalesOrderTempo_(salesOrder, suratJalan, totals, currentUser) {
  var order = salesOrder || {};
  var noSo = String(order.no_so || '').trim();
  var termPembayaran = String(order.term_pembayaran || '').trim();
  var totalTagihan = Number((totals && totals.total) || order.total_final || order.total || 0);
  var existing;
  var existingPaid;
  var sisaTagihan;
  var statusTagihan;
  var now;
  var tanggalJatuhTempo;
  var row;

  if (!noSo || normalizeText_(termPembayaran).indexOf('tempo') === -1) {
    return null;
  }

  if (!(totalTagihan > 0)) {
    return null;
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.TAGIHAN, APP_CONFIG.HEADERS.TAGIHAN);

  existing = getSheetData_(APP_CONFIG.SHEETS.TAGIHAN).find(function(item) {
    return normalizeText_(item.sumber_tagihan) === 'sales_order' &&
      String(item.no_so || '').trim() === noSo;
  }) || null;

  if (existing && normalizeText_(existing.status_tagihan) === 'lunas') {
    return existing;
  }

  now = getNowParts_();
  tanggalJatuhTempo = calculateTagihanDueDateFromVerification_(now.tanggal, termPembayaran) ||
    normalizeSheetDateToYmd_(order.tanggal_jatuh_tempo || '');
  existingPaid = existing ? Number(existing.total_bayar || 0) : 0;
  sisaTagihan = Math.max(0, totalTagihan - existingPaid);
  statusTagihan = sisaTagihan <= 0 ? 'Lunas' : (existingPaid > 0 ? 'Sebagian' : 'Belum Lunas');
  row = {
    sumber_tagihan: 'SALES_ORDER',
    no_so: noSo,
    no_surat_jalan: suratJalan && suratJalan.no_surat_jalan ? suratJalan.no_surat_jalan : '',
    kode_customer: order.customer_id || '',
    nama_customer: order.nama_customer_input || '',
    sales_id: order.sales_id || '',
    sales_nama: order.sales_nama || '',
    term_pembayaran: termPembayaran,
    tanggal_order: order.tanggal_order || '',
    tanggal_kirim: suratJalan && suratJalan.tanggal_kirim ? suratJalan.tanggal_kirim : (order.tanggal_kirim_rencana || ''),
    tanggal_jatuh_tempo: tanggalJatuhTempo,
    status_customer: '',
    status_pembayaran: termPembayaran,
    total_tagihan: totalTagihan,
    total_awal_tagihan: existing ? Number(existing.total_awal_tagihan || totalTagihan) : totalTagihan,
    total_bayar: existingPaid,
    sisa_tagihan: sisaTagihan,
    jumlah_nota_overdue: 0,
    jt_terdekat: tanggalJatuhTempo,
    catatan_piutang: 'Tagihan tempo dari SO ' + noSo,
    status_tagihan: statusTagihan
  };

  if (existing) {
    return updateRowByKey_(APP_CONFIG.SHEETS.TAGIHAN, 'tagihan_id', existing.tagihan_id, row);
  }

  row.tagihan_id = generateDocNumber_('TGH');
  row.tanggal_dibuat = now.tanggal + ' ' + now.jam;
  row.dibuat_oleh = currentUser && currentUser.user_id ? currentUser.user_id : '';
  row.tanggal_bayar_terakhir = '';
  row.dibayar_oleh_terakhir = '';
  row.catatan_bayar_terakhir = '';
  row.tanggal_lunas = '';
  row.dilunasi_oleh = '';
  row.catatan_pelunasan = '';
  appendRowByHeaders_(APP_CONFIG.SHEETS.TAGIHAN, row);
  return row;
}

function calculateTagihanDueDateFromVerification_(tanggalVerifikasi, termPembayaran) {
  var verificationDate = normalizeSheetDateToYmd_(tanggalVerifikasi);
  var tempoDays = getTempoDaysFromTermPembayaran_(termPembayaran);

  if (!verificationDate || tempoDays < 0) {
    return '';
  }

  return addDaysToYmdDate_(verificationDate, tempoDays);
}

function getTempoDaysFromTermPembayaran_(termPembayaran) {
  var text = String(termPembayaran || '').trim();
  var match = text.match(/(\d+)/);

  if (normalizeText_(text).indexOf('tempo') === -1) {
    return -1;
  }

  return match ? Number(match[1] || 0) : 0;
}

function addDaysToYmdDate_(ymdDate, daysToAdd) {
  var match = String(ymdDate || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  var date;

  if (!match) {
    return '';
  }

  date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  date.setDate(date.getDate() + Number(daysToAdd || 0));
  return Utilities.formatDate(date, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function markTagihanLunas_(currentUser, payload) {
  var safePayload = payload || {};
  var tagihanId = String(safePayload.tagihan_id || '').trim();
  var kodeCustomer = String(safePayload.kode_customer || '').trim();
  var catatanPelunasan = String(safePayload.catatan_pelunasan || '').trim();

  if (!tagihanId && !kodeCustomer) {
    throw new Error('Tagihan tidak valid (butuh tagihan_id atau kode_customer).');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.TAGIHAN, APP_CONFIG.HEADERS.TAGIHAN);
  var now = getNowParts_();
  var tagihanRows = getSheetData_(APP_CONFIG.SHEETS.TAGIHAN);
  var target = null;

  if (tagihanId) {
    target = tagihanRows.find(function(row) {
      return String(row.tagihan_id || '').trim() === tagihanId;
    }) || null;
  } else {
    target = tagihanRows.find(function(row) {
      return String(row.kode_customer || '').trim() === kodeCustomer && normalizeText_(row.status_tagihan) === 'belum lunas';
    }) || null;
  }

  if (!target) {
    throw new Error('Tagihan tidak ditemukan.');
  }

  if (normalizeText_(target.status_tagihan) === 'lunas') {
    return {
      success: true,
      message: 'Tagihan sudah lunas.',
      tagihan: target
    };
  }

  return applyTagihanPayment_(currentUser, {
    tagihan_id: target.tagihan_id,
    nominal_bayar: Math.max(0, Number(target.sisa_tagihan || target.total_tagihan || 0)),
    catatan_bayar: catatanPelunasan,
    mode: 'LUNAS'
  });
}

function applyTagihanPayment_(currentUser, payload) {
  var safePayload = payload || {};
  var tagihanId = String(safePayload.tagihan_id || '').trim();
  var nominalBayar = Number(safePayload.nominal_bayar || 0);
  var catatanBayar = String(safePayload.catatan_bayar || '').trim();
  var mode = String(safePayload.mode || '').trim().toUpperCase();

  if (!tagihanId) {
    throw new Error('Tagihan tidak valid.');
  }

  if (!(nominalBayar > 0) && mode !== 'LUNAS') {
    throw new Error('Nominal bayar harus lebih dari 0.');
  }

  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.TAGIHAN, APP_CONFIG.HEADERS.TAGIHAN);
  var now = getNowParts_();
  var tagihanRows = getSheetData_(APP_CONFIG.SHEETS.TAGIHAN);
  var target = tagihanRows.find(function(row) {
    return String(row.tagihan_id || '').trim() === tagihanId;
  }) || null;

  if (!target) {
    throw new Error('Tagihan tidak ditemukan.');
  }

  var currentStatus = normalizeText_(target.status_tagihan);
  if (currentStatus === 'lunas') {
    return {
      success: true,
      message: 'Tagihan sudah lunas.',
      tagihan: target
    };
  }

  var currentUserId = currentUser && currentUser.user_id ? currentUser.user_id : '';
  var isSalesOrderTagihan = normalizeText_(target.sumber_tagihan) === 'sales_order';
  var totalAwal = Number(target.total_awal_tagihan || target.total_tagihan || 0);
  var totalBayarLama = Number(target.total_bayar || 0);
  var totalTagihan = Number(target.total_tagihan || 0);
  var sisaLama = Number(target.sisa_tagihan || Math.max(0, totalTagihan - totalBayarLama));
  var bayarEfektif = mode === 'LUNAS' ? sisaLama : Math.min(nominalBayar, sisaLama);
  var totalBayarBaru = totalBayarLama + bayarEfektif;
  var sisaBaru = Math.max(0, totalTagihan - totalBayarBaru);
  var statusBaru = sisaBaru <= 0 ? 'Lunas' : (totalBayarBaru > 0 ? 'Sebagian' : 'Belum Lunas');

  var catatanPelunasan = String(target.catatan_pelunasan || '').trim();
  var paymentNoteParts = [
    'Bayar',
    now.tanggal,
    bayarEfektif > 0 ? ('Rp ' + formatNumberServer_(bayarEfektif)) : ''
  ].filter(Boolean);
  var paymentNote = paymentNoteParts.join(' ');
  if (catatanBayar) {
    paymentNote += '. ' + catatanBayar;
  }
  var nextCatatanPelunasan = catatanPelunasan ? (catatanPelunasan + ' | ' + paymentNote) : paymentNote;

  var updates = {
    total_awal_tagihan: totalAwal,
    total_bayar: totalBayarBaru,
    sisa_tagihan: sisaBaru,
    status_tagihan: statusBaru,
    tanggal_bayar_terakhir: now.tanggal + ' ' + now.jam,
    dibayar_oleh_terakhir: currentUserId,
    catatan_bayar_terakhir: catatanBayar,
    catatan_pelunasan: nextCatatanPelunasan
  };

  if (statusBaru === 'Lunas') {
    updates.tanggal_lunas = now.tanggal + ' ' + now.jam;
    updates.dilunasi_oleh = currentUserId;
  }

  var updatedTagihan = updateRowByKey_(APP_CONFIG.SHEETS.TAGIHAN, 'tagihan_id', tagihanId, updates);

  var existingCustomer = isSalesOrderTagihan ? null : findCustomerByCode_(target.kode_customer);
  if (existingCustomer) {
    var statusCustomer = normalizeText_(existingCustomer.status_customer);
    var nextCustomerStatus = existingCustomer.status_customer;
    var nextStatusPembayaran = 'Menunggak';
    var nextTotalTunggakan = sisaBaru;
    var nextNotaOverdue = Number(existingCustomer.jumlah_nota_overdue || 0);
    var shouldClearJt = false;

    if (sisaBaru <= 0) {
      nextStatusPembayaran = 'Lancar';
      nextTotalTunggakan = 0;
      nextNotaOverdue = 0;
      shouldClearJt = true;

      if (statusCustomer === 'menunggak') {
        nextCustomerStatus = 'Aktif';
      }
    } else {
      // Jika customer ditahan, buka jika sudah bayar minimal 50% dari total awal.
      var isHold = statusCustomer === 'ditahan' || normalizeText_(existingCustomer.status_pembayaran) === 'ditahan';
      if (isHold && totalAwal > 0 && totalBayarBaru / totalAwal >= 0.5) {
        nextCustomerStatus = 'Menunggak';
      }
    }

    updateRowByKey_(APP_CONFIG.SHEETS.MASTER_CUSTOMER, 'kode_customer', target.kode_customer, {
      status_customer: nextCustomerStatus,
      status_pembayaran: nextStatusPembayaran,
      total_tunggakan: nextTotalTunggakan,
      jumlah_nota_overdue: nextNotaOverdue,
      tanggal_jatuh_tempo_terdekat: shouldClearJt ? '' : existingCustomer.tanggal_jatuh_tempo_terdekat,
      catatan_piutang: appendCustomerCatatanPiutang_(existingCustomer.catatan_piutang, now, currentUserId, paymentNote)
    });
  }

  if (statusBaru === 'Lunas') {
    if (isSalesOrderTagihan && String(target.no_so || '').trim()) {
      markSlfCommissionReadyByNoSo_(target.no_so);
    } else {
      markSlfCommissionReadyByCustomer_(target.kode_customer);
    }
  }

  return {
    success: true,
    message: statusBaru === 'Lunas' ? 'Tagihan berhasil dilunasi.' : 'Pembayaran berhasil disimpan.',
    tagihan: updatedTagihan
  };
}

function recordTagihanPayment_(currentUser, payload) {
  var safePayload = payload || {};
  var tagihanId = String(safePayload.tagihan_id || '').trim();
  var nominalBayar = Number(safePayload.nominal_bayar || 0);
  var catatanBayar = String(safePayload.catatan_bayar || '').trim();

  return applyTagihanPayment_(currentUser, {
    tagihan_id: tagihanId,
    nominal_bayar: nominalBayar,
    catatan_bayar: catatanBayar,
    mode: 'CICIL'
  });
}

function appendCustomerCatatanPiutang_(catatanLama, nowParts, userId, catatanPelunasan) {
  var base = String(catatanLama || '').trim();
  var suffixParts = [
    (nowParts && nowParts.tanggal) ? nowParts.tanggal : '',
    userId ? ('oleh ' + userId) : '',
    '—'
  ].filter(Boolean);

  var suffix = suffixParts.join(' ') + ' ' + String(catatanPelunasan || '').trim();
  suffix = suffix.trim();

  if (!suffix) {
    return base;
  }

  if (!base) {
    return suffix;
  }

  return base + ' | ' + suffix;
}

function markSlfCommissionReadyByCustomer_(kodeCustomer) {
  var customerCode = String(kodeCustomer || '').trim();

  if (!customerCode) {
    return;
  }

  getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).filter(function(order) {
    return String(order.customer_id || '').trim() === customerCode &&
      normalizeText_(order.channel_sales) === 'slf' &&
      normalizeText_(order.status_order) === 'selesai' &&
      normalizeText_(order.status_komisi) === 'menunggu pembayaran';
  }).forEach(function(order) {
    updateSalesOrderCommissionStatus_(order.no_so, 'Siap Cair', {
      catatan: 'Tagihan lunas, komisi siap cair'
    });
  });
}

function markSlfCommissionReadyByNoSo_(noSo) {
  var noSoKey = String(noSo || '').trim();
  var order;

  if (!noSoKey) {
    return;
  }

  order = findSalesOrderByNoSo_(noSoKey);
  if (!order ||
    normalizeText_(order.channel_sales) !== 'slf' ||
    normalizeText_(order.status_order) !== 'selesai' ||
    normalizeText_(order.status_komisi) !== 'menunggu pembayaran') {
    return;
  }

  updateSalesOrderCommissionStatus_(noSoKey, 'Siap Cair', {
    catatan: 'Tagihan SO lunas, komisi siap cair'
  });
}
