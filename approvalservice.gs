function approveOrder(noSo, approverId, catatanApproval) {
  return decideApproval_(noSo, approverId, 'Disetujui', 'Siap Kirim', catatanApproval);
}

function rejectOrder(noSo, approverId, catatanApproval) {
  return decideApproval_(noSo, approverId, 'Ditolak', 'Ditolak', catatanApproval);
}

function testApproveLatestWaitingOrder() {
  var approval = getLatestWaitingApproval_();

  if (!approval) {
    throw new Error('Tidak ada order menunggu persetujuan untuk dites');
  }

  console.log(JSON.stringify(approveOrder(approval.no_so, 'U003', 'Disetujui untuk uji approval')));
}

function testRejectLatestWaitingOrder() {
  var approval = getLatestWaitingApproval_();

  if (!approval) {
    throw new Error('Tidak ada order menunggu persetujuan untuk dites');
  }

  console.log(JSON.stringify(rejectOrder(approval.no_so, 'U003', 'Ditolak untuk uji approval')));
}

function decideApproval_(noSo, approverId, statusApproval, statusOrderBaru, catatanApproval) {
  var approval = findApprovalByNoSo_(noSo);

  if (!approval) {
    throw new Error('Approval order tidak ditemukan untuk no_so: ' + noSo);
  }

  if (normalizeText_(approval.status_approval) !== 'menunggu') {
    throw new Error('Approval sudah diproses sebelumnya untuk no_so: ' + noSo);
  }

  var salesOrder = findSalesOrderByNoSo_(noSo);

  if (!salesOrder) {
    throw new Error('Sales order tidak ditemukan untuk no_so: ' + noSo);
  }

  var now = getNowParts_();
  var note = String(catatanApproval || '').trim();

  if (!note) {
    throw new Error('Catatan approval wajib diisi.');
  }

  var updatedApproval = updateRowByKey_(APP_CONFIG.SHEETS.APPROVAL_ORDER, 'approval_id', approval.approval_id, {
    status_approval: statusApproval,
    diputuskan_oleh: approverId,
    tanggal_keputusan: now.tanggal + ' ' + now.jam,
    catatan_approval: note
  });

  var updatedSalesOrder = updateRowByKey_(APP_CONFIG.SHEETS.SALES_ORDER, 'no_so', noSo, {
    status_order: statusOrderBaru,
    butuh_persetujuan: statusApproval === 'Disetujui' ? 'Tidak' : salesOrder.butuh_persetujuan
  });

  logStatusOrder_(noSo, salesOrder.status_order, statusOrderBaru, approverId, note);

  if (normalizeText_(statusOrderBaru) === 'siap kirim') {
    try {
      var nowForKpi = getNowParts_();
      recordKpiLogForOrderIfEligible_(noSo, approverId, nowForKpi.tanggal + ' ' + nowForKpi.jam);
    } catch (error) {
      console.log('KPI log gagal: ' + (error && error.message ? error.message : error));
    }
  }

  return {
    success: true,
    no_so: noSo,
    status_approval: updatedApproval.status_approval,
    status_order: updatedSalesOrder.status_order
  };
}

function getLatestWaitingApproval_() {
  var approvals = getSheetData_(APP_CONFIG.SHEETS.APPROVAL_ORDER).filter(function(row) {
    return normalizeText_(row.status_approval) === 'menunggu';
  });

  if (!approvals.length) {
    return null;
  }

  return approvals[approvals.length - 1];
}

function findApprovalByNoSo_(noSo) {
  var targetNoSo = String(noSo || '').trim();
  var matches = getSheetData_(APP_CONFIG.SHEETS.APPROVAL_ORDER).filter(function(row) {
    return String(row.no_so || '').trim() === targetNoSo;
  });

  if (!matches.length) {
    return null;
  }

  // Prioritaskan approval yang masih menunggu agar approve/reject tidak salah
  // mengenai riwayat approval lama untuk no_so yang sama.
  var pendingMatches = matches.filter(function(row) {
    return normalizeText_(row.status_approval) === 'menunggu';
  });

  if (pendingMatches.length) {
    return pendingMatches[pendingMatches.length - 1];
  }

  return matches[matches.length - 1];
}

function findSalesOrderByNoSo_(noSo) {
  return getSheetData_(APP_CONFIG.SHEETS.SALES_ORDER).find(function(row) {
    return String(row.no_so).trim() === String(noSo).trim();
  }) || null;
}
