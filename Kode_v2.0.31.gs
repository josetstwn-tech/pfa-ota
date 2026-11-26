function doGet(e) {
  try {
    var mode = (e && e.parameter && e.parameter.mode) ? String(e.parameter.mode).trim().toLowerCase() : '';
    if (mode === 'kartu' && typeof renderPublicKartuGantungHtml_ === 'function') {
      return renderPublicKartuGantungHtml_(e.parameter.kode || e.parameter.nama || '');
    }
    if (mode === 'sj' && typeof renderPublicSuratJalanHtml_ === 'function') {
      return renderPublicSuratJalanHtml_(e.parameter.no || '');
    }
  } catch (errMode) {}
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PLN Field Assistant (PFA)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Konfigurasi default update APK (bisa diubah kapan saja lewat Project Settings -> Script Properties)
var DEFAULT_APK_VERSION = '2.0.31';
var DEFAULT_APK_DOWNLOAD_URL = '';
var DEFAULT_APK_RELEASE_NOTES = 'Pembaruan PFA v2.0.9: Desain Ulang Portal PLN Corporate Pro (Light & PLN Night Dark Mode), Quick Search + Barcode Portal, Analisa AI Pengukuran, & Reliability Engine (Idempotency, Write Lock, Auto-Retry, Auto-Sync).';

function getAppUpdateMetadata_() {
  var props = PropertiesService.getScriptProperties();
  return {
    status: 'success',
    app: 'PLN Field Assistant (PFA)',
    latestVersion: String(props.getProperty('APK_LATEST_VERSION') || DEFAULT_APK_VERSION).trim(),
    minNativeVersion: String(props.getProperty('APK_MIN_NATIVE_VERSION') || '2.0.0').trim(),
    otaBundleUrl: String(props.getProperty('OTA_BUNDLE_URL') || '').trim(),
    betaOtaBundleUrl: String(props.getProperty('BETA_OTA_BUNDLE_URL') || '').trim(),
    downloadUrl: String(props.getProperty('APK_DOWNLOAD_URL') || DEFAULT_APK_DOWNLOAD_URL).trim(),
    releaseNotes: String(props.getProperty('APK_RELEASE_NOTES') || DEFAULT_APK_RELEASE_NOTES).trim(),
    serverTime: new Date().toISOString()
  };
}

function validateServerSessionToken_(token) {
  var val = String(token || '').trim();
  if (!val) return null;
  var cached = CacheService.getScriptCache().get('sess_' + val);
  if (!cached) return null;
  try {
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

function simpanMasukanPetugas_(req) {
  var run = function () {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Masukan_Petugas');
    if (!sheet) {
      sheet = ss.insertSheet('Masukan_Petugas');
      sheet.appendRow(['Timestamp', 'Versi', 'Petugas', 'Username', 'Kategori', 'Pesan']);
    }
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Jakarta';
    var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    var clean = (typeof pfaSanitizeCellValue_ === 'function') ? pfaSanitizeCellValue_ : function (v) { return v; };
    sheet.appendRow([
      stamp,
      clean(String(req.versi || '-')),
      clean(String(req.petugasNama || '-')),
      clean(String(req.username || '-')),
      clean(String(req.kategori || 'Umum')),
      clean(String(req.pesan || '').trim())
    ]);
    return { status: 'success', message: 'Masukan berhasil dicatat.' };
  };
  return (typeof pfaWithWriteLock_ === 'function') ? pfaWithWriteLock_(run) : run();
}

/**
 * REST API Gateway untuk Aplikasi Android (Capacitor) — PFA v2.0.9.
 * Memwajibkan sessionToken valid pada seluruh aksi selain login, ping, check_app_update, dan api_version_check.
 * Web App lama (doGet) tetap berjalan normal untuk Backoffice Admin di PC.
 */
function doPost(e) {
  var responseObj;
  try {
    var rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var payload = JSON.parse(rawBody);
    var action = String((payload && payload.action) || '').trim();
    var publicActions = ['login', 'ping', 'check_app_update', 'api_version_check', 'get_supabase_config'];

    if (publicActions.indexOf(action) === -1) {
      var sessUser = validateServerSessionToken_(payload && payload.sessionToken);
      if (!sessUser) {
        return ContentService
          .createTextOutput(JSON.stringify({
            status: 'error',
            code: 'SESSION_EXPIRED',
            message: 'Sesi login tidak valid atau telah berakhir. Silakan login kembali.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Perpanjang masa aktif token otomatis (sliding TTL 6 jam)
      try {
        CacheService.getScriptCache().put('sess_' + String(payload.sessionToken).trim(), JSON.stringify(sessUser), 21600);
      } catch (eTtl) {}
    }

    if (sessUser) payload._sessUser = sessUser;
    responseObj = apiHandler(payload);
  } catch (err) {
    responseObj = {
      status: 'error',
      message: 'Format permintaan API tidak valid: ' + (err && err.message ? err.message : String(err))
    };
  }
  return ContentService
    .createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function pfaExecuteReliableWrite_(req, writeFn) {
  var reqId = String((req && req.clientRequestId) || '').trim();
  if (reqId && typeof pfaCheckIdempotency_ === 'function') {
    var prev = pfaCheckIdempotency_(reqId);
    if (prev) return prev;
  }
  var exec = function () {
    var res = writeFn();
    if (reqId && typeof pfaSaveIdempotency_ === 'function') {
      pfaSaveIdempotency_(reqId, res);
    }
    return res;
  };
  return (typeof pfaWithWriteLock_ === 'function') ? pfaWithWriteLock_(exec) : exec();
}

function apiHandler(payload) {
  try {
    const req = payload || {};
    const action = String(req.action || '').trim();
    const sanitizeObj = (typeof pfaSanitizeObject_ === 'function') ? pfaSanitizeObject_ : function (o) { return o; };

    switch (action) {
      case 'ping':
      case 'check_app_update':
        return getAppUpdateMetadata_();

      case 'get_portal_summary':
        return (typeof pfaGetPortalSummary_ === 'function')
          ? pfaGetPortalSummary_(req, req._sessUser || validateServerSessionToken_(req.sessionToken))
          : { status: 'error', message: 'Modul Reliability_v208 belum terpasang.' };

      case 'api_version_check':
        return (typeof pfaApiVersionCheck_ === 'function')
          ? pfaApiVersionCheck_()
          : { status: 'success', apiVersion: DEFAULT_APK_VERSION };

      case 'log_client_error':
        return (typeof pfaLogClientError_ === 'function')
          ? pfaLogClientError_(req)
          : { status: 'success' };

      case 'kirim_masukan_pfa':
        return simpanMasukanPetugas_(req);

      case 'login':
        if (typeof pfaCheckLoginRateLimit_ === 'function') {
          var rl = pfaCheckLoginRateLimit_(req.username);
          if (rl.blocked) return { status: 'error', code: 'LOGIN_LOCKED', message: rl.message };
        }
        var loginRes = handleLogin(req.username, req.password);
        if (typeof pfaRecordLoginResult_ === 'function') {
          pfaRecordLoginResult_(req.username, Boolean(loginRes && loginRes.status === 'success'));
        }
        return loginRes;

      case 'get_master_kerusakan':
        return getMasterKerusakan();

      case 'cari_pelanggan':
        return (typeof cariPelangganHybrid_ === 'function') ? cariPelangganHybrid_(req.keyword) : cariPelanggan(req.keyword);

      case 'cek_duplikat_idpel':
        return cekDuplikatIdpel(req.idpel);

      case 'simpan_gangguan':
        return pfaExecuteReliableWrite_(req, function () {
          return simpanGangguan(sanitizeObj(req.data || {}), req.fotos || {}, Boolean(req.allowDuplicate));
        });

      case 'simpan_pengukuran':
        if (!req.allowDuplicate && req.data && req.data.id_pelanggan) {
          var dupUkur = pfaCheckDuplikatSheetByIdpel_((typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.REKAP_PENGUKURAN) || 'Rekap_Pengukuran', req.data.id_pelanggan, 1, 0, 21, 13);
          if (dupUkur && dupUkur.exists) {
            return { status: 'duplicate', message: 'IDPEL ' + req.data.id_pelanggan + ' sudah pernah diukur sebelumnya.', duplicate: dupUkur };
          }
        }
        return pfaExecuteReliableWrite_(req, function () {
          return simpanPengukuran(sanitizeObj(req.data || {}), req.fotos || {});
        });

      case 'ocr_foto_meter_ct':
        return ocrFotoMeterCt(req.imageBase64);

      case 'simpan_p2tl':
        if (!req.allowDuplicate && req.data && req.data.id_pelanggan && String(req.data.kategori_temuan || '') !== 'Non Pelanggan') {
          var dupP2tl = pfaCheckDuplikatSheetByIdpel_((typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.REKAP_P2TL) || 'Rekap_P2TL', req.data.id_pelanggan, 5, 0, 19, 3);
          if (dupP2tl && dupP2tl.exists) {
            return { status: 'duplicate', message: 'IDPEL ' + req.data.id_pelanggan + ' sudah pernah tercatat di P2TL sebelumnya.', duplicate: dupP2tl };
          }
        }
        return pfaExecuteReliableWrite_(req, function () {
          return simpanP2tl(sanitizeObj(req.data || {}), req.fotos || {});
        });

      case 'tanya_ai_lapangan':
        return tanyaAiLapangan(req.prompt, req.history || [], req.imageBase64 || '', req.useSearch !== false, req.aiMode || 'auto');

      case 'get_master_material':
        return getMasterMaterial();

      case 'simpan_master_material':
        return pfaExecuteReliableWrite_(req, function () {
          return simpanMasterMaterial(req.sessionToken || req.adminRole, sanitizeObj(req.item || {}));
        });

      case 'ajukan_material':
        return pfaExecuteReliableWrite_(req, function () {
          return ajukanTransaksiMaterial(req.sessionToken || req.userRole, sanitizeObj(req.data || {}), req.ttdBase64 || '');
        });

      case 'get_antrean_material':
        return getAntreanMaterial(req.sessionToken || req.userRole);

      case 'verifikasi_security_material':
        return pfaExecuteReliableWrite_(req, function () {
          return verifikasiSecurityMaterial(req.sessionToken || req.userRole, req.noSurat, req.keputusan, req.catatan, req.ttdBase64 || '', req.fotoBase64 || '');
        });

      case 'approve_admin_material':
        return pfaExecuteReliableWrite_(req, function () {
          return approveAdminGudangMaterial(req.sessionToken || req.userRole, req.noSurat, req.keputusan, req.catatan, req.ttdBase64 || '');
        });

      case 'hapus_master_material':
        return pfaExecuteReliableWrite_(req, function () {
          return hapusMasterMaterial(req.sessionToken || req.adminRole, req.kode);
        });

      case 'ubah_urutan_material':
        return pfaExecuteReliableWrite_(req, function () {
          return ubahUrutanMaterial(req.sessionToken || req.adminRole, req.kode, req.arah);
        });

      case 'sinkron_master_material_default':
        return pfaExecuteReliableWrite_(req, function () {
          return sinkronMasterMaterialDefault(req.sessionToken || req.adminRole);
        });

      case 'get_kartu_gantung':
        return getKartuGantungMaterial(req.kode || req.nama || '');

      case 'generate_kartu_gantung_pdf':
        return generateKartuGantungPdf(req.sessionToken || req.userRole, req.kode || req.nama || '');

      case 'get_surat_jalan_pdf':
        return getSuratJalanMaterialPdf(req.sessionToken || req.userRole, req.noSurat || '');

      case 'generate_qr_rak_pdf':
        return generateQrRakPdf(req.sessionToken || req.userRole, req.filterGroup || 'ALL', req.singleKode || '');

      case 'get_module_permissions':
        return { status: 'success', permissions: getModulePermissionsMap() };

      case 'admin_get_module_permissions':
        return adminGetModulePermissions(req.sessionToken || req.adminRole);

      case 'admin_save_module_permissions':
        return adminSaveModulePermissions(req.sessionToken || req.adminRole, req.permissions || {});

      case 'admin_get_monitoring_material':
        return adminGetMonitoringMaterial(req.sessionToken || req.adminRole, req.filterBulan);

      case 'cek_kct':
        return (typeof cekKctHybrid_ === 'function') ? cekKctHybrid_(req.keyword, req.unit) : cekKct(req.keyword, req.unit);

      case 'get_histori_petugas':
        return getHistoriPetugas(req.petugasNama, req.username);

      case 'admin_tambah_user':
        return adminTambahUser(req.sessionToken || req.adminRole, sanitizeObj(req.newUser || {}));

      case 'admin_get_users':
        return adminGetUsers(req.sessionToken || req.adminRole);

      case 'admin_update_user':
        return adminUpdateUser(req.sessionToken || req.adminRole, req.username, req.status, req.newPassword, req.role);

      case 'admin_get_rekap':
        return adminGetRekap(req.sessionToken || req.adminRole, req.filterTanggal);

      case 'admin_get_rekap_pengukuran':
        return adminGetRekapPengukuran(req.sessionToken || req.adminRole, req.filterTanggal);

      case 'admin_get_kb_list':
        return adminGetKbList(req.sessionToken || req.adminRole);

      case 'admin_save_kb_manual':
        return adminSaveKbManual(req.sessionToken || req.adminRole, sanitizeObj(req.item || {}));

      case 'admin_update_kb_status':
        return adminUpdateKbStatus(req.sessionToken || req.adminRole, req.idKb, req.status, req.judul, req.ringkasan, req.kataKunci);

      case 'admin_add_kb_topik':
        return adminAddKbTopik(req.sessionToken || req.adminRole, sanitizeObj(req.topik || {}));

      case 'admin_delete_kb_topik':
        return adminDeleteKbTopik(req.sessionToken || req.adminRole, req.idTopik);

      case 'admin_run_kb_riset':
        return adminRunKbRisetNow(req.sessionToken || req.adminRole, req.kategori, req.topikOrUrl, req.jenis);

      case 'admin_setup_kb_trigger':
        return adminSetupKbTrigger(req.sessionToken || req.adminRole, Boolean(req.enabled));

      case 'admin_import_kb_document':
        return adminImportKbDocument(req.sessionToken || req.adminRole, sanitizeObj(req.doc || {}));

      case 'admin_get_system_health':
        return adminGetSystemHealth(req.sessionToken || req.adminRole);

      case 'admin_test_gemini_health':
        return adminTestGeminiHealth(req.sessionToken || req.adminRole);

      case 'admin_clear_system_logs':
        return adminClearSystemLogs(req.sessionToken || req.adminRole);

      case 'admin_test_drive_write':
        return adminTestDriveWrite(req.sessionToken || req.adminRole);

      case 'admin_repair_missing_sheets':
        return adminRepairMissingSheets(req.sessionToken || req.adminRole);

      case 'admin_prune_old_logs':
        return adminPruneOldSystemLogs(req.sessionToken || req.adminRole, req.days || 30);

      case 'admin_scan_dus_ai':
        return (typeof adminScanDusAi === 'function')
          ? adminScanDusAi(req._sessUser, req.sessionToken || req.adminRole, req.imageBase64 || '')
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'admin_buat_drop_posko':
        return (typeof adminBuatDropPosko === 'function')
          ? adminBuatDropPosko(req._sessUser, req.sessionToken || req.adminRole, req.data || {})
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'admin_batal_drop_posko':
        return (typeof adminBatalDropPosko === 'function')
          ? adminBatalDropPosko(req._sessUser, req.sessionToken || req.adminRole, req.noBa)
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'posko_get_kiriman':
        return (typeof poskoGetKiriman === 'function')
          ? poskoGetKiriman(req._sessUser || validateServerSessionToken_(req.sessionToken), req)
          : { status: 'success', posko: '', pendingCount: 0, items: [] };

      case 'posko_konfirmasi_terima':
        return (typeof poskoKonfirmasiTerima === 'function')
          ? poskoKonfirmasiTerima(req._sessUser || validateServerSessionToken_(req.sessionToken), req)
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'posko_tolak_drop':
        return (typeof poskoTolakDrop === 'function')
          ? poskoTolakDrop(req._sessUser || validateServerSessionToken_(req.sessionToken), req)
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'admin_cari_meter_atau_dus':
        return (typeof adminCariMeterAtauDus === 'function')
          ? adminCariMeterAtauDus(req._sessUser, req.sessionToken || req.adminRole, req.keyword || '')
          : { status: 'error', message: 'PoskoKwh_v214.gs belum diperbarui ke v2.0.24.' };

      case 'get_supabase_config':
        return (typeof getSupabasePublicConfig_ === 'function')
          ? getSupabasePublicConfig_()
          : { status: 'success', enabled: false };

      case 'admin_sync_sheets_to_supabase':
        return (typeof adminSyncSheetsToSupabase === 'function')
          ? adminSyncSheetsToSupabase(req._sessUser, req.sessionToken || req.adminRole, req.tableName)
          : { status: 'error', message: 'Supabase_v220.gs belum diperbarui ke v2.0.24.' };

      case 'admin_upsert_supabase_batch':
        return (typeof adminUpsertSupabaseBatch === 'function')
          ? adminUpsertSupabaseBatch(req._sessUser, req.sessionToken || req.adminRole, req.tableName, req.rows || [])
          : { status: 'error', message: 'File Supabase_v220.gs belum ditambahkan ke Apps Script.' };

      case 'admin_get_saldo_kwh_posko':
        return (typeof adminGetSaldoKwhPosko === 'function')
          ? adminGetSaldoKwhPosko(req._sessUser, req.sessionToken || req.adminRole)
          : { status: 'error', message: 'File PoskoKwh_v214.gs belum ditambahkan ke Apps Script.' };

      case 'admin_set_posko_petugas':
        return (typeof adminSetPoskoPetugas === 'function')
          ? adminSetPoskoPetugas(req._sessUser, req.sessionToken || req.adminRole, req.username, req.namaPetugas, req.rolePetugas, req.posko)
          : { status: 'error', message: 'File PoskoKwh_v214.gs belum ditambahkan ke Apps Script.' };

      case 'admin_simpan_mutasi_kwh_posko':
        return (typeof adminSimpanMutasiKwhPosko === 'function')
          ? adminSimpanMutasiKwhPosko(req._sessUser, req.sessionToken || req.adminRole, req.data || {})
          : { status: 'error', message: 'File PoskoKwh_v214.gs belum ditambahkan ke Apps Script.' };

      case 'admin_backup_spreadsheet':
        return adminBackupSpreadsheetSnapshot(req.sessionToken || req.adminRole);

      default:
        return {
          status: 'error',
          message: 'Aksi "' + action + '" tidak dikenali.'
        };
    }
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    try {
      if (typeof logSystemError_ === 'function') {
        const reqObj = payload || {};
        logSystemError_(reqObj.action || 'unknown', reqObj.username || reqObj.petugasNama || reqObj.adminRole || '-', msg);
      }
    } catch (e2) {}
    return {
      status: 'error',
      message: msg
    };
  }
}


function pfaCheckDuplikatSheetByIdpel_(sheetName, idPel, idColIdx, tglColIdx, petColIdx, ketColIdx) {
  try {
    var cleanId = String(idPel || '').replace(/[^0-9A-Za-z]/g, '').trim().toUpperCase();
    if (!cleanId || cleanId === 'NONPELANGGAN') return null;
    var ss = getSpreadsheet();
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return null;
    var maxCol = Math.max(idColIdx, tglColIdx, petColIdx, ketColIdx) + 1;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(Math.max(sh.getLastColumn(), maxCol), 30)).getDisplayValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var rowId = String(rows[i][idColIdx] || '').replace(/[^0-9A-Za-z]/g, '').trim().toUpperCase();
      if (rowId === cleanId) {
        return {
          exists: true,
          latest: {
            idpel: rows[i][idColIdx] || idPel,
            tanggal: rows[i][tglColIdx] || '-',
            petugas: rows[i][petColIdx] || '-',
            keterangan: rows[i][ketColIdx] || '-'
          }
        };
      }
    }
  } catch (e) {}
  return null;
}
