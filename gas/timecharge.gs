/**** TimeCharge.gs — 作業報告書CSV（タイムチャージ） ****/
/* ▼ 設定（IDはあなたの環境に合わせて差し替え） */
const TC_DB_SPREADSHEET_ID = '18kdtg4E6lHgR40tbxicd-CIFhqVWKznUbxRLNh-MaFc'; // タイムチャージDB（スプシ）
const TC_OUTPUT_FOLDER_ID  = '132tucfV_REeFoRtX74aHJ8-Yd2PAyaZQ';            // 出力フォルダ
const TC_SUMMARY_FOLDER_ID = '1Qjgaphz-Cjf3he4UMgRvYo2N5vrRrLsy';             // ★修正: 月次集計スプシ保存フォルダ

// 入力シート名（空なら候補から自動で探す）
const TC_SHEET_OVERRIDE = ''; // 例: 'タイムチャージ'
const TC_SHEET_CANDIDATES = [
  'タイムチャージ','作業報告書','作業報告書CSV','2.タイムチャージ',
  '作業報告書CSV（入力）','TimeCharge','worklog','作業報告'
];

// プレビュー/CSVの安全弁
const TC_PREVIEW_LIMIT = 100;   // UI表示は上位100件
const TC_HARD_LIMIT    = 50000; // これ以上は切る（念のため）


/* ▼ 出力ヘッダ（作業報告書CSVファイルレイアウト） */
const TC_HEADERS = [
  '担当者','請求種別','得意先コード','得意先','ケースコード','ケース名','単価','発生日',
  '開始時間','終了時間','作業時間','金額（自動計算）','明細備考（内容）',
  '請求データ作成FLG','請求データ作成日','請求データ作成者','備考（社内）'
];

const TC_SUMMARY_HEADERS = [ // ★修正: 月次集計スプレッドシートの列順
  '担当者','請求種別','得意先コード','得意先名','ケースコード','ケース名','単価','カバー時間',
  '合計作業時間','超過時間','請求金額','請求データ作成FLG','請求データ作成日','請求データ作成者','備考（社内）'
];

/* ▼ 入力→出力のマッピング（“論理名”→シート上の候補ヘッダ群） */
const TC_ALIASES = {
  '日付':        ['日付（年月日）','日付(年月日)','発生日','日付','作業日'],
  '担当者':      ['担当者','担当','担当名','ユーザー','ユーザ'],
  '得意先コード':['得意先コード','顧客コード','取引先コード','クライアントコード','客先コード','取引先CD'],
  '得意先':      ['得意先','得意先名','顧客名','取引先','クライアント','会社名','客先','顧客'], // ★修正: 実シートの「得意先名」に対応

  // 以下は“あれば使う”列（必須ではない）
  '請求種別':    ['請求種別','種別'],
  'ケースコード':['ケースコード','案件コード','ケースID'],
  'ケース名':    ['ケース名（案件名）','案件名','ケース名','案件'],
  'カバー時間':  ['カバー時間','カバー分','カバー(分)'], // ★修正: 月次集計用にカバー時間を参照
  '単価/m':      ['単価/m','単価/分','分単価','単価（分）','単価(分)','単価'],
  '開始時間':    ['開始時間','開始','From','開始時刻'],
  '終了時間':    ['終了時間','終了','To','終了時刻'],
  '作業時間':    ['作業時間','作業分','作業時間（分）','作業時間(分)','時間（数量）','数量','時間'],
  '金額':        ['金額（自動計算）','金額','金額(自動計算)'],
  '明細備考':    ['明細備考（内容）','明細備考','内容','備考'],
  'FLG':         ['請求データ作成FLG','作成FLG','請求FLG'],
  '作成日':      ['請求データ作成日','作成日'],
  '作成者':      ['請求データ作成者','作成者','担当'],
  '社内備考':    ['備考（社内）','社内備考','備考(社内)']
};

// 必須（論理名）
const TC_REQUIRED = ['日付','担当者','得意先コード','得意先'];


/* ============ 公開API（UIから呼ばれる） ============ */

// ヘルスチェック（ヘッダ一覧と不足論理名を返す）
function ping_TimeCharge() {
  const t0 = new Date();
  const sh = tc_resolveSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(h=>String(h||''));
  const idxMap  = tc_buildHeaderIndex_(headers);
  const lacks   = TC_REQUIRED.filter(k => typeof idxMap[k] !== 'number');

  const tz = Session.getScriptTimeZone();
  const at = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
  tc_log_('INFO','TC ping',{ lastRow:lr, lastCol:lc, heads:headers, lacks, ms:(new Date())-t0 });
  return { ok: true, lastRow:lr, lastCol:lc, heads:headers, lacks, at, ms:(new Date())-t0 };
}

// プレビュー（対象年月 & FLG空）
function preview_TimeCharge(params) {
  const t0 = new Date();
  try {
    const ymRaw = params && (params.billingYm || params.closingDate || '');
    const ym = tc_parseYearMonth_(ymRaw); // ★修正: UIから渡された年月を厳密にパース
    if (!ym) throw new Error('請求対象年月が未指定/不正（YYYY/MM または YYYYMM）');

    const detail = tc_buildRows_(ym, { hardLimit: TC_HARD_LIMIT });
    const tz = Session.getScriptTimeZone();
    const at = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');

    const summary = tc_buildSummaryMessage_('プレビュー', detail);
    tc_log_('INFO','TC preview completed', Object.assign({ ms:(new Date())-t0 }, detail));

    if (!detail.rows.length) {
      return { ok:false, message:summary, generatedAt: at };
    }
    const body = detail.rows.slice(0, TC_PREVIEW_LIMIT);
    return {
      ok:true,
      rows: detail.picked,
      generatedAt: at,
      preview: [TC_HEADERS].concat(body),
      message: summary
    };
  } catch (e) {
    tc_log_('ERROR','TC preview failed',{ message:e.message, stack:e.stack });
    return { ok:false, message:e.message };
  }
}

// CSV作成
function createCsv_TimeCharge(params) {
  const t0 = new Date();
  try {
    if (!TC_OUTPUT_FOLDER_ID) throw new Error('TC_OUTPUT_FOLDER_ID が未設定です。');
    const ymRaw = params && (params.billingYm || params.closingDate || '');
    const ym = tc_parseYearMonth_(ymRaw); // ★修正: UIから渡された年月を厳密にパース
    if (!ym) throw new Error('請求対象年月が未指定/不正（YYYY/MM または YYYYMM）');

    const detail = tc_buildRows_(ym, { hardLimit: TC_HARD_LIMIT });
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const at  = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm:ss');

    if (!detail.rows.length) {
      tc_log_('INFO','TC CSV skipped(0)', Object.assign({ ms:(new Date())-t0 }, detail));
      return { ok:false, message: tc_buildSummaryMessage_('抽出', detail) };
    }

    const csv = tc_toCsv_([TC_HEADERS].concat(detail.rows));
    const filename = 'TimeCharge_' + Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss') + '.csv';
    const file = DriveApp.getFolderById(TC_OUTPUT_FOLDER_ID).createFile(filename, csv, MimeType.CSV);

    tc_log_('INFO','TC CSV created', Object.assign({ fileId:file.getId(), ms:(new Date())-t0 }, detail));

    return { ok:true, url:file.getUrl(), fileId:file.getId(), filename,
      rows:detail.picked, completedAt: at,
      message: tc_buildSummaryMessage_('CSV作成完了', detail)
    };
  } catch (e) {
    tc_log_('ERROR','TC CSV failed',{ message:e.message, stack:e.stack });
    return { ok:false, message:e.message };
  }
}

// ★修正: 対象年月でタイムチャージ月次集計スプシを再構築
function run_TimeChargeMonthlySummary(billingYm) {
  return rebuild_TimeChargeSummaryForMonth(billingYm);
}

// ★修正: 月次集計本体（指定年月でサマリースプシを再生成）
function rebuild_TimeChargeSummaryForMonth(billingYm) {
  const ym = tc_parseYearMonth_(billingYm);
  if (!ym) throw new Error('請求対象年月が未指定/不正（YYYY/MM または YYYYMM）');

  const sh = tc_resolveSheet_();
  const values = sh.getDataRange().getDisplayValues();
  const headers = values.length ? values[0].map(h => String(h || '').trim()) : [];
  const idxMap  = tc_buildHeaderIndex_(headers);
  const idxOpt  = tc_buildOptionalIndex_(headers);

  const missing = [];
  const req = (label, idx)=>{ if (typeof idx !== 'number') missing.push(label); };
  req('日付', idxMap['日付']);
  req('担当者', idxMap['担当者']);
  req('得意先コード', idxMap['得意先コード']);
  req('得意先', idxMap['得意先']);
  req('作業時間', idxOpt.作業時間);
  req('単価', idxOpt.単価m);
  req('カバー時間', idxOpt.カバー時間);
  const totalRows = Math.max(values.length - 1, 0);

  if (missing.length) {
    const msg = `[TC] 月次集計: 必須列不足（${missing.join(', ')}）`;
    tc_log_('WARN','TC summary missing headers',{ missing, headers });
    return { ok:false, message: msg, missingRequired: missing };
  }

  const groups = new Map();
  let scanned=0, picked=0, skippedBadDate=0, skippedYmMismatch=0;

  for (let r=1; r<values.length; r++){
    scanned++;
    const row = values[r];
    const ymd = tc_extractYearMonth_(row[idxMap['日付']]);
    if (!ymd) { skippedBadDate++; continue; }
    if (ymd.year !== ym.year || ymd.month !== ym.month) { skippedYmMismatch++; continue; }

    const person = tc_vBy_(row, idxMap, '担当者');
    const billType = idxOpt.請求種別 >=0 ? row[idxOpt.請求種別] : '';
    const custCode = tc_vBy_(row, idxMap, '得意先コード');
    const custName = tc_vBy_(row, idxMap, '得意先');
    const caseCode = idxOpt.ケースコード >=0 ? row[idxOpt.ケースコード] : '';
    const caseName = idxOpt.ケース名 >=0 ? row[idxOpt.ケース名] : '';
    const unit     = tc_toNumber_(idxOpt.単価m >=0 ? row[idxOpt.単価m] : 0);
    const cover    = tc_toNumber_(idxOpt.カバー時間 >=0 ? row[idxOpt.カバー時間] : 0);
    const workMin  = tc_toNumber_(idxOpt.作業時間 >=0 ? row[idxOpt.作業時間] : 0);

    const key = [person,billType,custCode,custName,caseCode,caseName,unit,cover].map(v=>String(v||'')).join('\u0001');
    if (!groups.has(key)) {
      groups.set(key, { person, billType, custCode, custName, caseCode, caseName, unit, cover, work:0 });
    }
    const g = groups.get(key);
    g.work += workMin;
    picked++;
  }

  const rows = [];
  groups.forEach(g=>{
    const over = Math.max(g.work - g.cover, 0);
    const price = over * g.unit;
    rows.push([
      g.person,
      g.billType,
      g.custCode,
      g.custName,
      g.caseCode,
      g.caseName,
      g.unit,
      g.cover,
      g.work,
      over,
      price,
      '', '', '', ''
    ]);
  });

  rows.sort((a,b)=>{
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    if (a[2] !== b[2]) return String(a[2]).localeCompare(String(b[2]), 'ja');
    if (a[4] !== b[4]) return String(a[4]).localeCompare(String(b[4]), 'ja');
    return String(a[3]).localeCompare(String(b[3]), 'ja');
  });

  const filename = 'TimeCharge_Summary_' + tc_fmtYearMonth_(ym.year, ym.month).replace('/', '-');
  const target = tc_openOrCreateSummarySheet_(filename);
  const data = [TC_SUMMARY_HEADERS].concat(rows);
  target.sheet.clearContents();
  target.sheet.getRange(1,1,data.length,TC_SUMMARY_HEADERS.length).setValues(data);

  const detail = {
    year: ym.year,
    month: ym.month,
    totalRows,
    scanned,
    picked,
    skippedBadDate,
    skippedYmMismatch,
    fileId: target.file.getId(),
    filename
  };
  tc_log_('INFO','TC monthly summary rebuilt', detail);

  const ymText = tc_fmtYearMonth_(ym.year, ym.month);
  const message = `[TC] 月次集計: ${rows.length}件（対象年月: ${ymText} / 総行数: ${totalRows} / 参照行: ${scanned} / 発生日不正: ${skippedBadDate} / 月一致せず: ${skippedYmMismatch}）`;
  return {
    ok: true,
    fileId: target.file.getId(),
    fileUrl: target.file.getUrl(),
    filename,
    rows: rows.length,
    message
  };
}


/* ============ 本体（抽出） ============ */
function tc_buildRows_(targetYm, opts){
  const hardLimit  = tc_numOrDefault_(opts && opts.hardLimit, TC_HARD_LIMIT);

  const sh = tc_resolveSheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      rows:[], year:targetYm.year, month:targetYm.month,
      scanned:0, totalRows:0, picked:0,
      skippedFlag:0, skippedBadDate:0, skippedYmMismatch:0,
      missingRequired: []
    };
  }

  const headers = values[0].map(h => String(h || '').trim());
  const idxMap  = tc_buildHeaderIndex_(headers);
  const missingRequired = TC_REQUIRED.filter(k => typeof idxMap[k] !== 'number');
  if (missingRequired.length) {
    tc_log_('WARN','TC build aborted: missing required columns',{ missingRequired, headers }); // ★修正: 必須列不足をログ出力
    return {
      rows:[], year:targetYm.year, month:targetYm.month,
      scanned:0, totalRows:values.length-1, picked:0,
      skippedFlag:0, skippedBadDate:0, skippedYmMismatch:0,
      missingRequired
    };
  }

  const idxOptional = tc_buildOptionalIndex_(headers);

  let scanned=0, picked=0, skippedFlag=0, skippedBadDate=0, skippedYmMismatch=0;
  const out = [];
  const totalRows = values.length - 1;

  for (let r=1; r<values.length; r++){
    if (out.length >= hardLimit) break;
    scanned++;

    const row = values[r];
    const dateCell = tc_vBy_(row, idxMap, '日付');
    const ymd = tc_extractYearMonth_(dateCell); // ★修正: 発生日から年月を抽出
    if (!ymd) { skippedBadDate++; continue; }
    if (ymd.year !== targetYm.year || ymd.month !== targetYm.month) { skippedYmMismatch++; continue; }

    if (idxOptional.FLG >= 0) {
      const flg = String(row[idxOptional.FLG] || '').trim();
      if (flg) { skippedFlag++; continue; }
    }

    const vOpt = (i)=> i>=0 ? (row[i] ?? '') : '';
    const tcRow = [
      tc_vBy_(row, idxMap, '担当者'),
      vOpt(idxOptional.請求種別),
      tc_vBy_(row, idxMap, '得意先コード'),
      tc_vBy_(row, idxMap, '得意先'),
      vOpt(idxOptional.ケースコード),
      vOpt(idxOptional.ケース名),
      vOpt(idxOptional.単価m),
      tc_vBy_(row, idxMap, '日付'),
      vOpt(idxOptional.開始時間),
      vOpt(idxOptional.終了時間),
      vOpt(idxOptional.作業時間),
      vOpt(idxOptional.金額),
      vOpt(idxOptional.明細備考),
      vOpt(idxOptional.FLG),
      vOpt(idxOptional.作成日),
      vOpt(idxOptional.作成者),
      vOpt(idxOptional.社内備考)
    ];
    out.push(tcRow);
    picked++;
  }

  const col = (h)=> TC_HEADERS.indexOf(h);
  out.sort((a,b)=>{
    const aY = (a[col('発生日')] || '').replace(/\D/g,'');
    const bY = (b[col('発生日')] || '').replace(/\D/g,'');
    if (aY !== bY) return aY < bY ? -1 : 1;
    const aC = a[col('得意先コード')] || '';
    const bC = b[col('得意先コード')] || '';
    if (aC !== bC) return aC < bC ? -1 : 1;
    const aS = a[col('開始時間')] || '';
    const bS = b[col('開始時間')] || '';
    return aS < bS ? -1 : (aS > bS ? 1 : 0);
  });

  tc_log_('DEBUG','TC build stats',{ scanned, picked, skippedFlag, skippedBadDate, skippedYmMismatch, hardLimit, targetYm });

  return {
    rows: out,
    year: targetYm.year,
    month: targetYm.month,
    scanned,
    totalRows,
    picked,
    skippedFlag,
    skippedBadDate,
    skippedYmMismatch,
    missingRequired: []
  };
}

function tc_buildOptionalIndex_(headers){
  const idxByName = (nameList)=> {
    for (const name of nameList) {
      const norm = tc_normHeader_(name);
      const pos  = headers.findIndex(h => tc_normHeader_(h) === norm);
      if (pos >= 0) return pos;
    }
    return -1;
  };
  return {
    請求種別      : idxByName(TC_ALIASES['請求種別']    || []),
    ケースコード  : idxByName(TC_ALIASES['ケースコード']|| []),
    ケース名      : idxByName(TC_ALIASES['ケース名']    || []),
    カバー時間    : idxByName(TC_ALIASES['カバー時間']  || []), // ★修正: 月次集計用
    単価m         : idxByName(TC_ALIASES['単価/m']      || []),
    開始時間      : idxByName(TC_ALIASES['開始時間']    || []),
    終了時間      : idxByName(TC_ALIASES['終了時間']    || []),
    作業時間      : idxByName(TC_ALIASES['作業時間']    || []),
    金額          : idxByName(TC_ALIASES['金額']        || []),
    明細備考      : idxByName(TC_ALIASES['明細備考']    || []),
    FLG           : idxByName(TC_ALIASES['FLG']         || []),
    作成日        : idxByName(TC_ALIASES['作成日']      || []),
    作成者        : idxByName(TC_ALIASES['作成者']      || []),
    社内備考      : idxByName(TC_ALIASES['社内備考']    || []),
  };
}


/* ============ ヘッダ解決ユーティリティ ============ */
function tc_normHeader_(s){
  return String(s||'')
    .trim()
    .replace(/\s+/g,'')
    .replace(/[()]/g, m => m==='(' ? '（' : '）');
}

function tc_buildHeaderIndex_(headers){
  const idxByNorm = {};
  headers.forEach((h,i)=> idxByNorm[tc_normHeader_(h)] = i);

  const map = {};
  for (const logic of Object.keys(TC_ALIASES)){
    const cands = (TC_ALIASES[logic] || []).map(tc_normHeader_);
    for (const cand of cands){
      if (cand in idxByNorm){ map[logic] = idxByNorm[cand]; break; }
    }
  }
  return map; // {論理名: index}
}

function tc_hasAllRequired_(idxMap){
  return TC_REQUIRED.every(k => typeof idxMap[k] === 'number');
}

function tc_vBy_(row, idxMap, logicKey){
  const i = idxMap[logicKey];
  return (typeof i === 'number') ? (row[i] ?? '') : '';
}


/* ============ シート解決 ============ */
function tc_resolveSheet_(){
  if (!TC_DB_SPREADSHEET_ID) throw new Error('TC_DB_SPREADSHEET_ID が未設定です。');
  const ss = SpreadsheetApp.openById(TC_DB_SPREADSHEET_ID);

  if (TC_SHEET_OVERRIDE) {
    const sh = ss.getSheetByName(TC_SHEET_OVERRIDE);
    if (sh) return sh;
    const list = ss.getSheets().map(s=>s.getName());
    throw new Error('指定シートが見つかりません: ' + TC_SHEET_OVERRIDE + ' / 存在: ' + JSON.stringify(list));
  }
  for (const name of TC_SHEET_CANDIDATES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  const list = ss.getSheets().map(s=>s.getName());
  throw new Error('タイムチャージ入力シートが見つかりません（候補: ' + TC_SHEET_CANDIDATES.join(', ') + '）/ 存在: ' + JSON.stringify(list));
}


/* ============ 日付/共通ユーティリティ ============ */
function tc_parseYearMonth_(value){
  if (value == null || value === '') return null;
  if (value instanceof Date) {
    return { year: value.getFullYear(), month: value.getMonth() + 1 };
  }
  const raw = String(value).trim();
  if (!raw) return null;

  const normalized = raw
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[年月]/g, '/')
    .replace(/日/g, '')
    .replace(/[.]/g, '/')
    .replace(/-/g, '/')
    .trim();

  const full = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (full) {
    const year = Number(full[1]);
    const month = Number(full[2]);
    if (month >= 1 && month <= 12) return { year, month };
    return null;
  }

  const ymSlash = normalized.match(/^(\d{4})\/(\d{1,2})$/);
  if (ymSlash) {
    const year = Number(ymSlash[1]);
    const month = Number(ymSlash[2]);
    if (month >= 1 && month <= 12) return { year, month };
    return null;
  }

  const digits = normalized.replace(/\D/g,'');
  if (digits.length >= 6) {
    const year = Number(digits.slice(0,4));
    const month = Number(digits.slice(4,6));
    if (month >=1 && month <=12) return { year, month };
  }

  return null;
}

function tc_extractYearMonth_(value){
  if (value == null || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    const d = value;
    if (isNaN(d.getTime())) return null;
    return { year: d.getFullYear(), month: d.getMonth() + 1 };
  }

  const raw = String(value).trim();
  if (!raw) return null;

  const normalized = raw
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[年月]/g, '/')
    .replace(/日/g, '')
    .replace(/[.\-]/g, '/')
    .trim();

  const full = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (full) {
    const year = Number(full[1]);
    const month = Number(full[2]);
    if (month >= 1 && month <= 12) return { year, month };
    return null;
  }

  const ym = normalized.match(/^(\d{4})\/(\d{1,2})$/);
  if (ym) {
    const year = Number(ym[1]);
    const month = Number(ym[2]);
    if (month >= 1 && month <= 12) return { year, month };
    return null;
  }

  const digits = normalized.replace(/\D/g, '');
  if (digits.length >= 6) {
    const year = Number(digits.slice(0,4));
    const month = Number(digits.slice(4,6));
    if (month >= 1 && month <= 12) return { year, month };
  }

  return null;
}

function tc_fmtYearMonth_(year, month){
  if (!year || !month) return '';
  return year + '/' + String(month).padStart(2, '0');
}

function tc_numOrDefault_(value, fallback){
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

function tc_toNumber_(value){ // ★修正: 数値化（カンマ・空文字を許容）
  const s = String(value == null ? '' : value).replace(/,/g,'').trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function tc_toCsv_(rows2d){
  return rows2d.map(r=> r.map(x=>{
    const s = (x==null)?'':String(x);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  }).join(',')).join('\n');
}

function tc_buildSummaryMessage_(label, detail){
  const ymText = tc_fmtYearMonth_(detail.year, detail.month) || '-';
  const total = typeof detail.totalRows === 'number' ? detail.totalRows : detail.scanned;
  const picked = typeof detail.picked === 'number' ? detail.picked : (detail.rows ? detail.rows.length : 0);
  const mismatch = detail.skippedYmMismatch || 0;
  const lacks = (detail.missingRequired && detail.missingRequired.length)
    ? ` / 必須列不足: ${detail.missingRequired.join(', ')}`
    : '';
  return `[TC] ${label}: ${picked}件（対象年月: ${ymText} / 総行数: ${total} / 採用: ${picked} / FLGありスキップ: ${detail.skippedFlag || 0} / 発生日不正: ${detail.skippedBadDate || 0} / 月一致せず: ${mismatch}${lacks}）`;
}

function tc_log_(level, message, data){
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  try { Logger.log('%s [%s] %s %s', ts, level, message, data ? JSON.stringify(data) : ''); } catch(_){}}

// ★修正: 月次集計スプシを作成/再利用して最初のシートを返す
function tc_openOrCreateSummarySheet_(filename){
  if (!TC_SUMMARY_FOLDER_ID) throw new Error('TC_SUMMARY_FOLDER_ID が未設定です。');
  const folder = DriveApp.getFolderById(TC_SUMMARY_FOLDER_ID);
  let file = null;
  const it = folder.getFilesByName(filename);
  if (it.hasNext()) {
    file = it.next();
  } else {
    const ssNew = SpreadsheetApp.create(filename);
    file = DriveApp.getFileById(ssNew.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch(_){ /* ignore */ }
  }
  const ss = SpreadsheetApp.openById(file.getId());
  let sheet = ss.getSheets()[0];
  if (!sheet) sheet = ss.insertSheet('Summary');
  sheet.setName('Summary');
  return { sheet, file };
}

/* ============ デバッグ用（任意） ============ */
function dev_ping_TC(){ Logger.log(JSON.stringify(ping_TimeCharge())); }
function dev_preview_TC(){ Logger.log(JSON.stringify(preview_TimeCharge({ billingYm:'2025/07' }))); }
function dev_csv_TC(){ Logger.log(JSON.stringify(createCsv_TimeCharge({ billingYm:'2025/07' }))); }
