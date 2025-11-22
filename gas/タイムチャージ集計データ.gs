/**** タイムチャージ集計データ.gs — タイムチャージ月次集計専用 ****/

// ★修正: 月次集計スプシ保存フォルダ
const TC_SUMMARY_FOLDER_ID = '1Qjgaphz-Cjf3he4UMgRvYo2N5vrRrLsy';

// ★修正: 月次集計スプレッドシートの列順
const TC_SUMMARY_HEADERS = [
  '担当者','請求種別','得意先コード','得意先名','ケースコード','ケース名','単価','カバー時間',
  '合計作業時間','超過時間','請求金額','請求データ作成FLG','請求データ作成日','請求データ作成者','備考（社内）'
];

// ★修正: 公開ラッパー（UI/トリガーから呼び出し）
function run_TimeChargeMonthlySummary(billingYm) {
  return rebuild_TimeChargeSummaryForMonth(billingYm);
}

// ★修正: 前月分を再構築するラッパー（トリガー用）
function run_TimeChargeMonthlySummaryForPreviousMonth() {
  const ym = tc_getPreviousYearMonthString();
  return rebuild_TimeChargeSummaryForMonth(ym);
}

// ★修正: 前月のYYYY/MM文字列を返すヘルパー
function tc_getPreviousYearMonthString() {
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}`;
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

// ★修正: 毎月1日に前月分の集計を走らせるトリガーを作成
function setup_TimeChargeMonthlyTrigger() {
  ScriptApp.newTrigger('run_TimeChargeMonthlySummaryForPreviousMonth')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
}

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

// end of タイムチャージ集計データ.gs
