/**** Customers (請求書CSV) — 発行日=作成日 / 支払期限=発行月末（プレビュー堅牢化 & 8桁日付対応） ****/
/* ▼ 設定 */
const CUST_DB_SPREADSHEET_ID = '1aZQtCRBtHV18RZsUNQTTds0EPscp0uZNBCvzODV99G8'; // 得意先マスタ専用スプシ
const CUST_SHEET_CANDIDATES  = ['1.得意先マスタ','得意先マスタ'];
const CUST_OUTPUT_FOLDER_ID  = '132tucfV_REeFoRtX74aHJ8-Yd2PAyaZQ';           // 出力フォルダ
const CUST_PREVIEW_LIMIT     = 100;

/* ▼ UI（Webアプリ） */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('ui') // ← 'ui.html' を用意
    .setTitle('アンビシャス 請求ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ▼ 請求書CSVのファイルレイアウト（列順） */
const INVOICE_HEADERS = [
  '得意先コード','得意先名（正式名称）','案件名',
  '請求月','支払期限','請求日','発行日','請求締日',
  '法人区分','契約パターン','契約種別','超過請求','請求元名称','源泉徴収',
  'インボイス番号','振込口座情報','振込先口座1','振込先口座2','振込先口座3',
  '請求先宛名','請求項目（概要）','顧問ケースNo','顧問料/月','数量',
  'カバー時間/分','タイムチャージ','タイムチャージ/分',
  '請求書肩書き','請求書記載の担当者','相手担当者',
  '請求先郵便番号','請求先住所１','請求先住所２',
  '請求備考','備考（社内）','取引開始','取引終了'
];

/* ▼ 列エイリアス（表記ゆれ吸収） */
const CUST_ALIASES = {
  '請求月': ['請求月','請求月(区分)','請求月（区分）'],
  '取引開始': ['取引開始','取引開始日'],
  '取引終了': ['取引終了','取引終了日'],
  '得意先コード': ['得意先コード','顧客コード','取引先コード'],
  '得意先名（正式名称）': ['得意先名（正式名称）','取引先名','顧客名','得意先名'],
};

/* ================== 公開API（UIから呼ばれる） ================== */
// ヘルスチェック（簡易）
function healthcheck(){
  try{
    const sh = resolveCustomerSheet_();
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    return { ok:true, message:`OK ${lr}行 x ${lc}列` };
  }catch(e){
    return { ok:false, message:e.message };
  }
}

// プレビュー：請求締日で抽出し、発行日/支払期限を自動セット
function preview_CustomersByClosing(params){
  try{
    const closing = parseDateStrict_((params && params.closingDate)||'');
    if (!closing) throw new Error('請求締日が未指定/不正（YYYY/MM/DD または YYYYMMDD）');

    const result = buildCustomersRowsSafe_(closing);
    const at = nowStamp_();

    if (!result.ok) {
      return { ok:false, message:`[請求書] プレビュー失敗: ${result.reason}（${at}）` };
    }
    const rows = result.rows || [];
    if (!rows.length){
      return { ok:false, message:`[請求書] 抽出 0件（${at}）` };
    }
    const preview = [INVOICE_HEADERS].concat(rows.slice(0, CUST_PREVIEW_LIMIT));
    return { ok:true, rows: rows.length, generatedAt: at,
             preview, message:`[請求書] プレビュー: ${rows.length}件（${at}）` };
  }catch(e){
    logC_('ERROR','preview_CustomersByClosing failed',{ msg:e.message, stack:e.stack });
    return { ok:false, message:e.message };
  }
}

// CSV作成：請求締日で抽出し、発行日/支払期限を自動セット
function createCsv_CustomersByClosing(params){
  try{
    const closing = parseDateStrict_((params && params.closingDate)||'');
    if (!closing) throw new Error('請求締日が未指定/不正（YYYY/MM/DD または YYYYMMDD）');
    if (!CUST_OUTPUT_FOLDER_ID) throw new Error('出力フォルダIDが未設定です。');

    const result = buildCustomersRowsSafe_(closing);
    const at = nowStamp_();

    if (!result.ok) {
      return { ok:false, message:`[請求書] CSV作成失敗: ${result.reason}（${at}）` };
    }
    const rows = result.rows || [];
    if (!rows.length){
      return { ok:false, message:`[請求書] 抽出 0件（${at}）` };
    }

    // ★修正: タイムチャージ月次集計を取り込み、請求書行にマージ
    const billingYm = deriveBillingYmFromClosing_(closing);
    let summaryRows = [];
    try {
      summaryRows = loadTimeChargeSummaryRows_(billingYm);
    } catch(e) {
      logC_('ERROR','merge timecharge summary failed',{ msg:e.message, stack:e.stack });
      return { ok:false, message:`[請求書] CSV作成失敗: タイムチャージ集計データ結合でエラー: ${e.message}` };
    }

    if (summaryRows.length) {
      const idxCode = INVOICE_HEADERS.indexOf('得意先コード');
      const idxName = INVOICE_HEADERS.indexOf('得意先名（正式名称）');
      const idxContract = INVOICE_HEADERS.indexOf('契約種別');
      const idxItem = INVOICE_HEADERS.indexOf('請求項目（概要）');
      const idxMonthly = INVOICE_HEADERS.indexOf('顧問料/月');
      const idxQty = INVOICE_HEADERS.indexOf('数量');
      const idxInternal = INVOICE_HEADERS.indexOf('備考（社内）');
      const idxClosing = INVOICE_HEADERS.indexOf('請求締日');
      const idxIssue = INVOICE_HEADERS.indexOf('発行日');
      const idxDue = INVOICE_HEADERS.indexOf('支払期限');

      const issue = today_();
      const due = endOfMonth_(issue);
      const issueStr = fmtYmd_(issue);
      const dueStr = fmtYmd_(due);
      const closingStr = fmtYmd_(closing);

      const baseByCustCode = new Map();
      for (const row of rows) {
        const code = String(row[idxCode] || '').trim();
        if (code && !baseByCustCode.has(code)) {
          baseByCustCode.set(code, row);
        }
      }

      summaryRows.forEach(row => {
        const base = baseByCustCode.get(String(row.custCode || '').trim());
        const newRow = base ? base.slice() : new Array(INVOICE_HEADERS.length).fill('');

        if (!base) {
          if (idxCode >= 0) newRow[idxCode] = row.custCode || '';
          if (idxName >= 0) newRow[idxName] = row.custName || '';
          if (idxClosing >= 0) newRow[idxClosing] = closingStr;
          if (idxIssue >= 0) newRow[idxIssue] = issueStr;
          if (idxDue >= 0) newRow[idxDue] = dueStr;
        }

        const itemParts = [row.billType, row.caseCode, row.caseName].filter(v=>String(v||'').trim());
        if (idxContract >= 0) newRow[idxContract] = row.billType || '';
        if (idxItem >= 0) newRow[idxItem] = itemParts.join(' ');
        if (idxMonthly >= 0) newRow[idxMonthly] = row.unitPerMin;
        if (idxQty >= 0) newRow[idxQty] = row.overMinutes;
        if (idxInternal >= 0) newRow[idxInternal] = row.internalNote || newRow[idxInternal];

        rows.push(newRow);
      });
    }

    const csv  = toCsv_([INVOICE_HEADERS].concat(rows));
    const tz   = Session.getScriptTimeZone();
    const now  = new Date();
    const name = 'MakeLeaps請求CSV_' + Utilities.formatDate(now,tz,'yyyyMMdd_HHmmss') + '.csv';
    const file = DriveApp.getFolderById(CUST_OUTPUT_FOLDER_ID).createFile(name, csv, MimeType.CSV);

    logC_('INFO','invoice CSV created',{ rows:rows.length, fileId:file.getId() });
    return { ok:true, url:file.getUrl(), fileId:file.getId(), filename:name, rows:rows.length,
             completedAt: at, message:`[請求書] CSV作成完了: ${rows.length}件（${at}）` };
  }catch(e){
    logC_('ERROR','createCsv_CustomersByClosing failed',{ msg:e.message, stack:e.stack });
    return { ok:false, message:e.message };
  }
}

/* ================== 本体：行生成（安全ラッパ） ================== */
function buildCustomersRowsSafe_(closingDate){
  try{
    const rows = buildCustomersRows_(closingDate);
    return { ok:true, rows };
  }catch(e){
    return { ok:false, reason: e && e.message ? e.message : '不明なエラー' };
  }
}

/* ================== 本体：行生成（エイリアス対応 & 統計ログ） ================== */
/**
 * ルール：
 *  - 「請求月」区分が「当月請求」→ pivot = (締日 + 1ヶ月)
 *  - 「請求月」区分が「翌月請求」→ pivot = 締日
 *  - 取引開始日 ≤ pivot ≤ 取引終了日 で採用（終了日空は上限なし）
 *  - 発行日 = 実行日、支払期限 = 発行日の当月末
 */
function buildCustomersRows_(closingDate){
  const sh = resolveCustomerSheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(x=>String(x||'').trim());
  const idxBy = (logic)=> {
    const list = (CUST_ALIASES[logic] || [logic]);
    for (const name of list) {
      const pos = headers.indexOf(name);
      if (pos >= 0) return pos;
    }
    return -1;
  };

  // 必須列
  const iCode   = idxBy('得意先コード');
  const iName   = idxBy('得意先名（正式名称）');
  const iKbn    = idxBy('請求月');
  const iStart  = idxBy('取引開始');
  const iEnd    = idxBy('取引終了');

  if (iCode<0 || iName<0 || iKbn<0 || iStart<0) {
    throw new Error('得意先マスタの必須列が不足しています（得意先コード/得意先名（正式名称）/請求月/取引開始[日]）');
  }

  // 発行日/支払期限
  const issue = today_();
  const due   = endOfMonth_(issue);
  const issueStr   = fmtYmd_(issue);
  const dueStr     = fmtYmd_(due);
  const closingStr = fmtYmd_(closingDate);

  const pivotThis  = new Date(closingDate);             // 翌月請求: pivot=締日
  const pivotNext  = addMonthsKeepEnd_(closingDate, 1); // 当月請求: pivot=締日+1M

  const pick = (row, colName)=>{
    const i = headers.indexOf(colName);
    return i>=0 ? row[i] : '';
  };

  const out = [];

  // 統計
  let scanned=0, picked=0, skippedNoStart=0, skippedNoKbn=0, skippedOutRange=0;

  for (let r=1; r<values.length; r++){
    scanned++;
    const row = values[r];
    const code = String(row[iCode]||'').trim();
    const name = String(row[iName]||'').trim();
    const kbn  = String(row[iKbn]||'').trim();
    if (!code || !name || !kbn) { if(!kbn) skippedNoKbn++; continue; }

    // 8桁テキスト対応の緩いパース
    const startD = parseDateLoose_(row[iStart]);
    const endD   = (iEnd>=0 ? parseDateLoose_(row[iEnd]) : null);
    if (!startD) { skippedNoStart++; continue; }

    const pivot = (kbn === '当月請求') ? pivotNext : (kbn === '翌月請求' ? pivotThis : null);
    if (!pivot) { skippedNoKbn++; continue; }
    if (!inRangeIncl_(pivot, startD, endD)) { skippedOutRange++; continue; }

    const csvRow = [
      code,                                   // 得意先コード
      name,                                   // 得意先名（正式名称）
      '',                                     // 案件名（後日ケースから）
      kbn,                                    // 請求月（区分）
      dueStr,                                 // 支払期限 = 発行日の当月末
      String(pick(row,'請求日')||''),         // 請求日（任意）
      issueStr,                               // 発行日 = 実行日
      closingStr,                             // 請求締日

      String(pick(row,'法人区分')||''),
      String(pick(row,'契約パターン')||''),
      String(pick(row,'契約種別')||''),
      String(pick(row,'超過請求')||''),
      String(pick(row,'請求元名称')||''),
      String(pick(row,'源泉徴収')||''),

      String(pick(row,'インボイス番号')||''),
      String(pick(row,'振込口座情報')||''),
      String(pick(row,'振込先口座1')||''),
      String(pick(row,'振込先口座2')||''),
      String(pick(row,'振込先口座3')||''),

      String(pick(row,'請求先宛名')||''),
      String(pick(row,'請求項目（概要）')||''),
      String(pick(row,'顧問ケースNo')||pick(row,'顧問ケースNo.')||''),
      String(pick(row,'顧問料/月')||''),
      String(pick(row,'数量')||''),

      String(pick(row,'カバー時間/分')||''),
      String(pick(row,'タイムチャージ')||''),
      String(pick(row,'タイムチャージ/分')||''),

      String(pick(row,'請求書肩書き')||''),
      String(pick(row,'請求書記載の担当者')||''),
      String(pick(row,'相手担当者')||''),

      String(pick(row,'請求先郵便番号')||''),
      String(pick(row,'請求先住所１')||''),
      String(pick(row,'請求先住所２')||''),
      String(pick(row,'請求備考')||''),
      String(pick(row,'備考（社内）')||''),

      fmtYmd_(startD),                        // 取引開始（確認用）
      endD ? fmtYmd_(endD) : ''               // 取引終了（確認用）
    ];
    out.push(csvRow);
    picked++;
  }

  logC_('INFO','customers preview stats', { scanned, picked, skippedNoStart, skippedNoKbn, skippedOutRange });

  out.sort((a,b)=>{
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    return String(a[1]).localeCompare(String(b[1]), 'ja');
  });

  return out;
}

/* ================== ユーティリティ ================== */
function resolveCustomerSheet_(){
  const ss = SpreadsheetApp.openById(CUST_DB_SPREADSHEET_ID);
  for (const name of CUST_SHEET_CANDIDATES){
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  const list = ss.getSheets().map(s=>s.getName());
  throw new Error('得意先マスタのシートが見つかりません（候補: ' + CUST_SHEET_CANDIDATES.join(', ') + '）/ 存在: ' + JSON.stringify(list));
}

// 8桁YYYYMMDD・全角数字・- . 区切りも吸収する緩いパーサ
function parseDateLoose_(v){
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    const d = v;
    return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const s = String(v).trim()
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[.\-]/g, '/');
  const digits = s.replace(/\D/g, '');
  if (/^\d{8}$/.test(digits)) {
    const y = +digits.slice(0,4), m = +digits.slice(4,6) - 1, d = +digits.slice(6,8);
    const dt = new Date(y, m, d);
    return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
}

function parseDateStrict_(v){
  if (!v) return null;
  if (Object.prototype.toString.call(v)==='[object Date]') return v;
  const s = String(v).trim().replace(/[.\-]/g,'/');
  const digits = s.replace(/\D/g,'');
  if (/^\d{8}$/.test(digits)) {
    const y=+digits.slice(0,4), m=+digits.slice(4,6)-1, d=+digits.slice(6,8);
    const dt = new Date(y,m,d); return isNaN(dt.getTime()) ? null : dt;
  }
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(s)) {
    const [y,m,d] = s.split('/').map(Number);
    const dt = new Date(y, m-1, d); return isNaN(dt.getTime()) ? null : dt;
  }
  return null;
}

function inRangeIncl_(pivot, start, end){
  const p = norm_(pivot).getTime();
  const s = norm_(start).getTime();
  const e = end ? norm_(end).getTime() : Number.POSITIVE_INFINITY;
  return (s <= p && p <= e);
}

function addMonthsKeepEnd_(date, months){
  const y=date.getFullYear(), m=date.getMonth(), d=date.getDate();
  const t=new Date(y, m+months, 1);
  const last=new Date(t.getFullYear(), t.getMonth()+1, 0).getDate();
  t.setDate(Math.min(d, last));
  return norm_(t);
}

function endOfMonth_(date){
  const y=date.getFullYear(), m=date.getMonth();
  const e = new Date(y, m+1, 0);
  return norm_(e);
}

function today_(){
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const s = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  const [Y,M,D] = s.split('/').map(Number);
  return new Date(Y, M-1, D);
}

function fmtYmd_(d){
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function norm_(d){ d.setHours(0,0,0,0); return d; }

function toCsv_(rows2d){
  return rows2d.map(r=> r.map(x=>{
    const s = (x==null)?'':String(x);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  }).join(',')).join('\n');
}

function nowStamp_(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function logC_(level, msg, data){
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  try{ Logger.log('%s [%s] %s %s', ts, level, msg, data?JSON.stringify(data):''); }catch(_){}}

// ★修正: 締日から請求対象年月(YYYY/MM)を導出
function deriveBillingYmFromClosing_(closing){
  if (!closing || Object.prototype.toString.call(closing) !== '[object Date]') return '';
  const y = closing.getFullYear();
  const m = closing.getMonth() + 1;
  return y + '/' + String(m).padStart(2,'0');
}

// ★修正: タイムチャージ月次集計（Summaryシート）を読み込む
function loadTimeChargeSummaryRows_(billingYm){
  if (!billingYm) return [];
  const filename = 'TimeCharge_Summary_' + String(billingYm).replace(/\//g, '-');
  try {
    const folder = DriveApp.getFolderById(TC_SUMMARY_FOLDER_ID);
    const it = folder.getFilesByName(filename);
    if (!it.hasNext()) return [];
    const file = it.next();
    const ss = SpreadsheetApp.openById(file.getId());
    const sh = ss.getSheetByName('Summary');
    if (!sh) return [];

    const values = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return [];
    const headers = values[0].map(h=>String(h||'').trim());

    const required = [
      '担当者','請求種別','得意先コード','得意先名','ケースコード','ケース名','単価','カバー時間',
      '合計作業時間','超過時間','請求金額','請求データ作成FLG','請求データ作成日','請求データ作成者','備考（社内）'
    ];
    const idx = {};
    for (const name of required) {
      const pos = headers.indexOf(name);
      if (pos < 0) throw new Error('Summaryシートのヘッダ不足: ' + name);
      idx[name] = pos;
    }

    const toNum = (v)=>{
      const n = Number(String(v||'').replace(/,/g,''));
      return isNaN(n) ? 0 : n;
    };

    const out = [];
    for (let r=1; r<values.length; r++){
      const row = values[r];
      out.push({
        person: row[idx['担当者']] || '',
        billType: row[idx['請求種別']] || '',
        custCode: row[idx['得意先コード']] || '',
        custName: row[idx['得意先名']] || '',
        caseCode: row[idx['ケースコード']] || '',
        caseName: row[idx['ケース名']] || '',
        unitPerMin: toNum(row[idx['単価']]),
        coverMinutes: toNum(row[idx['カバー時間']]),
        totalMinutes: toNum(row[idx['合計作業時間']]),
        overMinutes: toNum(row[idx['超過時間']]),
        amount: toNum(row[idx['請求金額']]),
        flag: row[idx['請求データ作成FLG']] || '',
        createdAt: row[idx['請求データ作成日']] || '',
        createdBy: row[idx['請求データ作成者']] || '',
        internalNote: row[idx['備考（社内）']] || '',
      });
    }
    return out;
  } catch(e) {
    if (String(e && e.message || '').indexOf('Summaryシートのヘッダ不足') >= 0) throw e;
    logC_('WARN','loadTimeChargeSummaryRows_ failed',{ msg:e.message, stack:e.stack });
    return [];
  }
}
