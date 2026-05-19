// ============================================================
// エー・ファクトリー 営業管理システム - Apps Script v3
// 営業・商材・案件をスプレッドシートで管理するバージョン
// ============================================================
// 【シート構成】
//   「設定_営業」シート  ：営業名・役職・ステータス・個人コード
//   「設定_商材」シート  ：商材名・金額・案件種別・インセンティブ方式
//   「案件マスタ」シート ：案件の全データ（新規追加）
//   「営業実績」シート   ：日報データが自動蓄積
// ============================================================

const SHEET_RESULT        = '営業実績';
const SHEET_CUSTOMERS     = '顧客DB';
const SHEET_PERSONS       = '設定_営業';
const SHEET_PRODUCTS      = '設定_商材';
const SHEET_GOALS         = '設定_目標';
const SHEET_GOALS_HISTORY = '目標履歴';
const SHEET_DEALS         = '案件マスタ';
const SHEET_ACTIVITIES    = '日次活動';
const SHEET_WEEKLY_GOALS  = '設定_週次目標';

const HEADERS = [
  '送信日時','日付','担当者','商材',
  '自分でアポ','テレアポ経由','紹介会社','代理店紹介',
  'ビジェント','代理店開拓',
  'ヒアリング','提案','クロージング','売上(万円)',
  'メモ','ステータス','企業名','プロジェクト名','費用(万円)','粗利(万円)'
];

// 商材マスタの列定義（v3: 種別・インセンティブ率に簡略化）
const PRODUCT_HEADERS_V3 = [
  '商材名','種別','売上単価（円）','費用（円）','インセンティブ率','価格タイプ'
];

// 案件マスタの列定義
const DEAL_HEADERS = [
  '案件ID','登録日','担当者','顧客ID','会社名','商材名',
  'フェーズ','確度ランク',
  '売上（単価）','費用（単価）','コース数','件数','月数',
  '売上予定額','費用（合計）','粗利',
  'インセンティブ','売上予定月','入金ステータス','入金確認日',
  'メモ','引継担当者','引継日','担当者コード','理由','最終更新日',
  '計上会社','B売上単価','B費用単価','B件数'
];

// ============================================================
// 会社名正規化（法人格・空白除去・小文字統一）
// "株式会社loty" と "loty" を同一とみなすための共通ヘルパー
// ============================================================
function normalizeCompany_(s) {
  return String(s || '').trim()
    .replace(/株式会社|有限会社|合同会社|一般社団法人|一般財団法人|特定非営利活動法人/g, '')
    .replace(/[\s　（()）・]/g, '')
    .toLowerCase();
}

// ============================================================
// GETリクエスト
// ============================================================
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'form';
  if (mode === 'all')       return getAllData();
  if (mode === 'ai')        return getAIAdvice(e && e.parameter);
  if (mode === 'data')      return getData();
  if (mode === 'master') {
    const cache = CacheService.getScriptCache();
    const cachedMaster = cache.get(MASTER_CACHE_KEY);
    if (cachedMaster) {
      return ContentService.createTextOutput(cachedMaster).setMimeType(ContentService.MimeType.JSON);
    }
    var m = getMaster();
    var mObj = JSON.parse(m.getContent());
    mObj.topProducts = getTopProducts();
    mObj.companies   = getCompanies();
    const masterPayload = JSON.stringify(mObj);
    try { cache.put(MASTER_CACHE_KEY, masterPayload, 300); } catch(_) {} // 5分キャッシュ
    return ContentService.createTextOutput(masterPayload).setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'goals')           return getGoals(e && e.parameter && e.parameter.fy);
  if (mode === 'goals_history')   return getGoalsHistory(e && e.parameter && e.parameter.fy);
  if (mode === 'generateTestData') return generateTestData();
  if (mode === 'clearTestData')    return clearTestData();
  if (mode === 'customers') return getCustomerList();
  if (mode === 'customer')  return getCustomerDetail(e && e.parameter && e.parameter.code);
  if (mode === 'deals')     return getDeals(e && e.parameter && e.parameter.person);
  if (mode === 'deal')          return getDeal(e && e.parameter && e.parameter.id);
  if (mode === 'detectIndustry') return json(detectIndustry(e && e.parameter && e.parameter.companyName));
  if (mode === 'weeklyData')    return json(getWeeklyData(e.parameter.person, e.parameter.weekStart));
  if (mode === 'weeklyAdvice')  return json(getWeeklyAdvice(e.parameter.person, e.parameter.weekStart));
  return HtmlService
    .createHtmlOutput(buildFormHtml())
    .setTitle('営業日報入力 - エー・ファクトリー')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTopProducts() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 当日キャッシュがあればそのまま返す
  if (props.getProperty('TOP_PRODUCTS_DATE') === today) {
    return JSON.parse(props.getProperty('TOP_PRODUCTS') || '[]');
  }

  // 案件マスタから直近30日の商材使用頻度を集計
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const rows = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  const counts = {};
  for (let i = 1; i < rows.length; i++) {
    const raw = rows[i][1]; // B列：登録日
    const name = rows[i][5]; // F列：商材名
    if (!name) continue;
    const d = (typeof raw.getFullYear === 'function') ? raw : new Date(raw);
    if (isNaN(d) || d < cutoff) continue;
    counts[name] = (counts[name] || 0) + 1;
  }

  const top10 = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name]) => name);

  props.setProperty('TOP_PRODUCTS', JSON.stringify(top10));
  props.setProperty('TOP_PRODUCTS_DATE', today);
  return top10;
}

// ============================================================
// CacheService ヘルパー
// ============================================================
const MASTER_CACHE_KEY = 'master_response_v1';

function invalidateMasterCache_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([MASTER_CACHE_KEY, MASTER_DATA_CACHE_KEY]);
  } catch(_) {}
}

// ============================================================
// LockService ヘルパー（同時書き込み防止）
// ============================================================
function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return json({ success: false, error: '他の処理が実行中です。しばらくしてから再試行してください。' });
    }
    return fn();
  } catch (err) {
    return json({ success: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// ============================================================
// POSTリクエスト
// ============================================================
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const action = d.action || 'entry';

    // ── 読み取り専用（ロック不要） ──────────────────────────
    if (action === 'getCustomers')     return json(getCustomers(d));
    if (action === 'getCustomerStats') return json(getCustomerStats(d.afcStaff));
    if (action === 'getContactMaster') return json({ success: true, data: getContacts(d.customerId || '') });
    if (action === 'detectIndustry')  return json({ success: true, data: detectIndustry(d.companyName) });
    if (action === 'migrateOldDB')    return json({ success: true, data: migrateOldCustomerDB() });
    if (action === 'clearAll')        return json({ success: false, error: '外部からの実行は許可されていません' });

    // ── 書き込み系（LockService で保護） ───────────────────
    return withLock(() => {
      if (action === 'addPerson')           return addPerson(d.name, d.role, d.code || '');
      if (action === 'setPersonStatus')     return setPersonStatus(d.name, d.status);
      if (action === 'deletePerson')        return deletePerson(d.name);
      if (action === 'addProduct')          return addProductToSheet(d.name, d.kind, d.unitPrice, d.cost, d.incentiveRate, d.priceType);
      if (action === 'saveGoals')           return saveGoals(d);
      if (action === 'changeCompanyPerson') return changeCompanyPerson(d.code, d.person);
      if (action === 'initCustomerHistory') return initCustomerHistoryColumn();
      if (action === 'addDeal')             return addDeal(d);
      if (action === 'updateDeal')          return updateDeal(d);
      if (action === 'updateDealStatus')    return updateDealStatus(d.id, d.phase, d.rankLabel);
      if (action === 'confirmPayment')      return confirmPayment(d.id, d.date);
      if (action === 'handoverDeal')        return handoverDeal(d.id, d.newPerson, d.date);
      if (action === 'addCustomer')         return json({ success: true, data: addCustomer(d) });
      if (action === 'updateCustomer')      return json({ success: true, data: updateCustomer(d) });
      if (action === 'deleteCustomer')      return json({ success: true, data: deleteCustomer(d.id) });
      if (action === 'addContact')          return json({ success: true, data: addContact(d) });
      if (action === 'updateContact')       return json({ success: true, data: updateContact(d) });
      if (action === 'deleteContact')       return json({ success: true, data: deleteContact(d.id) });
      if (action === 'addActivity')         return json(addActivity(d));
      if (action === 'setWeeklyTarget')     return json(setWeeklyTarget(d));
      if (action === 'setMonthlyKPITarget') return json(setMonthlyKPITarget(d));
      return addEntry(d);
    });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// ============================================================
// 案件登録
// ============================================================
const VALID_DEAL_RANKS = ['売上', '決定', 'A', 'B', 'C', '失注'];

function addDeal(d) {
  if (!d.person || !String(d.person).trim())           return json({ success: false, error: '営業担当（個人コード）が空です' });
  if (!d.companyName || !String(d.companyName).trim()) return json({ success: false, error: '会社名が空です' });
  if (!d.expectedMonth || !String(d.expectedMonth).trim()) return json({ success: false, error: '売上予定月が空です' });
  if (!/^\d{4}-\d{2}$/.test(String(d.expectedMonth).trim())) return json({ success: false, error: '売上予定月の形式が不正です（例: 2026-04）' });
  if (d.rankLabel && !VALID_DEAL_RANKS.includes(String(d.rankLabel).trim())) return json({ success: false, error: '確度ランクが無効です: ' + d.rankLabel });

  const sheet = getOrCreateDealSheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const id = generateDealId();

  const pDetail = getProductDetail(d.productName);

  // フォームから単価・コース数・件数・月数を受け取り、合計を計算
  const unitSales = Number(d.unitSales) || (pDetail ? pDetail.unitPrice : 0);
  const unitCost  = Number(d.unitCost)  || (pDetail ? pDetail.cost : 0);
  const courses   = Math.max(1, Number(d.courses) || 1);
  const qty       = Math.max(1, Number(d.qty)     || 1);
  const months    = Math.max(1, Math.min(24, Number(d.months) || (pDetail ? pDetail.months : 1)));

  const totalSales  = unitSales * courses * qty * months;
  const totalCost   = unitCost  * courses * qty * months;
  const grossProfit = (unitSales - unitCost) * courses * qty * months;
  const monthlyGP   = (unitSales - unitCost) * courses * qty;

  const incentiveRate = pDetail ? pDetail.incentiveRate : 0;
  const incentive = calcIncentive(monthlyGP, months, incentiveRate);

sheet.appendRow([
    id,                          // A: 案件ID
    today,                       // B: 登録日
    d.person || '',              // C: 担当者
    d.customerId || '',          // D: 顧客ID
    d.companyName || '',         // E: 会社名
    d.productName || '',         // F: 商材名
    d.phase || 'ヒアリング中',   // G: フェーズ
    d.rankLabel || 'C',          // H: 確度ランク
    unitSales,                   // I: 売上（単価）
    unitCost,                    // J: 費用（単価）
    courses,                     // K: コース数
    qty,                         // L: 件数
    months,                      // M: 月数
    totalSales,                  // N: 売上予定額（合計）
    totalCost,                   // O: 費用（合計）
    grossProfit,                 // P: 粗利（合計）
    incentive,                   // Q: インセンティブ
    d.expectedMonth || '',       // R: 売上予定月
    '未入金',                    // S: 入金ステータス
    '',                          // T: 入金確認日
    d.memo || '',                // U: メモ
    '',                          // V: 引継担当者
    '',                          // W: 引継日
    '',                          // X: 担当者コード
    d.reason || '',              // Y: 理由
    today,                       // Z: 最終更新日
    d.billingCompany || '',      // AA: 計上会社
    Number(d.bUnitSales) || 0,  // AB: B売上単価
    Number(d.bUnitCost)  || 0,  // AC: B費用単価
    Number(d.bQty)       || 0   // AD: B件数
  ]);

  return json({ success: true, id, incentive, grossProfit });
}

// ============================================================
// 案件一覧取得（営業フィルタ対応）
// ============================================================
function getDeals(person) {
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: true, deals: [], count: 0 });

  // 個人コード→営業名の解決マップを構築
  const personCodeMap = {};
  getPersonDetails().forEach(p => { if (p.code) personCodeMap[p.code] = p.name; });

  const vals = sheet.getRange(1, 1, lastRow, DEAL_HEADERS.length).getValues();
  const headers = vals[0];

  const MON_MAP = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',
                   Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  function normYM(v) {
    if (!v) return '';
    // GASのスプレッドシートDate型はinstanceof Dateが効かない場合がある → duck typing
    if (typeof v === 'object' && typeof v.getFullYear === 'function') {
      return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
    }
    const s = String(v).trim();
    // ISO文字列: "2025-07-31T15:00:00.000Z"
    const mIso = s.match(/^(\d{4})-(\d{2})-\d{2}T/);
    if (mIso) return mIso[1]+'-'+mIso[2];
    // "2025-08" 形式はそのまま
    if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
    // "Aug-25" 形式 → "2025-08"
    const m1 = s.match(/^([A-Z][a-z]{2})-(\d{2})$/i);
    if (m1) {
      const key = m1[1].charAt(0).toUpperCase()+m1[1].slice(1,3).toLowerCase();
      const yr = parseInt(m1[2]);
      return (yr <= 50 ? '20' : '19') + m1[2] + '-' + (MON_MAP[key] || '01');
    }
    return '';
  }
  function normDate(v) {
    if (!v) return '';
    if (typeof v === 'object' && typeof v.getFullYear === 'function') {
      return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    const s = String(v).trim();
    const mIso = s.match(/^(\d{4}-\d{2}-\d{2})T/);
    if (mIso) return mIso[1];
    return s;
  }

  let deals = vals.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    let rawPersonCode = '';
    headers.forEach((h, i) => {
      if (h === '担当者コード') {
        // X列は後でC列の値で上書きするためスキップ
        return;
      } else if (h === '担当者') {
        // C列: 個人コード→名前に解決（旧データ=名前文字列はそのまま）
        rawPersonCode = String(r[i] || '').trim();
        o[h] = personCodeMap[rawPersonCode] || rawPersonCode;
      } else if (h === '売上予定月') {
        o[h] = normYM(r[i]);
      } else if (h === '登録日' || h === '入金確認日' || h === '引継日') {
        o[h] = normDate(r[i]);
      } else if (typeof r[i] === 'object' && typeof r[i].getFullYear === 'function') {
        o[h] = Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        o[h] = r[i];
      }
    });
    // C列の生コード（個人コード）を担当者コードとして公開
    o['担当者コード'] = rawPersonCode;
    return o;
  });

  if (person) deals = deals.filter(d => d['担当者'] === person);

  return json({ success: true, deals, count: deals.length });
}

// ============================================================
// 案件詳細取得
// ============================================================
function getDeal(id) {
  if (!id) return json({ success: false, error: 'IDが空です' });
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: false, error: '案件が見つかりません' });

  const vals = sheet.getRange(1, 1, lastRow, DEAL_HEADERS.length).getValues();
  const headers = vals[0];
  const row = vals.slice(1).find(r => String(r[0]).trim() === id.trim());
  if (!row) return json({ success: false, error: '見つかりません: ' + id });

  const deal = {};
  headers.forEach((h, i) => {
    deal[h] = row[i] instanceof Date
      ? Utilities.formatDate(row[i], 'Asia/Tokyo', 'yyyy-MM-dd') : row[i];
  });
  return json({ success: true, deal });
}

// ============================================================
// 案件更新（フェーズ・確度・金額など）
// ============================================================
function updateDeal(d) {
  if (!d.id) return json({ success: false, error: 'IDが空です' });
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, DEAL_HEADERS.length).getValues();

  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === d.id.trim()) {
      const rowNum = i + 2;
      const rowData = vals[i]; // メモリ上の行データを直接編集し、最後に1回だけ書き込む
      const setField = (key, val) => {
        if (val === undefined || val === null) return;
        const idx = DEAL_HEADERS.indexOf(key);
        if (idx >= 0) rowData[idx] = val;
      };

      // 常に更新可能なフィールド
      setField('確度ランク', d.rankLabel);
      setField('売上予定月', d.expectedMonth);
      setField('メモ', d.memo);
      setField('会社名', d.companyName);
      setField('顧客ID', d.customerId);
      setField('フェーズ', d.phase);
      setField('理由', d.reason);
      setField('入金ステータス', d.paymentStatus);
      setField('計上会社', d.billingCompany);
      if (d.bUnitSales !== undefined) setField('B売上単価', Number(d.bUnitSales) || 0);
      if (d.bUnitCost  !== undefined) setField('B費用単価', Number(d.bUnitCost)  || 0);
      if (d.bQty       !== undefined) setField('B件数',     Number(d.bQty)       || 0);

      const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

      // ダッシュボードモーダルからの粗利直接更新
      if (d.grossProfit !== undefined) {
        setField('粗利', Number(d.grossProfit));
        setField('最終更新日', today);
        sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
        return json({ success: true });
      }

      // フォームからの単価・コース数・件数・月数更新（全再計算）
      if (d.unitSales !== undefined || d.unitCost !== undefined ||
          d.courses !== undefined || d.qty !== undefined || d.months !== undefined) {
        const unitSales = Number(d.unitSales !== undefined ? d.unitSales : rowData[DEAL_HEADERS.indexOf('売上（単価）')]);
        const unitCost  = Number(d.unitCost  !== undefined ? d.unitCost  : rowData[DEAL_HEADERS.indexOf('費用（単価）')]);
        const courses   = Math.max(1, Number(d.courses !== undefined ? d.courses : rowData[DEAL_HEADERS.indexOf('コース数')]) || 1);
        const qty       = Math.max(1, Number(d.qty     !== undefined ? d.qty     : rowData[DEAL_HEADERS.indexOf('件数')])   || 1);
        const months    = Math.max(1, Math.min(24, Number(d.months !== undefined ? d.months : rowData[DEAL_HEADERS.indexOf('月数')]) || 1));
        const totalSales  = unitSales * courses * qty * months;
        const totalCost   = unitCost  * courses * qty * months;
        const grossProfit = (unitSales - unitCost) * courses * qty * months;
        const monthlyGP   = (unitSales - unitCost) * courses * qty;
        const pDetail = getProductDetail(String(rowData[DEAL_HEADERS.indexOf('商材名')]));
        const incentiveRate = pDetail ? pDetail.incentiveRate : 0;
        const incentive = calcIncentive(monthlyGP, months, incentiveRate);
        setField('売上（単価）', unitSales);
        setField('費用（単価）', unitCost);
        setField('コース数', courses);
        setField('件数', qty);
        setField('月数', months);
        setField('売上予定額', totalSales);
        setField('費用（合計）', totalCost);
        setField('粗利', grossProfit);
        setField('インセンティブ', incentive);
        setField('最終更新日', today);
        sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
        return json({ success: true, grossProfit, incentive });
      }

      setField('最終更新日', today);
      sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      return json({ success: true });
    }
  }
  return json({ success: false, error: '見つかりません: ' + d.id });
}

// ============================================================
// 入金確認（継続案件の月次入金対応）
// ============================================================
function confirmPayment(id, payDate) {
  if (!id) return json({ success: false, error: 'IDが空です' });
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === id.trim()) {
      const row = i + 2;
      const dateStr = payDate || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      sheet.getRange(row, DEAL_HEADERS.indexOf('入金ステータス') + 1).setValue('入金済み');
      sheet.getRange(row, DEAL_HEADERS.indexOf('入金確認日') + 1).setValue(dateStr);
      return json({ success: true });
    }
  }
  return json({ success: false, error: '見つかりません: ' + id });
}

// ============================================================
// 案件引継（退職時）
// ============================================================
function handoverDeal(id, newPerson, handoverDate) {
  if (!id || !newPerson) return json({ success: false, error: 'IDまたは引継先が空です' });
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, DEAL_HEADERS.length).getValues();

  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === id.trim()) {
      const row = i + 2;
      const dateStr = handoverDate || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      sheet.getRange(row, DEAL_HEADERS.indexOf('引継担当者') + 1).setValue(newPerson);
      sheet.getRange(row, DEAL_HEADERS.indexOf('引継日') + 1).setValue(dateStr);
      return json({ success: true });
    }
  }
  return json({ success: false, error: '見つかりません: ' + id });
}

// ============================================================
// テストデータ生成（管理者用：GETで呼び出し）
// ============================================================
function generateTestData() {
  const sheet = getOrCreateDealSheet();
  clearTestData(); // 既存テストデータを先にクリア

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const thisYM = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  const nextD  = new Date(); nextD.setMonth(nextD.getMonth() + 1);
  const nextYM = Utilities.formatDate(nextD, 'Asia/Tokyo', 'yyyy-MM');

  // 列順: A案件ID B登録日 C担当者 D顧客ID E会社名 F商材名 Gフェーズ H確度ランク
  //       I売上単価 J費用単価 Kコース数 L件数 M月数 N売上予定額 O費用合計 P粗利
  //       Qインセンティブ R売上予定月 S入金ステータス T入金確認日 Uメモ
  //       V引継担当者 W引継日 X担当者コード Y理由 Z最終更新日
  //       AA計上会社 ABB売上単価 ACB費用単価 ADB件数
  const rows = [
    // TEST-001: 正常な売上確定（ストック12ヶ月・インセンティブあり）
    ['TEST-001', today, 'テスト担当者', '', '株式会社LOTY', 'HubCast_直販',
     '売上', '売上', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, Math.floor((35000-13000)*12*0.05),
     thisYM, '入金済み', today, '[TEST]001_正常売上_ストック12',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-002: 表記揺れ検証（' loTy ' → normalizeCompanyで株式会社LOTYと同一扱いになるか）
    ['TEST-002', today, 'テスト担当者', '', ' loTy ', 'HubCast_直販',
     'ヒアリング中', 'A', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, 0,
     nextYM, '未入金', '', '[TEST]002_表記揺れ(loTy)',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-003: ストック按分テスト（件数2・月数12で月次粗利が半分になるか）
    ['TEST-003', today, 'テスト担当者', '', 'ストックテスト株式会社', 'HubCast_直販',
     'ヒアリング中', 'B', 35000, 13000, 1, 2, 12,
     35000*2*12, 13000*2*12, (35000-13000)*2*12, 0,
     thisYM, '未入金', '', '[TEST]003_ストック按分（件数2）',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-004: スポット商材（月数1）
    ['TEST-004', today, 'テスト担当者', '', 'スポットテスト株式会社', 'IT支援研修事業',
     'クロージング中', 'A', 1000000, 0, 1, 1, 1,
     1000000, 0, 1000000, Math.floor(1000000*0.03),
     nextYM, '未入金', '', '[TEST]004_スポット商材（月数1）',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-005: 売上-費用=粗利 計算整合性テスト
    ['TEST-005', today, 'テスト担当者', '', '計算検証株式会社', 'リスキリング補助_直販',
     '提案中', 'B', 200000, 0, 1, 1, 1,
     200000, 0, 200000, 0,
     nextYM, '未入金', '', '[TEST]005_計算整合性チェック',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-006: B行テスト（計上会社・B売上単価・B費用単価・B件数あり）
    ['TEST-006', today, 'テスト担当者', '', 'B行テスト株式会社', 'HubCast_代理店',
     'クロージング中', 'A', 35000, 27000, 1, 3, 12,
     35000*3*12, 27000*3*12, (35000-27000)*3*12, 0,
     nextYM, '未入金', '', '[TEST]006_B行テスト',
     '', '', '', '', today, 'エー・ファクトリー株式会社', 40000, 30000, 2],

    // TEST-007: インセンティブ計算テスト（決定ランク）
    ['TEST-007', today, 'テスト担当者', '', 'インセンティブテスト社', 'HubCast_直販',
     '売上', '決定', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, Math.floor((35000-13000)*12*0.05),
     thisYM, '未入金', '', '[TEST]007_インセンティブ計算',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-008: 複数コース数×件数のテスト
    ['TEST-008', today, 'テスト担当者', '', '複数件数テスト社', 'HubCast_直販',
     '提案中', 'C', 35000, 13000, 2, 3, 12,
     35000*2*3*12, 13000*2*3*12, (35000-13000)*2*3*12, 0,
     nextYM, '未入金', '', '[TEST]008_複数コース×件数',
     '', '', '', '', today, '', 0, 0, 0],

    // TEST-009: 失注案件（パイプライン集計から除外されるか）
    ['TEST-009', today, 'テスト担当者', '', '失注テスト株式会社', 'HubCast_直販',
     '失注', '失注', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, 0,
     nextYM, '未入金', '', '[TEST]009_失注案件',
     '', '', '', '商談が長引いた', today, '', 0, 0, 0],

    // TEST-010: 大文字・空白混じり表記揺れ（株式会社　LOTY → lotyと同一か）
    ['TEST-010', today, 'テスト担当者', '', '株式会社　LOTY', '現地長期補助_直販',
     'ヒアリング中', 'C', 600000, 0, 1, 1, 1,
     600000, 0, 600000, 0,
     nextYM, '未入金', '', '[TEST]010_表記揺れ（全角スペース入り）',
     '', '', '', '', today, '', 0, 0, 0],
  ];

  rows.forEach(r => sheet.appendRow(r));
  return json({ success: true, count: rows.length });
}

// ============================================================
// テストデータ削除（管理者用：GETで呼び出し）
// ============================================================
function clearTestData() {
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: true, count: 0 });

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const toDelete = [];
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).startsWith('TEST-')) toDelete.push(i + 2);
  }
  toDelete.forEach(r => sheet.deleteRow(r));
  return json({ success: true, count: toDelete.length });
}

// ============================================================
// インセンティブ計算
// ============================================================
function calcIncentive(grossProfitMonthly, months, incentiveRate) {
  const gp   = Number(grossProfitMonthly) || 0;
  const m    = Number(months) || 1;
  const rate = Number(incentiveRate) || 0;
  // ストック・スポット問わず totalGP × rate（スポットはm=1なので実質同じ）
  return Math.floor(gp * m * rate);
}

// ============================================================
// 案件ID生成
// ============================================================
function generateDealId() {
  const ts   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  const rand = Math.random().toString(36).substr(2, 4).toUpperCase();
  return 'DL-' + ts + '-' + rand;
}

// ============================================================
// 商材詳細1件取得（ヘルパー）
// ============================================================
function getProductDetail(name) {
  if (!name) return null;
  const sheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const vals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const row = vals.find(r => String(r[0]).trim() === String(name).trim());
  if (!row) return null;
  const kind = String(row[1] || 'スポット').trim();
  return {
    name:          String(row[0]).trim(),
    kind:          kind,
    months:        kind === 'ストック' ? 12 : 1,
    unitPrice:     Number(row[2]) || 0,
    cost:          Number(row[3]) || 0,
    incentiveRate: Number(row[4]) || 0,
    priceType:     String(row[5]).trim() || (Number(row[2]) > 0 ? '固定' : '都度見積もり'),
  };
}

// ============================================================
// 日報データ登録
// ============================================================
function addEntry(d) {
  if (!d.person || !String(d.person).trim()) return json({ success: false, error: '営業担当（個人コード）が空です' });
  if (!d.date   || !String(d.date).trim())   return json({ success: false, error: '日付が空です' });

  const sheet = getOrCreateResultSheet();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  if (d.activity !== undefined) {
    const act = d.activity || {};
    sheet.appendRow([
      now, d.date||'', d.person||'', '【活動】',
      n(act.selfApo), n(act.telApo), n(act.refApo), n(act.agentRef),
      n(act.bizent), n(act.selfAgent),
      n(act.hearing), n(act.proposal), n(act.closing), 0,
      act.memo||'', '', '', '', '', ''
    ]);
    const deals = d.deals || [];
    deals.forEach(function(deal) {
      sheet.appendRow([
        now, d.date||'', d.person||'', deal.product||'',
        0, 0, 0, 0, 0, 0, 0, 0, deal.status==='売上'?1:0,
        deal.status==='売上' ? n(deal.sales) : 0,
        '',
        deal.status||'',
        deal.company||'',
        deal.project||'',
        n(deal.cost),
        n(deal.grossProfit)
      ]);
    });
    return json({ success: true, rows: 1 + deals.length });
  }

  sheet.appendRow([
    now, d.date||'', d.person||'', d.product||'',
    n(d.selfApo), n(d.telApo), n(d.refApo), n(d.agentRef),
    n(d.bizent), n(d.selfAgent),
    n(d.hearing), n(d.proposal), n(d.closing), n(d.sales),
    '', '', '', '', '', ''
  ]);
  return json({ success: true });
}

// ============================================================
// 営業追加
// ============================================================
function addPerson(name, role, code) {
  if (!name || !name.trim()) return json({ success: false, error: '名前が空です' });
  name = name.trim();
  role = (role || 'スタッフ').trim();
  code = (code || '').trim();
  const sheet = getOrCreatePersonSheet();
  const existing = getPersonDetails();
  if (existing.some(p => p.name === name)) return json({ success: false, error: '既に存在します: ' + name });
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, 4).setValues([[name, role, '在籍中', code]]);
  invalidateMasterCache_();
  return json({ success: true, name: name, role: role, code: code });
}

// ============================================================
// 営業ステータス変更
// ============================================================
function setPersonStatus(name, status) {
  if (!name) return json({ success: false, error: '名前が空です' });
  name = name.trim();
  const sheet = getOrCreatePersonSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: false, error: '営業が見つかりません' });
  const vals = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === name) {
      sheet.getRange(i + 2, 3).setValue(status);
      invalidateMasterCache_();
      return json({ success: true, name: name, status: status });
    }
  }
  return json({ success: false, error: '見つかりません: ' + name });
}

// ============================================================
// 営業削除
// ============================================================
function deletePerson(name) {
  if (!name) return json({ success: false, error: '名前が空です' });
  name = name.trim();
  const sheet = getOrCreatePersonSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: false, error: '営業が見つかりません' });
  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === name) {
      sheet.deleteRow(i + 2);
      invalidateMasterCache_();
      return json({ success: true, name: name });
    }
  }
  return json({ success: false, error: '見つかりません: ' + name });
}

// ============================================================
// 商材追加（新列対応）
// ============================================================
function addProductToSheet(name, kind, unitPrice, cost, incentiveRate, priceType) {
  if (!name || !name.trim()) return json({ success: false, error: '商材名が空です' });
  name = name.trim();
  const sheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const existing = getListFromSheet(sheet);
  if (existing.includes(name)) return json({ success: false, error: '既に存在します: ' + name });
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, 6).setValues([[
    name,
    kind          || 'スポット',
    Number(unitPrice) || 0,
    Number(cost)      || 0,
    Number(incentiveRate) || 0,
    priceType     || '固定',
  ]]);
  invalidateMasterCache_();
  return json({ success: true, name: name });
}

// ============================================================
// 目標設定 取得
// ============================================================
// fy: 'FY2025' のような文字列。省略時は現在のFY
function getGoals(fy) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetFy = fy || getCurrentFY_();

    // まず目標履歴シートから最新版を取得
    const histSheet = ss.getSheetByName(SHEET_GOALS_HISTORY);
    if (histSheet) {
      const data = histSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === targetFy && data[i][2]) {
          try {
            return json({ success: true, goals: JSON.parse(data[i][2]), savedAt: String(data[i][1]) });
          } catch(e) {}
        }
      }
    }

    // フォールバック: 旧設定_目標シート
    const sheet = ss.getSheetByName(SHEET_GOALS);
    if (!sheet) return json({ success: true, goals: null });
    const val = String(sheet.getRange(1, 1).getValue()).trim();
    if (!val) return json({ success: true, goals: null });
    return json({ success: true, goals: JSON.parse(val) });
  } catch(err) {
    return json({ success: true, goals: null, error: err.message });
  }
}

// 目標履歴一覧取得（新しい順）
function getGoalsHistory(fy) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetFy = fy || getCurrentFY_();
    const sheet = ss.getSheetByName(SHEET_GOALS_HISTORY);
    if (!sheet) return json({ success: true, history: [] });
    const data = sheet.getDataRange().getValues();
    const history = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === targetFy && data[i][2]) {
        try {
          history.push({ savedAt: String(data[i][1]), goals: JSON.parse(data[i][2]) });
        } catch(e) {}
      }
    }
    history.reverse(); // 新しい順
    return json({ success: true, history });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// FY文字列を返すヘルパー（'FY2025' など）
function getCurrentFY_() {
  const now = new Date();
  const m = now.getMonth() + 1; // 1-12
  const y = now.getFullYear();
  return 'FY' + (m >= 8 ? y : y - 1);
}

// ============================================================
// 目標設定 保存（目標履歴シートに追記）
// ============================================================
function saveGoals(d) {
  try {
    const persons = d.goals && d.goals.persons;
    if (!persons || !Array.isArray(persons) || persons.length === 0) {
      return json({ success: false, error: '担当者データが空のため保存しませんでした' });
    }
    const hasNonZero = persons.some(p =>
      Object.values(p.months || {}).some(v => Number(v) > 0)
    );
    if (!hasNonZero) {
      return json({ success: false, error: '全担当者の目標額が0のため保存しませんでした' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fy = d.fy || getCurrentFY_();
    const goalsJson = JSON.stringify(d.goals);

    // 目標履歴シートに追記
    let sheet = ss.getSheetByName(SHEET_GOALS_HISTORY);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_GOALS_HISTORY);
      const hdr = sheet.getRange(1, 1, 1, 3);
      hdr.setValues([['年度', '保存日時', 'データ(JSON)']]);
      hdr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold');
      sheet.setColumnWidth(1, 80);
      sheet.setColumnWidth(2, 160);
      sheet.setColumnWidth(3, 800);
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([fy, now, goalsJson]);

    // 設定_目標シートにも最新データをバックアップ
    let goalsSheet = ss.getSheetByName(SHEET_GOALS);
    if (!goalsSheet) goalsSheet = ss.insertSheet(SHEET_GOALS);
    goalsSheet.getRange(1, 1).setValue(goalsJson);

    return json({ success: true });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// ============================================================
// AIアドバイス（Gemini API をGAS経由で呼び出す）
// ============================================================
function getAIAdvice(params) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return json({ success: false, error: 'GASにAPIキーが未設定です。setGeminiApiKey()を実行してください。' });

    const person = (params && params.person) || null;
    const isTeam = !person || person === '__all__';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const today = new Date();
    const currentFY = today.getMonth() >= 7 ? today.getFullYear() : today.getFullYear() - 1;
    const todayYM = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM');

    // FY月リスト
    const fyMonths = [];
    for (let m = 8; m <= 12; m++) fyMonths.push(currentFY + '-' + String(m).padStart(2,'0'));
    for (let m = 1; m <= 7;  m++) fyMonths.push((currentFY+1) + '-' + String(m).padStart(2,'0'));

    // 実績データ読み込み
    const rSheet = getOrCreateResultSheet();
    const rVals = rSheet.getDataRange().getValues();
    const rHdr = rVals[0];
    const actualData = rVals.slice(1).filter(r => r[1]).map(r => {
      const o = {};
      rHdr.forEach((h,i) => { o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i],'Asia/Tokyo','yyyy-MM-dd') : r[i]; });
      return o;
    });

    const fyActual = actualData.filter(r => {
      if (!fyMonths.includes(String(r['日付']).slice(0,7))) return false;
      return isTeam || r['担当者'] === person;
    });
    const thisMonthActual = fyActual.filter(r => String(r['日付']).slice(0,7) === todayYM);

    const actualTotal    = Math.round(fyActual.reduce((s,r) => s+(Number(r['売上(万円)'])||0),0));
    const thisMonthGP    = Math.round(thisMonthActual.reduce((s,r) => s+(Number(r['売上(万円)'])||0),0));
    const meeting        = thisMonthActual.reduce((s,r) => s+(Number(r['ヒアリング'])||0),0);
    const selfApo        = thisMonthActual.reduce((s,r) => s+(Number(r['自分でアポ'])||0),0);
    const telApo         = thisMonthActual.reduce((s,r) => s+(Number(r['テレアポ経由'])||0),0);
    const refCount       = thisMonthActual.reduce((s,r) => s+(Number(r['紹介会社'])||0),0);

    // 案件データ
    const dSheet = getOrCreateDealSheet();
    const dVals = dSheet.getDataRange().getValues();
    const dHdr = dVals[0];
    const allDeals = dVals.slice(1).filter(r=>r[0]).map(r=>{
      const o={}; dHdr.forEach((h,i)=>{o[h]=r[i];}); return o;
    });
    const activeDeals = allDeals.filter(d => (isTeam || d['担当者']===person) && d['確度ランク']!=='失注');
    const yomiTotal = activeDeals.filter(d=>['決定','A','B'].includes(String(d['確度ランク'])))
      .reduce((s,d)=>s+Math.round((Number(d['粗利'])||0)/10000),0);

    // 目標
    let annualTarget=0, monthlyTarget=0;
    const gSheet = ss.getSheetByName(SHEET_GOALS);
    if (gSheet) {
      try {
        const goals = JSON.parse(String(gSheet.getRange(1,1).getValue()).trim());
        if (isTeam) {
          annualTarget  = (goals.persons||[]).reduce((s,p)=>s+Math.round(p.salesTarget||0),0);
          monthlyTarget = (goals.persons||[]).reduce((s,p)=>s+Math.round(p.monthlyTarget||0),0);
        } else {
          const pg = (goals.persons||[]).find(p=>p.name===person)||{};
          annualTarget  = Math.round(pg.salesTarget||0);
          monthlyTarget = Math.round(pg.monthlyTarget||0);
        }
      } catch(e) {}
    }

    const elapsed = Math.max(1, today.getMonth()>=7 ? today.getMonth()-7 : today.getMonth()+5);
    const prompt = `あなたはプロの営業マネジャーです。以下の営業データを分析し、今月の改善アドバイスを日本語で3〜4点、箇条書きで具体的に提示してください。数字を必ず使ってください。

【対象】${isTeam?'チーム全体':'担当者: '+person}
【期間】FY${currentFY}（8月〜翌7月）経過${elapsed}ヶ月
【今期累計粗利】${actualTotal}万円 ／ 年間目標 ${annualTarget>0?annualTarget+'万円':'未設定'}
【今月粗利実績】${thisMonthGP}万円 ／ 月間目標 ${monthlyTarget>0?monthlyTarget+'万円':'未設定'}
【今月KPI】有効面談${meeting}件、自アポ${selfApo}件、テレアポ${telApo}件、紹介${refCount}件
【アクティブ案件】${activeDeals.length}件 ／ ヨミ合計 ${yomiTotal}万円
【確度内訳】決定:${activeDeals.filter(d=>d['確度ランク']==='決定').length}件 A:${activeDeals.filter(d=>d['確度ランク']==='A').length}件 B:${activeDeals.filter(d=>d['確度ランク']==='B').length}件 C:${activeDeals.filter(d=>d['確度ランク']==='C').length}件

アドバイスは実践的かつ前向きなトーンで。`;

    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key='+apiKey,
      { method:'POST', headers:{'Content-Type':'application/json'},
        payload:JSON.stringify({contents:[{parts:[{text:prompt}]}],
          generationConfig:{temperature:0.7,maxOutputTokens:1024}}),
        muteHttpExceptions:true }
    );
    const data = JSON.parse(res.getContentText());
    if (data.error) return json({ success:false, error:data.error.message });
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '応答がありません';
    return json({ success:true, text, person:person||'チーム全体', month:todayYM });
  } catch(err) {
    return json({ success:false, error:err.message });
  }
}

// GASエディタから実行してAPIキーを設定する
function setGeminiApiKey() {
  const key = 'ここにAPIキーを貼り付ける'; // ← APIキーに書き換えてから実行！
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('Gemini APIキーを設定しました');
}

// ============================================================
// パイプライン集計（担当者 × 確度ランク）
// ============================================================
function buildPipelineByPerson_(deals) {
  const ranks = ['売上', '決定', 'A', 'B', 'C', '失注'];
  const byPerson = {};
  deals.forEach(function(d) {
    const person = String(d['担当者'] || '');
    const rank   = String(d['確度ランク'] || '');
    const gp     = Number(d['粗利']) || 0;
    if (!person) return;
    if (!byPerson[person]) {
      byPerson[person] = {};
      ranks.forEach(function(r) { byPerson[person][r] = 0; });
    }
    if (byPerson[person][rank] !== undefined) byPerson[person][rank] += gp;
  });
  return byPerson;
}

// ============================================================
// 全データ一括取得（キャッシュ対応用）
// ============================================================
function getAllData() {
  try {
    // 各データを内部で直接取得（HTTPコール不要）
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // master
    const masterResult = JSON.parse(getMaster().getContent());

    // goals（目標履歴シートから最新版を取得）
    let goalsVal = null;
    try {
      const goalsResult = JSON.parse(getGoals(getCurrentFY_()).getContent());
      if (goalsResult.success) goalsVal = goalsResult.goals;
    } catch(e) {}

    // actual data
    const resultSheet = getOrCreateResultSheet();
    const rVals = resultSheet.getDataRange().getValues();
    let actualRows = [];
    if (rVals.length > 1) {
      const headers = rVals[0];
      actualRows = rVals.slice(1).filter(r => r[1]).map(r => {
        const o = {};
        headers.forEach((h, i) => {
          o[h] = r[i] instanceof Date
            ? Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd') : r[i];
        });
        return o;
      });
    }

    // deals
    const dealsResult = JSON.parse(getDeals(null).getContent());

    // パイプライン by 担当者（確度ランク × 担当者別 粗利集計）
    const pipelineByPerson = buildPipelineByPerson_(dealsResult.deals || []);

    // 確度マッピング（確度ランク → 着地確率）
    const confidenceMap = getConfidenceMapping_();

    // 全担当者の今週行動KPI集計
    const weeklyKpi = getWeeklyKPISummaryAll_(getThisWeekMonday_());

    return json({
      success: true,
      master: masterResult,
      goals: { success: true, goals: goalsVal },
      actual: { success: true, data: actualRows, count: actualRows.length },
      deals: dealsResult,
      pipelineByPerson,
      confidenceMap,
      weeklyKpi,
      cachedAt: new Date().toISOString()
    });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// ============================================================
// 確度マッピング読み込み（「確度マッピング」シートまたはデフォルト値）
// シート構成: A列=確度ランク名, B列=確率(0〜100の整数または0.0〜1.0の小数)
// ============================================================
function getConfidenceMapping_() {
  const defaults = { 売上: 1.0, 決定: 0.9, A: 0.8, B: 0.5, C: 0.2, 失注: 0 };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('確度マッピング');
    if (!sh || sh.getLastRow() <= 1) return defaults;
    const rows = sh.getDataRange().getValues().slice(1);
    rows.forEach(function(r) {
      const rank = String(r[0]).trim();
      let pct = Number(r[1]);
      if (!rank || isNaN(pct)) return;
      // 1より大きい値はパーセント表記とみなして0〜1に変換
      if (pct > 1) pct = pct / 100;
      defaults[rank] = pct;
    });
  } catch(e) {}
  return defaults;
}

// ============================================================
// 今週月曜日の日付文字列を返す（yyyy-MM-dd）
// ============================================================
function getThisWeekMonday_() {
  const today = new Date();
  const day = today.getDay(); // 0=Sun, 1=Mon, ...
  const diff = day === 0 ? -6 : 1 - day;
  const monday = new Date(today);
  monday.setDate(today.getDate() + diff);
  return Utilities.formatDate(monday, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// ============================================================
// 全担当者の今週行動KPI集計（架電・面談・紹介 の 実績 vs 目標）
// ============================================================
function getWeeklyKPISummaryAll_(weekStart) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const persons = getPersonDetails().filter(function(p) { return p.status === '在籍中'; });
  const goalSheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  const actSheet  = ss.getSheetByName(SHEET_ACTIVITIES);
  const weekDates = getWeekDateRange(weekStart);

  // シートデータをループ外で一括読み込み（担当者数×2回→各1回に削減）
  const gRows = (goalSheet && goalSheet.getLastRow() > 1)
    ? goalSheet.getDataRange().getValues().slice(1) : [];
  const aRows = (actSheet && actSheet.getLastRow() > 1)
    ? actSheet.getDataRange().getValues().slice(1) : [];

  return persons.map(function(p) {
    const person = p.name;

    // 週次目標を取得
    var target = { calls: 0, meetings: 0, referrals: 0 };
    const gRow = gRows.find(function(r) {
      return String(r[0]).trim() === person && toDateStr(r[1]) === weekStart;
    });
    if (gRow) {
      target = { calls: Number(gRow[2])||0, meetings: Number(gRow[3])||0, referrals: Number(gRow[4])||0 };
    }

    // 週次実績を集計
    var actual = { calls: 0, meetings: 0, referrals: 0 };
    aRows.forEach(function(r) {
      if (String(r[0]).trim() === person && weekDates.includes(toDateStr(r[1]))) {
        actual.calls     += Number(r[2]) || 0;
        actual.meetings  += Number(r[3]) || 0;
        actual.referrals += Number(r[4]) || 0;
      }
    });

    return { person: person, role: p.role, target: target, actual: actual };
  });
}

// ============================================================
// 実績データ取得
// ============================================================
function getData() {
  try {
    const sheet = getOrCreateResultSheet();
    const vals = sheet.getDataRange().getValues();
    if (vals.length <= 1) return json({ success:true, data:[], count:0 });
    const headers = vals[0];
    const rows = vals.slice(1).filter(r => r[1]).map(r => {
      const o = {};
      headers.forEach((h,i) => {
        o[h] = r[i] instanceof Date
          ? Utilities.formatDate(r[i],'Asia/Tokyo','yyyy-MM-dd') : r[i];
      });
      return o;
    });
    return json({ success:true, data:rows, count:rows.length });
  } catch(err) {
    return json({ success:false, error:err.message });
  }
}

// ============================================================
// マスタデータ取得（営業・商材・顧客）
// ============================================================
const MASTER_DATA_CACHE_KEY = 'master_data_v1';

function getMaster() {
  // getAllData() などの内部呼び出しでもキャッシュが効くよう関数レベルでキャッシュ
  const cache = CacheService.getScriptCache();
  const cachedStr = cache.get(MASTER_DATA_CACHE_KEY);
  if (cachedStr) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  const personDetails = getPersonDetails();
  const persons = personDetails.filter(p => p.status === '在籍中').map(p => p.name);

  const productSheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const pVals = productSheet.getDataRange().getValues();
  const productDetails = pVals.slice(1).filter(r => String(r[0]).trim()).map(r => {
    const kind = String(r[1] || 'スポット').trim();
    return {
      name:          String(r[0]).trim(),
      kind:          kind,
      months:        kind === 'ストック' ? 12 : 1,
      unitPrice:     Number(r[2]) || 0,
      cost:          Number(r[3]) || 0,
      incentiveRate: Number(r[4]) || 0,  // 0.10% → GASは0.001として読む
      priceType:     String(r[5]).trim() || (Number(r[2]) > 0 ? '固定' : '都度見積もり'),
      bUnitPrice:    Number(r[6]) || 0,  // ← 追加
      bCost:         Number(r[7]) || 0,  // ← 追加
    };
  });
  const products = productDetails.map(p => p.name);

  const customers = getOldCustomers();
  const payload = JSON.stringify({ success: true, persons, personDetails, products, productDetails, customers });
  try { cache.put(MASTER_DATA_CACHE_KEY, payload, 300); } catch(_) {} // 5分キャッシュ
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 設定_会社シート取得
// ============================================================
function getCompanies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('設定_会社');
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1)
    .filter(function(r){ return r[3] === '有効'; })
    .map(function(r){
      return {
        id:       String(r[0]),
        name:     String(r[1]),
        category: String(r[2]),
        status:   String(r[3])
      };
    });
}

// ============================================================
// 営業詳細リスト取得
// ============================================================
function getPersonDetails() {
  const sheet = getOrCreatePersonSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  // A:営業名 B:役職 C:ステータス D:個人コード E:メールアドレス
  const vals = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return vals
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name:   String(r[0]).trim(),
      role:   String(r[1]).trim() || 'スタッフ',
      status: String(r[2]).trim() || '在籍中',
      code:   String(r[3] || '').trim(),
      email:  String(r[4] || '').trim()
    }));
}

// ============================================================
// 顧客DBの担当営業変更
// ============================================================
function changeCompanyPerson(code, newPerson) {
  try {
    if (!code || !newPerson) return json({ success: false, error: 'コードまたは営業名が空です' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
    if (!sheet) return json({ success: false, error: '顧客DBが見つかりません' });

    const histCol = ensureHistoryColumn(sheet);
    const lastRow = sheet.getLastRow();
    const allCodes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    for (let i = 0; i < allCodes.length; i++) {
      if (String(allCodes[i][0]).trim() === code.trim()) {
        const row = i + 2;
        const currentPerson = String(sheet.getRange(row, 2).getValue()).trim();
        if (currentPerson && currentPerson !== newPerson) {
          const histCell = sheet.getRange(row, histCol);
          const existing = String(histCell.getValue()).trim();
          const entry = currentPerson + '(〜' + today + ')';
          histCell.setValue(existing ? existing + ', ' + entry : entry);
        }
        sheet.getRange(row, 2).setValue(newPerson);
        return json({ success: true });
      }
    }
    return json({ success: false, error: '顧客コードが見つかりません: ' + code });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// ============================================================
// 顧客DBに前任者履歴列を追加
// ============================================================
function initCustomerHistoryColumn() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
    if (!sheet) return json({ success: false, error: '顧客DBが見つかりません' });
    const col = ensureHistoryColumn(sheet);
    return json({ success: true, column: col });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

function ensureHistoryColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let histCol = headers.findIndex(h => String(h).trim() === '前任者履歴') + 1;
  if (histCol > 0) return histCol;

  histCol = sheet.getLastColumn() + 1;
  const headerCell = sheet.getRange(1, histCol);
  headerCell.setValue('前任者履歴');
  headerCell.setBackground('#4a4a6a').setFontColor('#fff').setFontWeight('bold');
  sheet.setColumnWidth(histCol, 220);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const bVals = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const histVals = bVals.map(r => {
      const name = String(r[0]).trim();
      return [name ? name + '(〜' + today + '以前)' : ''];
    });
    sheet.getRange(2, histCol, lastRow - 1, 1).setValues(histVals);
  }
  return histCol;
}

// ============================================================
// 顧客一覧（軽量版）
// ============================================================
function getCustomerList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!sheet) return json({ success: true, customers: [] });
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: true, customers: [] });

  const readCols = Math.min(10, sheet.getLastColumn());
  const vals = sheet.getRange(1, 1, lastRow, readCols).getValues();
  const headers = vals[0].map(h => String(h).trim());

  const codeIdx    = headers.indexOf('顧客コード');
  const personIdx  = headers.indexOf('AFC営業担当');
  const companyIdx = headers.indexOf('企業名');
  const prefIdx    = headers.indexOf('都道府県');

  const customers = vals.slice(1)
    .map(r => ({
      code:       String(codeIdx    >= 0 ? r[codeIdx]    : '').trim(),
      person:     String(personIdx  >= 0 ? r[personIdx]  : '').trim(),
      company:    String(companyIdx >= 0 ? r[companyIdx] : '').trim(),
      prefecture: String(prefIdx    >= 0 ? r[prefIdx]    : '').trim()
    }))
    .filter(r => r.company || r.code);

  return json({ success: true, customers });
}

// ============================================================
// 顧客詳細（1件 + 活動履歴 + 案件履歴）
// ============================================================
function getCustomerDetail(code) {
  if (!code) return json({ success: false, error: 'コードが空です' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!sheet) return json({ success: false, error: '顧客DBが見つかりません' });

  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = vals[0].map(h => String(h).trim());
  const codeIdx = headers.indexOf('顧客コード');

  const row = vals.slice(1).find(r => String(r[codeIdx]).trim() === code.trim());
  if (!row) return json({ success: false, error: '見つかりません: ' + code });

  const customer = {};
  headers.forEach((h, i) => { if (h) customer[h] = String(row[i] || '').trim(); });

  const activities = getActivitiesForCompany(customer['企業名']);
  const deals = getDealsForCustomer(code, customer['企業名']);
  const personDetails = getPersonDetails();

  return json({ success: true, customer, activities, deals, personDetails });
}

// 顧客の案件履歴を取得
function getDealsForCustomer(customerId, companyName) {
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const vals = sheet.getRange(1, 1, lastRow, DEAL_HEADERS.length).getValues();
  const headers = vals[0];
  const cidIdx  = headers.indexOf('顧客ID');
  const cmpIdx  = headers.indexOf('会社名');
  const normTarget = normalizeCompany_(companyName);

  return vals.slice(1).filter(r => {
    if (r[0] === '') return false;
    if (customerId && String(r[cidIdx]).trim() === customerId.trim()) return true;
    if (companyName) {
      const normRow = normalizeCompany_(r[cmpIdx]);
      if (normRow && normTarget && (normRow === normTarget || normRow.includes(normTarget) || normTarget.includes(normRow))) return true;
    }
    return false;
  }).map(r => {
    const o = {};
    headers.forEach((h, i) => {
      o[h] = r[i] instanceof Date
        ? Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd') : r[i];
    });
    return o;
  });
}

function getActivitiesForCompany(companyName) {
  if (!companyName) return [];
  const sheet = getOrCreateResultSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const vals = sheet.getRange(1, 1, lastRow, HEADERS.length).getValues();
  const headers = vals[0];
  const companyIdx = headers.indexOf('企業名');
  if (companyIdx < 0) return [];

  const normTarget = normalizeCompany_(companyName);
  if (!normTarget) return [];

  return vals.slice(1)
    .filter(r => r[1])
    .filter(r => {
      const normRow = normalizeCompany_(r[companyIdx]);
      if (!normRow) return false;
      return normRow === normTarget
        || normRow.includes(normTarget)
        || normTarget.includes(normRow);
    })
    .map(r => {
      const o = {};
      headers.forEach((h, i) => {
        o[h] = r[i] instanceof Date
          ? Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd') : r[i];
      });
      return o;
    })
    .sort((a, b) => (String(b['日付']) > String(a['日付']) ? 1 : -1))
    .slice(0, 50);
}

// ============================================================
// 旧・顧客DB取得（getMaster用。新getCustomersはgas_customer_master.jsで定義）
// ============================================================
function getOldCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const vals = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(r => String(r[1]||r[0]).trim())
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(r[i]||'').trim(); });
      return obj;
    });
}

// ============================================================
// 入力フォームHTML（日報）
// ============================================================
function buildFormHtml() {
  const scriptUrl = ScriptApp.getService().getUrl();
  const persons  = getPersonDetails().filter(p => p.status === '在籍中').map(p => p.name);
  const products = getListFromSheet(getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3));
  const personOpts  = persons.map(p  => `<option value="${p}">${p}</option>`).join('');
  const productOpts = products.map(p => `<option value="${p}">${p}</option>`).join('');

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>営業日報入力</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--blue:#2563eb;--blue-bg:#eff6ff;--green:#16a34a;--green-bg:#f0fdf4;
  --red:#dc2626;--red-bg:#fef2f2;--amber:#d97706;--amber-bg:#fffbeb;
  --text:#1a1917;--text2:#6b6860;--text3:#9c9a94;
  --border:#e0ddd4;--bg:#f5f4f1;--surface:#fff;--r:8px;--r-lg:14px}
body{font-family:'Hiragino Sans','Noto Sans JP',sans-serif;background:var(--bg);
  color:var(--text);font-size:15px;min-height:100vh;padding-bottom:80px}
.wrap{max-width:500px;margin:0 auto;padding:16px}
.header{text-align:center;padding:20px 0 18px}
.header-co{font-size:11px;color:var(--text3);letter-spacing:.1em}
.header-title{font-size:21px;font-weight:700;margin-top:4px}
.header-date{font-size:13px;color:var(--text2);margin-top:4px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r-lg);padding:18px;margin-bottom:12px}
.sec{font-size:11px;font-weight:700;color:var(--text3);letter-spacing:.06em;text-transform:uppercase;
  margin-bottom:14px;display:flex;align-items:center;gap:8px}
.sec::after{content:'';flex:1;height:1px;background:var(--border)}
.fg{margin-bottom:12px}.fg:last-child{margin-bottom:0}
.fl{font-size:12px;font-weight:500;color:var(--text2);display:block;margin-bottom:5px}
.fi,.fsel{width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:var(--r);
  font-size:15px;font-family:inherit;background:var(--surface);color:var(--text);
  -webkit-appearance:none;transition:.15s}
.fi:focus,.fsel:focus{outline:none;border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.fsel{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%239c9a94' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 12px center}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
.ki .fl{font-size:11px}
.kii{width:100%;padding:10px;border:1px solid var(--border);border-radius:var(--r);
  font-size:20px;font-weight:700;font-family:inherit;text-align:center;
  background:var(--surface);color:var(--text);-webkit-appearance:none}
.kii:focus{outline:none;border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.apo-chip{background:var(--blue-bg);color:var(--blue);padding:9px 14px;
  border-radius:var(--r);font-size:13px;font-weight:500;text-align:center;margin-top:10px}
.submit-btn{width:100%;padding:17px;background:var(--blue);color:#fff;border:none;
  border-radius:var(--r-lg);font-size:17px;font-weight:700;cursor:pointer;
  font-family:inherit;transition:.15s;margin-top:4px}
.submit-btn:active{transform:scale(.98);background:#1d4ed8}
.submit-btn:disabled{background:var(--border);color:var(--text3);cursor:not-allowed}
.err{background:var(--red-bg);color:var(--red);padding:10px 14px;border-radius:var(--r);
  font-size:13px;margin-bottom:10px;display:none}
.success{display:none;text-align:center;padding:48px 20px}
.success-icon{font-size:72px;margin-bottom:16px}
.success-title{font-size:22px;font-weight:700;margin-bottom:8px}
.success-sub{font-size:14px;color:var(--text2);margin-bottom:28px}
.again-btn{padding:13px 32px;background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r);font-size:14px;font-weight:500;cursor:pointer;font-family:inherit}
hr{border:none;border-top:1px solid var(--border);margin:14px 0}
.add-row{display:flex;gap:8px;margin-top:8px}
.add-input{flex:1;padding:8px 10px;border:1px solid var(--border);border-radius:var(--r);
  font-size:13px;font-family:inherit;background:var(--surface);color:var(--text)}
.add-input:focus{outline:none;border-color:var(--blue)}
.add-btn{padding:8px 14px;background:var(--blue);color:#fff;border:none;border-radius:var(--r);
  font-size:13px;font-weight:500;cursor:pointer;font-family:inherit;white-space:nowrap}
.add-btn:disabled{background:var(--border);color:var(--text3);cursor:not-allowed}
.add-msg{font-size:12px;margin-top:6px;min-height:16px}
.add-msg.ok{color:var(--green)}.add-msg.ng{color:var(--red)}
details summary{font-size:12px;color:var(--text3);cursor:pointer;list-style:none;padding:6px 0}
details summary::before{content:'＋ '}
details[open] summary::before{content:'－ '}
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <div class="header-co">エー・ファクトリー株式会社</div>
    <div class="header-title">営業日報入力</div>
    <div class="header-date" id="hdate"></div>
  </div>

  <div id="form-area">
    <div class="err" id="err"></div>

    <div class="card">
      <div class="sec">基本情報</div>
      <div class="fg">
        <label class="fl">営業担当</label>
        <select class="fsel" id="person">
          <option value="">選択してください</option>
          ${personOpts}
        </select>
        <details style="margin-top:6px">
          <summary>営業担当を追加する</summary>
          <div class="add-row">
            <input class="add-input" id="add-person-input" placeholder="氏名を入力（例：山田 花子）">
            <button class="add-btn" onclick="addMaster('person')">追加</button>
          </div>
          <div class="add-msg" id="add-person-msg"></div>
        </details>
      </div>
      <div class="fg">
        <label class="fl">日付</label>
        <input class="fi" type="date" id="date">
      </div>
      <div class="fg">
        <label class="fl">商材</label>
        <select class="fsel" id="product">
          <option value="">選択してください</option>
          ${productOpts}
        </select>
        <div style="font-size:11px;color:var(--text3);margin-top:5px;padding:6px 10px;border-radius:5px">
          ※ 商材の追加・変更は管理者にご連絡ください
        </div>
      </div>
    </div>

    <div class="card">
      <div class="sec">A. アポ取り</div>
      <div style="font-size:12px;font-weight:500;color:var(--text2);margin-bottom:10px">面談アポ</div>
      <div class="g2" style="margin-bottom:12px">
        <div class="ki"><label class="fl">自分でアポ</label><input class="kii" type="number" id="selfApo" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
        <div class="ki"><label class="fl">テレアポ経由</label><input class="kii" type="number" id="telApo" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
        <div class="ki"><label class="fl">紹介会社</label><input class="kii" type="number" id="refApo" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
        <div class="ki"><label class="fl">代理店紹介</label><input class="kii" type="number" id="agentRef" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
      </div>
      <hr>
      <div style="font-size:12px;font-weight:500;color:var(--text2);margin-bottom:10px">代理店開拓アポ</div>
      <div class="g2">
        <div class="ki"><label class="fl">ビジェントから</label><input class="kii" type="number" id="bizent" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
        <div class="ki"><label class="fl">自己開拓</label><input class="kii" type="number" id="selfAgent" value="0" min="0" inputmode="numeric" oninput="calcApo()"></div>
      </div>
      <div class="apo-chip">今日のアポ合計：<strong id="apototal">0</strong> 件</div>
    </div>

    <div class="card">
      <div class="sec">B. 面談プロセス</div>
      <div class="g3">
        <div class="ki"><label class="fl">ヒアリング</label><input class="kii" type="number" id="hearing" value="0" min="0" inputmode="numeric"></div>
        <div class="ki"><label class="fl">提案</label><input class="kii" type="number" id="proposal" value="0" min="0" inputmode="numeric"></div>
        <div class="ki"><label class="fl">成約</label><input class="kii" type="number" id="closing" value="0" min="0" inputmode="numeric"></div>
      </div>
    </div>

    <div class="card">
      <div class="sec">C. 売上実績</div>
      <div class="ki">
        <label class="fl">売上金額（万円）</label>
        <input class="kii" type="number" id="sales" value="0" min="0" inputmode="numeric" style="font-size:26px;padding:14px">
      </div>
    </div>

    <button class="submit-btn" id="sbtn" onclick="submit()">送信する</button>
  </div>

  <div class="success" id="success">
    <div class="success-icon">✅</div>
    <div class="success-title">送信完了！</div>
    <div class="success-sub" id="smsg"></div>
    <button class="again-btn" onclick="reset()">続けて入力する</button>
  </div>
</div>

<script>
const URL = '${scriptUrl}';
const NUM_IDS = ['selfApo','telApo','refApo','agentRef','bizent','selfAgent','hearing','proposal','closing','sales'];

function init() {
  const d = new Date();
  document.getElementById('hdate').textContent =
    d.toLocaleDateString('ja-JP',{year:'numeric',month:'long',day:'numeric',weekday:'short'});
  document.getElementById('date').value = d.toISOString().slice(0,10);
  const lp = localStorage.getItem('last_person');
  if (lp) document.getElementById('person').value = lp;
}

function calcApo() {
  const t = ['selfApo','telApo','refApo','agentRef','bizent','selfAgent']
    .reduce((s,id) => s + (parseInt(document.getElementById(id).value)||0), 0);
  document.getElementById('apototal').textContent = t;
}

async function addMaster(type) {
  const inputId = type === 'person' ? 'add-person-input' : 'add-product-input';
  const msgId   = type === 'person' ? 'add-person-msg'  : 'add-product-msg';
  const selId   = type === 'person' ? 'person' : 'product';
  const action  = type === 'person' ? 'addPerson' : 'addProduct';
  const name    = document.getElementById(inputId).value.trim();
  const msgEl   = document.getElementById(msgId);
  if (!name) { showAddMsg(msgEl, '名前を入力してください', false); return; }
  const btn = document.querySelector(\`[onclick="addMaster('\${type}')"]\`);
  btn.disabled = true; btn.textContent = '追加中...';
  try {
    await fetch(URL, {
      method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ action, name })
    });
    const sel = document.getElementById(selId);
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    sel.appendChild(opt);
    sel.value = name;
    document.getElementById(inputId).value = '';
    showAddMsg(msgEl, '✅ ' + name + ' を追加しました', true);
  } catch(e) {
    showAddMsg(msgEl, '追加に失敗しました', false);
  }
  btn.disabled = false; btn.textContent = '追加';
}

function showAddMsg(el, msg, ok) {
  el.textContent = msg;
  el.className = 'add-msg ' + (ok ? 'ok' : 'ng');
  setTimeout(() => { el.textContent = ''; el.className = 'add-msg'; }, 3000);
}

async function submit() {
  const person  = document.getElementById('person').value;
  const date    = document.getElementById('date').value;
  const product = document.getElementById('product').value;
  const errEl   = document.getElementById('err');
  errEl.style.display = 'none';
  if (!person)  { showErr('営業担当を選択してください'); return; }
  if (!date)    { showErr('日付を入力してください'); return; }
  if (!product) { showErr('商材を選択してください'); return; }
  const btn = document.getElementById('sbtn');
  btn.disabled = true; btn.textContent = '送信中...';
  const data = { date, person, product,
    selfApo:v('selfApo'), telApo:v('telApo'), refApo:v('refApo'), agentRef:v('agentRef'),
    bizent:v('bizent'), selfAgent:v('selfAgent'),
    hearing:v('hearing'), proposal:v('proposal'), closing:v('closing'), sales:v('sales') };
  try {
    await fetch(URL, { method:'POST', mode:'no-cors',
      headers:{'Content-Type':'application/json'}, body:JSON.stringify(data) });
    localStorage.setItem('last_person', person);
    document.getElementById('smsg').textContent =
      person + 'さん、' + date + ' の日報を送信しました。お疲れ様でした！';
    document.getElementById('form-area').style.display = 'none';
    document.getElementById('success').style.display = 'block';
  } catch(e) {
    showErr('送信に失敗しました。ネットワークを確認してください。');
    btn.disabled = false; btn.textContent = '送信する';
  }
}

function v(id) { return parseInt(document.getElementById(id).value) || 0; }
function showErr(msg) {
  const e = document.getElementById('err');
  e.textContent = msg; e.style.display = 'block';
  window.scrollTo({top:0,behavior:'smooth'});
}
function reset() {
  document.getElementById('success').style.display = 'none';
  document.getElementById('form-area').style.display = 'block';
  NUM_IDS.forEach(id => document.getElementById(id).value = 0);
  calcApo();
  document.getElementById('sbtn').disabled = false;
  document.getElementById('sbtn').textContent = '送信する';
  document.getElementById('date').value = new Date().toISOString().slice(0,10);
}

init();
</script>
</body>
</html>`;
}

// ============================================================
// ユーティリティ
// ============================================================
function n(v) { return Number(v) || 0; }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getListFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals.map(r => String(r[0]).trim()).filter(v => v !== '');
}

function getOrCreateResultSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RESULT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESULT);
  }
  // ヘッダーが未設定（1行目が空）の場合は自動で追加
  const firstCell = sheet.getRange(1,1).getValue();
  if (!firstCell) {
    const hr = sheet.getRange(1,1,1,HEADERS.length);
    hr.setValues([HEADERS]);
    hr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,140); sheet.setColumnWidth(2,100);
    sheet.setColumnWidth(3,100); sheet.setColumnWidth(4,180);
    sheet.setColumnWidth(15,200); sheet.setColumnWidth(16,80);
    sheet.setColumnWidth(17,160); sheet.setColumnWidth(18,200);
  }
  return sheet;
}

// 営業実績シートのヘッダーを手動でリセットする（GASエディタから実行）
function initResultSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RESULT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESULT);
  }
  const hr = sheet.getRange(1,1,1,HEADERS.length);
  hr.setValues([HEADERS]);
  hr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,140); sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,100); sheet.setColumnWidth(4,180);
  sheet.setColumnWidth(15,200); sheet.setColumnWidth(16,80);
  sheet.setColumnWidth(17,160); sheet.setColumnWidth(18,200);
  return '営業実績シートのヘッダーを設定しました';
}

function getOrCreateMasterSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const hr = sheet.getRange(1,1,1,headers.length);
    hr.setValues([headers]);
    hr.setBackground('#2d6a4f').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
  }
  return sheet;
}

// 営業専用シート
function getOrCreatePersonSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PERSONS);
  if (!sheet) sheet = ss.insertSheet(SHEET_PERSONS);
  const h1 = String(sheet.getRange(1,1).getValue()).trim();
  if (!h1) {
    const headers = ['営業名','役職','ステータス','個人コード'];
    const hr = sheet.getRange(1,1,1,4);
    hr.setValues([headers]);
    hr.setBackground('#2d6a4f').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 120);
  }
  return sheet;
}

// 案件マスタシート
function getOrCreateDealSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DEALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DEALS);
    const hr = sheet.getRange(1, 1, 1, DEAL_HEADERS.length);
    hr.setValues([DEAL_HEADERS]);
    hr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,  200); // 案件ID
    sheet.setColumnWidth(2,  100); // 登録日
    sheet.setColumnWidth(3,  100); // 担当者
    sheet.setColumnWidth(4,  110); // 顧客ID
    sheet.setColumnWidth(5,  180); // 会社名
    sheet.setColumnWidth(6,  200); // 商材名
    sheet.setColumnWidth(7,  120); // フェーズ
    sheet.setColumnWidth(8,   80); // 確度ランク
    sheet.setColumnWidth(9,  110); // 売上（単価）
    sheet.setColumnWidth(10, 110); // 費用（単価）
    sheet.setColumnWidth(11,  70); // コース数
    sheet.setColumnWidth(12,  60); // 件数
    sheet.setColumnWidth(13,  60); // 月数
    sheet.setColumnWidth(14, 120); // 売上予定額
    sheet.setColumnWidth(15, 110); // 費用（合計）
    sheet.setColumnWidth(16, 110); // 粗利
    sheet.setColumnWidth(17, 130); // インセンティブ
    sheet.setColumnWidth(18, 110); // 売上予定月
    sheet.setColumnWidth(19, 110); // 入金ステータス
    sheet.setColumnWidth(20, 110); // 入金確認日
    sheet.setColumnWidth(21, 250); // メモ
    sheet.setColumnWidth(22, 110); // 引継担当者
    sheet.setColumnWidth(23, 100); // 引継日
  }
  return sheet;
}


// ============================================================
// 全データ消去（初期化用）
// ============================================================
function clearAll() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_PERSONS);
    if (pSheet && pSheet.getLastRow() > 1) pSheet.deleteRows(2, pSheet.getLastRow() - 1);
    const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);
    if (prodSheet && prodSheet.getLastRow() > 1) prodSheet.deleteRows(2, prodSheet.getLastRow() - 1);
    const gSheet = ss.getSheetByName(SHEET_GOALS);
    if (gSheet) { gSheet.getRange(1,1).clearContent(); gSheet.getRange(1,2).clearContent(); }
    const rSheet = ss.getSheetByName(SHEET_RESULT);
    if (rSheet && rSheet.getLastRow() > 1) rSheet.deleteRows(2, rSheet.getLastRow() - 1);
    const dSheet = ss.getSheetByName(SHEET_DEALS);
    if (dSheet && dSheet.getLastRow() > 1) dSheet.deleteRows(2, dSheet.getLastRow() - 1);
    return json({ success: true });
  } catch(err) {
    return json({ success: false, error: err.message });
  }
}

// ============================================================
// スプレッドシートを開いたときのメニュー
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 顧客管理')
    .addItem('顧客コードを生成する（初回のみ）', 'insertCustomerCodes')
    .addItem('【初回のみ】案件マスタのヘッダーをリセットする', 'resetDealSheetHeaders')
    .addToUi();
}

// ============================================================
// 案件マスタのヘッダーを新形式にリセット（データはそのまま残す）
// ============================================================
function resetDealSheetHeaders() {
  const ui = SpreadsheetApp.getUi();

  // 1回目：警告
  const confirm1 = ui.alert(
    '⚠️ 案件マスタの全データが削除されます',
    'この操作は取り消せません。案件マスタシートを削除して新しい列構造で作り直します。\n\n本当に実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm1 !== ui.Button.YES) { ui.alert('キャンセルしました。'); return; }

  // 2回目：入力確認
  const confirm2 = ui.prompt(
    '最終確認',
    '削除を実行するには「DELETE」と入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm2.getSelectedButton() !== ui.Button.OK) { ui.alert('キャンセルしました。'); return; }
  if (confirm2.getResponseText().trim() !== 'DELETE') { ui.alert('入力が違います。キャンセルしました。'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEALS);
  if (sheet) ss.deleteSheet(sheet);
  getOrCreateDealSheet();
  ui.alert('完了しました。案件マスタを新しい列構造で作り直しました。');
}

// ============================================================
// 週次活動データ取得
// ============================================================
function getWeeklyData(person, weekStart) {
  if (!person || !weekStart) return { success: false, error: 'person と weekStart は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 週次目標を取得
  const goalSheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  let target = null;
  if (goalSheet && goalSheet.getLastRow() > 1) {
    const rows = goalSheet.getDataRange().getValues().slice(1);
    const goalRow = rows.find(r => String(r[0]) === person && toDateStr(r[1]) === weekStart);
    if (goalRow) target = { calls: Number(goalRow[2]) || 0, meetings: Number(goalRow[3]) || 0, referrals: Number(goalRow[4]) || 0 };
  }

  // 当週の日次活動を取得（月〜金）
  const actSheet = ss.getSheetByName(SHEET_ACTIVITIES);
  const daily = [];
  if (actSheet && actSheet.getLastRow() > 1) {
    const rows = actSheet.getDataRange().getValues().slice(1);
    const weekDates = getWeekDateRange(weekStart);
    rows.forEach(r => {
      const rowPerson = String(r[0]);
      const rowDate   = toDateStr(r[1]);
      if (rowPerson === person && weekDates.includes(rowDate)) {
        daily.push({ date: rowDate, calls: Number(r[2]) || 0, meetings: Number(r[3]) || 0, referrals: Number(r[4]) || 0 });
      }
    });
  }

  return { success: true, target, daily };
}

function getWeekDateRange(mondayStr) {
  const dates = [];
  const base = new Date(mondayStr);
  for (let i = 0; i < 5; i++) {
    const d = new Date(base);
    d.setDate(base.getDate() + i);
    dates.push(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'));
  }
  return dates;
}

// ============================================================
// 日次活動記録（addActivity）
// ============================================================
function addActivity(d) {
  const person = String(d.person || '').trim();
  const date   = String(d.date   || '').trim();
  const calls    = Number(d.calls)    || 0;
  const meetings = Number(d.meetings) || 0;
  const referrals = d.referrals || 0;
  const comment   = d.comment   || '';
  if (!person || !date) return { success: false, error: 'person と date は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVITIES);
    sheet.appendRow(['担当者', '日付', '架電数', '面談数', '紹介数', 'コメント']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 同日のレコードがあれば更新、なければ追記
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === person && String(vals[i][1]).slice(0,10) === date) {
        sheet.getRange(i + 2, 3, 1, 4).setValues([[calls, meetings, referrals, comment]]);
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([person, date, calls, meetings, referrals, comment]);
  return { success: true, updated: false };
}

// ============================================================
// 週次目標保存（setWeeklyTarget）
// ============================================================
function setWeeklyTarget(d) {
  const person    = String(d.person    || '').trim();
  const weekStart = String(d.weekStart || '').trim();
  const callTarget    = Number(d.callTarget)    || 0;
  const meetingTarget = Number(d.meetingTarget) || 0;
  if (!person || !weekStart) return { success: false, error: 'person と weekStart は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY_GOALS);
    sheet.appendRow(['担当者', '週開始日', '架電目標', '面談目標']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 同担当者・同週があれば更新
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === person && String(vals[i][1]).slice(0,10) === weekStart) {
        sheet.getRange(i + 2, 3, 1, 2).setValues([[callTarget, meetingTarget]]);
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([person, weekStart, callTarget, meetingTarget]);
  return { success: true, updated: false };
}

// ============================================================
// 月次KPI目標保存（setMonthlyKPITarget）
// ============================================================
function setMonthlyKPITarget(d) {
  const person  = String(d.person || '').trim();
  const month   = String(d.month  || '').trim(); // "2026-05"
  const monthly = d.monthly || { calls: 0, meetings: 0, referrals: 0 };
  const weeks   = d.weeks   || [];
  if (!person || !month) return { success: false, error: 'person と month は必須です' };

  // Script Properties に月次データを保存
  const props = PropertiesService.getScriptProperties();
  props.setProperty('KPI_' + person + '_' + month, JSON.stringify({ monthly, weeks }));

  // SHEET_WEEKLY_GOALS に週別目標を書き込む
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY_GOALS);
    sheet.appendRow(['担当者', '週開始日', '架電目標', '面談目標', '紹介目標']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 既存の同一担当者・同月データを全削除（Date型バグ対策・重複防止）
  const all = sheet.getDataRange().getValues();
  for (let i = all.length - 1; i >= 1; i--) {
    if (String(all[i][0]).trim() === person && toDateStr(all[i][1]).slice(0, 7) === month) {
      sheet.deleteRow(i + 1);
    }
  }

  // 週別目標を書き込む
  weeks.forEach(w => {
    const weekStart = String(w.weekStart || '').slice(0, 10);
    if (!weekStart) return;
    sheet.appendRow([person, weekStart, w.calls || 0, w.meetings || 0, w.referrals || 0]);
  });

  return { success: true };
}

function toDateStr(v) {
  if (v && typeof v.getFullYear === 'function') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).slice(0, 10);
}

// ============================================================
// カレンダーイベント取得（アポ・架電・面談系）
// ============================================================
function getCalendarEventsForWeek(personName, weekStart) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PERSONS);
    if (!sheet) return [];

    const rows = sheet.getDataRange().getValues().slice(1);
    const personRow = rows.find(r => String(r[0]).trim() === personName);
    if (!personRow || !personRow[4]) return []; // E列 = メールアドレス

    const email = String(personRow[4]).trim();
    const cal   = CalendarApp.getCalendarById(email);
    if (!cal) return [];

    const startDate = new Date(weekStart);
    const endDate   = new Date(weekStart);
    endDate.setDate(endDate.getDate() + 5); // 月〜金

    const keywords = /アポ|架電|商談|面談|訪問|MTG|ミーティング|打ち合わせ/i;

    return cal.getEvents(startDate, endDate)
      .filter(e => keywords.test(e.getTitle()))
      .map(e => ({
        title: e.getTitle(),
        start: Utilities.formatDate(e.getStartTime(), 'Asia/Tokyo', 'M/d(E) HH:mm'),
        end:   Utilities.formatDate(e.getEndTime(),   'Asia/Tokyo', 'HH:mm'),
      }));
  } catch(e) {
    return [];
  }
}

// ============================================================
// Gemini 週次アドバイス
// ============================================================
function getWeeklyAdvice(person, weekStart) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, error: 'GASにAPIキーが未設定です' };

  const wdata = getWeeklyData(person, weekStart);
  const target = wdata.target || { calls: 0, meetings: 0, referrals: 0 };
  const daily  = wdata.daily  || [];
  const totalCalls     = daily.reduce((s, d) => s + d.calls,     0);
  const totalMeetings  = daily.reduce((s, d) => s + d.meetings,  0);
  const totalReferrals = daily.reduce((s, d) => s + d.referrals, 0);

  // カレンダーイベント取得
  const calEvents = getCalendarEventsForWeek(person, weekStart);
  const calSection = calEvents.length > 0
    ? `\nGoogleカレンダーの予定（アポ・架電・面談系）：\n` +
      calEvents.map(e => `  - ${e.start}〜${e.end}：${e.title}`).join('\n')
    : '\nGoogleカレンダー：取得できませんでした（データなし）';

  const prompt = `あなたは営業コーチです。以下の週次データを元に、担当者「${person}」へ実践的なアドバイスを日本語300字以内で提供してください。

今週（${weekStart}〜）の実績：
- 架電数：${totalCalls} / 目標 ${target.calls} 件
- 有効面談数：${totalMeetings} / 目標 ${target.meetings} 件
- 紹介数：${totalReferrals} / 目標 ${target.referrals} 件
${calSection}

カレンダーに予定がある場合はその内容も踏まえ、前向きで具体的なアドバイスをMarkdown形式で返してください。`;

  try {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      { method: 'POST', headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.7, maxOutputTokens: 512 } }),
        muteHttpExceptions: true }
    );
    const body = JSON.parse(res.getContentText());
    const advice = body?.candidates?.[0]?.content?.parts?.[0]?.text || 'アドバイスを取得できませんでした';
    return { success: true, advice };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// 顧客コード生成
// ============================================================
function insertCustomerCodes() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!sheet) { SpreadsheetApp.getUi().alert('「顧客DB」シートが見つかりません。'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { SpreadsheetApp.getUi().alert('データがありません。'); return; }

  const headerA = String(sheet.getRange(1, 1).getValue()).trim();
  if (headerA === '顧客コード') {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert('顧客コード列はすでに存在します。', '空欄のセルだけコードを追加しますか？', ui.ButtonSet.YES_NO);
    if (result !== ui.Button.YES) return;
    fillMissingCodes(sheet, lastRow);
    ui.alert('完了しました。');
    return;
  }

  sheet.insertColumnBefore(1);
  const headerCell = sheet.getRange(1, 1);
  headerCell.setValue('顧客コード');
  headerCell.setBackground('#2d6a4f').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 110);
  fillMissingCodes(sheet, lastRow);
  SpreadsheetApp.getUi().alert('完了しました！顧客コードをA列に追加しました。\n形式：AF-XXXXX');
}

function fillMissingCodes(sheet, lastRow) {
  const codeRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const existing  = codeRange.getValues();
  const updates   = existing.map(row => {
    const val = String(row[0]).trim();
    if (val && val !== '' && val !== '0') return [val];
    return [generateCustomerCode()];
  });
  codeRange.setValues(updates);
}

function generateCustomerCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = 'AF-';
  for (let i = 0; i < 5; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  return code;
}

// ============================================================
// 商材一括登録（初回のみ実行）
// ============================================================
function initProducts() {
  // [name, unitPrice, months, cost, dealType, priceType, incentiveType, incentiveTarget]
  const products = [
    ['リスキリング研修_直販',    200000,  1,       0, '通常',          '固定',         '10%単発',  '○'],
    ['リスキリング研修_代理店',  200000,  1,  150000, '通常',          '固定',         '10%単発',  '○'],
    ['HubCast_直販',             35000, 12,   13000, '通常',          '固定',         '1ヶ月分',  '○'],
    ['HubCast_代理店',           35000, 12,   27000, '通常',          '固定',         '10%継続',  '○'],
    ['IT導入補助金',           1000000,  1,       0, '助成金・補助金', '固定',         '0.1%',     '○'],
    ['社長紹介_直販',           600000,  1,       0, '通常',          '固定',         '10%単発',  '○'],
    ['社長紹介_代理店',         600000,  1,  200000, '通常',          '固定',         '10%単発',  '○'],
    ['健康診断事務代理',              0,  1,       0, '通常',          '都度見積もり', '10%単発',  '○'],
    ['保健師電話健康相談',            0,  1,       0, '通常',          '都度見積もり', '10%単発',  '○'],
    ['社労士案件',                    0,  1,       0, '通常',          '都度見積もり', '10%単発',  '○'],
  ];
  products.forEach(([name, unitPrice, months, cost, dealType, priceType, incentiveType, incentiveTarget]) => {
    addProductToSheet(name, unitPrice, months, cost, dealType, priceType, incentiveType, incentiveTarget);
  });
  Logger.log('商材登録完了: ' + products.length + '件');
}

// ============================================================
// 一括インポート（インポートテンプレートから案件マスタへ）
// ============================================================
// 使い方：
//   1. スプレッドシートに「インポート用」シートを作成
//   2. 1行目はヘッダー行（内容は何でもOK）、2行目以降にデータを貼り付け
//   3. 列順（13列）：
//      A:担当者 B:会社名 C:確度ランク D:売上予定月
//      E:売上(円) F:費用(円) G:粗利(円) H:商材名 I:メモ
//      J:計上会社 K:B売上単価(円) L:B費用単価(円) M:B件数
//      ※ J〜M列は省略可（空欄なら空文字・0として登録）
//   4. この関数をGASエディタから実行
// ============================================================
function importFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('インポート用');
  if (!src) { Logger.log('「インポート用」シートが見つかりません'); return; }

  const vals = src.getDataRange().getValues();
  if (vals.length <= 1) { Logger.log('データがありません'); return; }

  const dest = ss.getSheetByName(SHEET_DEALS);

  // 営業名→個人コードの変換マップを構築
  const personNameToCode = {};
  getPersonDetails().forEach(p => { if (p.name && p.code) personNameToCode[p.name] = p.code; });

  const VALID_RANKS = ['売上', '決定', 'A', 'B', 'C', '失注'];
  let added = 0, skipped = 0, errors = [];

  // 重複チェック用に既存データを一度だけ読み込む
  const iPersonIdx  = DEAL_HEADERS.indexOf('担当者');
  const iCompanyIdx = DEAL_HEADERS.indexOf('会社名');
  const iMonthIdx   = DEAL_HEADERS.indexOf('売上予定月');
  const existingRows = dest.getLastRow() > 1
    ? dest.getDataRange().getValues().slice(1)
    : [];

  vals.slice(1).forEach((r, i) => {
    const rowNum       = i + 2;
    const personName   = String(r[0]||'').trim();
    // 営業名→個人コードに変換（コードが未設定なら名前をそのまま保存）
    const person       = personNameToCode[personName] || personName;
    const company      = String(r[1]||'').trim();
    const rank         = String(r[2]||'').trim();
    const expMonth     = String(r[3]||'').trim();
    const sales        = Number(r[4])||0;
    const cost         = Number(r[5])||0;
    const gp           = Number(r[6])||0;
    const product      = String(r[7]||'').trim();
    const memo         = String(r[8]||'').trim();
    const billingCo    = String(r[9]||'').trim();
    const bUnitSales   = Number(r[10])||0;
    const bUnitCost    = Number(r[11])||0;
    const bQty         = Number(r[12])||0;

    // バリデーション
    if (!personName || !company || !expMonth) {
      errors.push('行'+rowNum+': 営業担当・会社名・売上予定月は必須');
      skipped++; return;
    }
    if (!VALID_RANKS.includes(rank)) {
      errors.push('行'+rowNum+': 確度ランク「'+rank+'」が無効（売上/決定/A/B/C/失注）');
      skipped++; return;
    }
    if (!/^\d{4}-\d{2}$/.test(expMonth)) {
      errors.push('行'+rowNum+': 売上予定月の形式が無効（例: 2026-04）');
      skipped++; return;
    }

    // 重複チェック（同営業コード・同会社・同月）
    const isDuplicate = existingRows.some(dr =>
      String(dr[iPersonIdx])===person &&
      String(dr[iCompanyIdx])===company &&
      String(dr[iMonthIdx]).slice(0,7)===expMonth
    );
    if (isDuplicate) {
      errors.push('行'+rowNum+': 重複スキップ（'+personName+'/'+company+'/'+expMonth+'）');
      skipped++; return;
    }

    // 案件ID生成
    const now = new Date();
    const id = 'DEAL-IMP-' + Utilities.formatDate(now,'Asia/Tokyo','yyyyMMddHHmmss') + '-' + (added+1);
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    const phase = (rank === '売上' || rank === '決定') ? '完了' : 'ヒアリング中';
    const payStatus = rank === '売上' ? '入金済み' : '未入金';

    const rowMap = {
      '案件ID': id, '登録日': today, '担当者': person, '顧客ID': '',
      '会社名': company, '商材名': product, 'フェーズ': phase, '確度ランク': rank,
      '売上（単価）': sales, '費用（単価）': cost, 'コース数': 1, '件数': 1, '月数': 1,
      '売上予定額': sales, '費用（合計）': cost, '粗利': gp,
      'インセンティブ': 0, '売上予定月': expMonth,
      '入金ステータス': payStatus, '入金確認日': '', 'メモ': memo,
      '引継担当者': '', '引継日': '', '担当者コード': '', '理由': '', '最終更新日': today,
      '計上会社': billingCo, 'B売上単価': bUnitSales, 'B費用単価': bUnitCost, 'B件数': bQty
    };
    const newRow = DEAL_HEADERS.map(h => rowMap[h] !== undefined ? rowMap[h] : '');
    dest.appendRow(newRow);
    existingRows.push(newRow); // 同一インポート内での重複防止
    added++;
  });

  Logger.log('【インポート完了】 追加:'+added+'件 / スキップ:'+skipped+'件');
  if (errors.length) Logger.log('【エラー詳細】\n'+errors.join('\n'));
}
