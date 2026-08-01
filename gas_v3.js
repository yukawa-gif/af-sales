// ============================================================
// エー・ファクトリー 営業管理システム - Apps Script v3
// 営業・商材・案件をスプレッドシートで管理するバージョン
// ============================================================
// 【シート構成】
//   「設定_営業」シート  ：個人コード・営業名・役職・ステータス
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
const SHEET_COST_SCHEDULE = '費用計上明細';
const SHEET_PRICE_HISTORY = '商材価格変更履歴';

// 商材価格変更履歴の列定義
const PRICE_HISTORY_HEADERS = [
  '変更日時','商材コード','商材名','変更項目','旧値','新値',
  '反映件数','スキップ件数','反映した案件ID','実行者'
];

// 費用計上明細の列定義（案件ごとに費用を複数月へ分散計上するための明細）
const COST_SCHEDULE_HEADERS = ['案件ID', '計上月', '金額'];

const HEADERS = [
  '送信日時','日付','個人コード','営業名','商材',
  '自分でアポ','テレアポ経由','紹介会社','代理店紹介',
  'ビジェント','代理店開拓',
  'ヒアリング','提案','クロージング','売上(万円)',
  'メモ','ステータス','企業名','プロジェクト名','費用(万円)','粗利(万円)'
];

// 商材マスタの列定義
// A:商材コード B:商材名 C:種別 D:売上単価 E:費用 F:インセンティブ率
// G:価格タイプ H:B単価（円） I:B費用（円） J:インセンティブ固定額（円）
// J列に値がある商材は、インセンティブ率（F列）による粗利×率の計算ではなく
// 「固定額 × 件数」で計算する（例: HubCastは1件成約につき21,000円固定）
const PRODUCT_HEADERS_V3 = [
  '商材コード','商材名','種別','売上単価（円）','費用（円）','インセンティブ率',
  '価格タイプ','B単価（円）','B費用（円）','インセンティブ固定額（円）'
];

// 案件マスタの列定義
// C:個人コード D:営業名 は設定_営業のA列・B列に対応
const DEAL_HEADERS = [
  '案件ID','登録日','個人コード','営業名','顧客ID','会社名','商材名',
  'フェーズ','確度ランク',
  '売上（単価）','費用（単価）','コース数','件数','月数',
  '売上予定額','費用（合計）','粗利',
  'インセンティブ','売上予定月','入金ステータス','入金確認日',
  'メモ','引継営業名','引継日','理由','最終更新日',
  '計上会社','B売上単価','B費用単価','B件数',
  '商材コード',  // AE: index 30
  '継続課金','継続終了月',  // AF/AG: 顧問契約など、止めるまで毎月自動計上する案件のフラグと終了月
  'インセンティブ計上済み', // AH: index 33。ストック案件は契約全期間分を初回の決定→売上確定時に一括計上するため、
                            // 二重計上・多重計上を防ぐための一度払ったら立てるフラグ（分割・繰越・編集時に参照）
  '次回アクション日'        // AI: index 34。営業が次にアクションすべき日（AI提案 or 手動入力、yyyy-MM-dd）
];

// ============================================================
// 会社名正規化（法人格・空白除去・小文字統一）
// "株式会社loty" と "loty" を同一とみなすための共通ヘルパー
// ============================================================
// ============================================================
// E2Eシステムテストが挿入する検証用データ（案件ID: TEST-001〜010、
// 担当者名: テスト担当者）を判定する。月別推移・KPI・インセンティブ等
// 本番集計に使うエンドポイント（getAllData）はこれで必ず除外すること。
// ============================================================
function isTestDeal_(d) {
  var id = String((d && d['案件ID']) || '').trim();
  if (id.indexOf('TEST-') === 0) return true;
  var person = String((d && (d['担当者'] || d['営業名'])) || '').trim();
  return person === 'テスト担当者';
}

function normalizeCompany_(s) {
  return String(s || '').trim()
    .replace(/株式会社|有限会社|合同会社|一般社団法人|一般財団法人|特定非営利活動法人/g, '')
    .replace(/[\s　（()）・]/g, '')
    // 全角英数字 → 半角（例: "ＡＢＣ" → "abc"）
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .toLowerCase();
}

// ============================================================
// GETリクエスト
// ============================================================
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'form';
  if (mode === 'all')       return getAllData(e && e.parameter && e.parameter.fresh === '1');
  if (mode === 'ai')        return getAIAdvice(e && e.parameter);
  if (mode === 'data')      return getData();
  if (mode === 'master') {
    var masterCache = null;
    try { masterCache = CacheService.getScriptCache(); } catch(e) {}
    if (masterCache) {
      try {
        var masterHit = masterCache.get(GAS_MASTER_CACHE_KEY);
        if (masterHit) {
          return ContentService.createTextOutput(masterHit)
            .setMimeType(ContentService.MimeType.JSON);
        }
      } catch(e) {}
    }
    var m = getMaster();
    var mObj = JSON.parse(m.getContent());
    mObj.topProducts = getTopProducts();
    mObj.companies   = getCompanies();
    var masterPayload = JSON.stringify(mObj);
    if (masterCache) {
      try { masterCache.put(GAS_MASTER_CACHE_KEY, masterPayload, 5 * 60); } catch(e) {}
    }
    return ContentService
      .createTextOutput(masterPayload)
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'goals')           return getGoals(e && e.parameter && e.parameter.fy);
  if (mode === 'goals_history')   return getGoalsHistory(e && e.parameter && e.parameter.fy);
  if (mode === 'generateTestData') return withLock(() => generateTestData());
  if (mode === 'clearTestData')    return withLock(() => clearTestData());
  if (mode === 'syncDealHeaders')  return withLock(() => syncDealHeaders());
  if (mode === 'suggestNextAction') return suggestNextActionDate(e && e.parameter);
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
    const name = rows[i][6]; // G列：商材名
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
// GAS CacheService キャッシュ無効化ヘルパー
// ============================================================
const GAS_ALL_CACHE_KEY    = 'getAllData_v1';
const GAS_MASTER_CACHE_KEY = 'getMaster_v2';
const GAS_CACHE_CHUNK      = 90000; // CacheService は 100KB/エントリ制限 → 90KB ずつ分割

// チャンク分割してキャッシュに書き込む
function putChunkedCache_(cache, key, value, ttl) {
  try {
    if (value.length <= GAS_CACHE_CHUNK) {
      cache.put(key, value, ttl);
      cache.remove(key + '_n'); // 旧チャンクメタを削除
    } else {
      const n = Math.ceil(value.length / GAS_CACHE_CHUNK);
      const keys = [key + '_n'];
      for (let i = 0; i < n; i++) {
        cache.put(key + '_' + i, value.slice(i * GAS_CACHE_CHUNK, (i + 1) * GAS_CACHE_CHUNK), ttl);
        keys.push(key + '_' + i);
      }
      cache.put(key + '_n', String(n), ttl);
      cache.remove(key); // 旧単一キーを削除
    }
  } catch(e) {}
}

// チャンク分割されたキャッシュを読み込む
function getChunkedCache_(cache, key) {
  try {
    const simple = cache.get(key);
    if (simple) return simple;
    const n = Number(cache.get(key + '_n') || '0');
    if (!n) return null;
    const parts = [];
    for (let i = 0; i < n; i++) {
      const c = cache.get(key + '_' + i);
      if (!c) return null; // チャンク欠損 → キャッシュミス
      parts.push(c);
    }
    return parts.join('');
  } catch(e) { return null; }
}

function invalidateAllDataCache_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(GAS_ALL_CACHE_KEY + '_n') || '0');
    const keys = [GAS_ALL_CACHE_KEY, GAS_MASTER_CACHE_KEY, GAS_ALL_CACHE_KEY + '_n'];
    for (let i = 0; i < n; i++) keys.push(GAS_ALL_CACHE_KEY + '_' + i);
    c.removeAll(keys);
  } catch(e) {}
}

// ============================================================
// LockService ヘルパー（同時書き込み防止）
// 書き込み操作はすべてここを通るため、finally でキャッシュも無効化する
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
    invalidateAllDataCache_(); // 書き込み後に allData キャッシュを破棄
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
    if (action === 'migrateOldDB')         return json({ success: true, data: migrateOldCustomerDB() });
    if (action === 'clearAll')        return json({ success: false, error: '外部からの実行は許可されていません' });

    // ── 書き込み系（LockService で保護） ───────────────────
    return withLock(() => {
      if (action === 'addPerson')           return addPerson(d.name, d.role, d.code || '');
      if (action === 'setPersonStatus')     return setPersonStatus(d.name, d.status);
      if (action === 'deletePerson')        return deletePerson(d.name);
      if (action === 'addProduct')          return addProductToSheet(d.code, d.name, d.kind, d.unitPrice, d.cost, d.incentiveRate, d.priceType, d.incentiveFixedAmount);
      if (action === 'updateProduct')       return updateProduct(d);
      if (action === 'saveGoals')           return saveGoals(d);
      if (action === 'changeCompanyPerson') return changeCompanyPerson(d.code, d.person);
      if (action === 'initCustomerHistory') return initCustomerHistoryColumn();
      if (action === 'addDeal')             return addDeal(d);
      if (action === 'updateDeal')          return updateDeal(d);
      if (action === 'deleteDeal')          return deleteDeal(d.id);
      if (action === 'saveCostSchedule')    return saveCostSchedule(d);
      if (action === 'updateDealStatus')    return updateDealStatus(d.id, d.phase, d.rankLabel);
      if (action === 'confirmPayment')      return confirmPayment(d.id, d.date);
      if (action === 'handoverDeal')        return handoverDeal(d.id, d.newPerson, d.date);
      if (action === 'buildCustomersFromDeals') return json({ success: true, data: buildCustomersFromDeals() });
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
  if (!d.rankLabel || !String(d.rankLabel).trim())      return json({ success: false, error: '確度ランクが空です' });
  if (!VALID_DEAL_RANKS.includes(String(d.rankLabel).trim())) return json({ success: false, error: '確度ランクが無効です: ' + d.rankLabel });

  const sheet = getOrCreateDealSheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  // deal_form.html は no-cors POST でレスポンスを読めないため、費用計上明細を紐付ける場合は
  // クライアント側で生成したIDを使う（未指定時は従来通りサーバー側で生成）
  const id = (d.id && String(d.id).trim()) || generateDealId();

  // 商材コード優先、なければ商材名でlookup
  const pDetail = getProductDetail(d.productCode || d.productName);

  // フォームから単価・コース数・件数・月数を受け取り、合計を計算
  const unitSales = Number(d.unitSales) || (pDetail ? pDetail.unitPrice : 0);
  const unitCost  = Number(d.unitCost)  || (pDetail ? pDetail.cost : 0);
  const courses   = Math.max(1, Number(d.courses) || 1);
  const qty       = Math.max(1, Number(d.qty)     || 1);
  const months    = Math.max(1, Math.min(24, Number(d.months) || (pDetail ? pDetail.months : 1)));
  // B行（複合商材：社労士顧問・リスキリング・嘱託産業医など。月数に関わらず一回きりの金額）
  const bUnitSales = Number(d.bUnitSales) || 0;
  const bUnitCost  = Number(d.bUnitCost)  || 0;
  const bQty       = Number(d.bQty)       || 0;
  const bLumpGP    = (bUnitSales - bUnitCost) * bQty;

  const totalSales  = unitSales * courses * qty * months + bUnitSales * bQty;
  const totalCost   = unitCost  * courses * qty * months + bUnitCost  * bQty;
  const monthlyGP   = (unitSales - unitCost) * courses * qty;
  const grossProfit = monthlyGP * months + bLumpGP;

  const incentiveRate = pDetail ? pDetail.incentiveRate : 0;
  const incentiveFixedAmount = pDetail ? pDetail.incentiveFixedAmount : 0;
  // インセンティブはA行（月次経常GP）のみを対象とし、B行の一回きり金額は含めない
  const incentive = calcIncentive(monthlyGP, months, incentiveRate, incentiveFixedAmount, qty);

  const codeMap = {};
  getPersonDetails().forEach(p => { if (p.code) codeMap[p.code] = p.name; });
  const personName = codeMap[String(d.person || '').trim()] || String(d.person || '').trim();

  // 顧客IDルックアップ / 未登録は顧客マスタに自動追加
  const custSheet_ = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('顧客マスタ');
  let customerId_ = String(d.customerId || '').trim();
  if (!customerId_ && custSheet_) {
    const norm_ = normalizeCompany_(d.companyName || '');
    const custData_ = custSheet_.getDataRange().getValues();
    const match_ = custData_.find((r, i) => i > 0 && normalizeCompany_(String(r[1])) === norm_);
    if (match_) {
      customerId_ = String(match_[0]);
    } else {
      const newId_ = getNextId_(custSheet_, 'CUS');
      custSheet_.appendRow([
        newId_, d.companyName || '', '', '', '',
        String(d.person || ''), personName, '取引中', new Date(), new Date(), ''
      ]);
      customerId_ = newId_;
    }
  }

  sheet.appendRow([
    id,                          // A: 案件ID
    today,                       // B: 登録日
    d.person || '',              // C: 個人コード
    personName,                  // D: 営業名
    customerId_,                 // E: 顧客ID
    d.companyName || '',         // F: 会社名
    d.productName || '',         // G: 商材名
    d.phase || 'ヒアリング中',   // H: フェーズ
    String(d.rankLabel).trim(),  // I: 確度ランク
    unitSales,                   // J: 売上（単価）
    unitCost,                    // K: 費用（単価）
    courses,                     // L: コース数
    qty,                         // M: 件数
    months,                      // N: 月数
    totalSales,                  // O: 売上予定額（合計）
    totalCost,                   // P: 費用（合計）
    grossProfit,                 // Q: 粗利（合計）
    incentive,                   // R: インセンティブ
    d.expectedMonth || '',       // S: 売上予定月
    '未入金',                    // T: 入金ステータス
    '',                          // U: 入金確認日
    d.memo || '',                // V: メモ
    '',                          // W: 引継営業名
    '',                          // X: 引継日
    d.reason || '',              // Y: 理由
    today,                       // Z: 最終更新日
    d.billingCompany || '',      // AA: 計上会社
    Number(d.bUnitSales) || 0,  // AB: B売上単価
    Number(d.bUnitCost)  || 0,  // AC: B費用単価
    Number(d.bQty)       || 0,  // AD: B件数
    d.productCode || (pDetail ? pDetail.code : '') || '',  // AE: 商材コード
    !!d.recurring,                // AF: 継続課金
    '',                            // AG: 継続終了月（登録時は常に空＝継続中）
    String(d.rankLabel || '').trim() === '売上',  // AH: インセンティブ計上済み
    d.nextActionDate || '',       // AI: 次回アクション日
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

  // 個人コード→営業名、営業名→個人コードの双方向マップを構築
  const personCodeMap = {};
  const personNameToCode = {};
  getPersonDetails().forEach(p => {
    if (p.code) {
      personCodeMap[p.code] = p.name;
      if (p.name) personNameToCode[p.name] = p.code;
    }
  });

  const vals = sheet.getRange(1, 1, lastRow, DEAL_HEADERS.length).getValues();
  const headers = vals[0];

  const MON_MAP = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',
                   Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  function normYM(v) {
    if (!v) return '';
    // GASのスプレッドシートDate型はinstanceof Dateが効かない場合がある → duck typing
    if (typeof v === 'object' && typeof v.getFullYear === 'function') {
      try {
        const formatted = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
        if (/^\d{4}-\d{2}$/.test(formatted)) return formatted;
      } catch(e) {}
      return '';
    }
    const s = String(v).trim();
    // ISO文字列: "2025-07-31T15:00:00.000Z"
    const mIso = s.match(/^(\d{4})-(\d{2})-\d{2}T/);
    if (mIso) return mIso[1]+'-'+mIso[2];
    // "2025-08-01" や "2025-08" 形式
    if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
    // "2025/08/01" や "2025/08" 形式（スラッシュ区切り）
    const mSlash = s.match(/^(\d{4})\/(\d{2})/);
    if (mSlash) return mSlash[1]+'-'+mSlash[2];
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

  // 費用計上明細（案件ID → [{month, amount}, ...]）を1回だけ読み込む
  const costSchedMap = {};
  const csSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COST_SCHEDULE);
  if (csSheet && csSheet.getLastRow() > 1) {
    const csVals = csSheet.getRange(2, 1, csSheet.getLastRow() - 1, COST_SCHEDULE_HEADERS.length).getValues();
    csVals.forEach(r => {
      const dealId = String(r[0] || '').trim();
      const month = normYM(r[1]);
      const amount = Number(r[2]) || 0;
      if (!dealId || !month || amount <= 0) return;
      (costSchedMap[dealId] = costSchedMap[dealId] || []).push({ month, amount });
    });
  }

  let deals = vals.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    headers.forEach((h, i) => {
      if (h === '個人コード') {
        const raw = String(r[i] || '').trim();
        // 旧レコード対応: 個人コード列に営業名が入っている場合は逆引きでコードに変換
        o[h] = personCodeMap[raw] !== undefined ? raw : (personNameToCode[raw] || raw);
        o['担当者コード'] = o[h]; // 後方互換
      } else if (h === '営業名') {
        // シートのVLOOKUP値を使い、personCodeMapで最新名に解決
        const code = o['個人コード'] || '';
        o[h] = personCodeMap[code] || String(r[i] || '').trim() || code;
        o['担当者'] = o[h]; // 後方互換
      } else if (h === '売上予定月' || h === '継続終了月') {
        o[h] = normYM(r[i]);
      } else if (h === '継続課金') {
        o[h] = r[i] === true || String(r[i]).trim().toUpperCase() === 'TRUE';
      } else if (h === '登録日' || h === '入金確認日' || h === '引継日' || h === '次回アクション日') {
        o[h] = normDate(r[i]);
      } else if (typeof r[i] === 'object' && typeof r[i].getFullYear === 'function') {
        o[h] = Utilities.formatDate(r[i], 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        o[h] = r[i];
      }
    });
    // 後方互換キーを末尾で強制上書き（ヘッダー処理順に依存しない保証）
    o['担当者']     = o['営業名']   || '';
    o['担当者コード'] = o['個人コード'] || '';
    if (o['売上予定月']) {
      o['売上予定月'] = String(o['売上予定月']).trim().slice(0, 7);
    }
    if (o['継続終了月']) {
      o['継続終了月'] = String(o['継続終了月']).trim().slice(0, 7);
    }
    o['費用スケジュール'] = costSchedMap[o['案件ID']] || [];
    return o;
  });

  if (person) deals = deals.filter(d => d['個人コード'] === person || d['営業名'] === person || d['担当者'] === person);

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
  if (d.recurringEndMonth && !/^\d{4}-\d{2}$/.test(String(d.recurringEndMonth).trim())) {
    return json({ success: false, error: '継続終了月の形式が不正です（例: 2026-04）' });
  }
  // d.grossProfit は常に円単位で受け取る（フロントの dm-amount フォームが円入力）
  // 以前あった「10000未満→万円とみなして変換」の安全弁は誤変換を招くため削除済み
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, DEAL_HEADERS.length).getValues();

  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === d.id.trim()) {
      const row = i + 2;
      const prevRank = String(vals[i][DEAL_HEADERS.indexOf('確度ランク')] || '').trim();
      const setCol = (key, val) => {
        if (val === undefined || val === null) return;
        const col = DEAL_HEADERS.indexOf(key) + 1;
        if (col > 0) sheet.getRange(row, col).setValue(val);
      };

      // 常に更新可能なフィールド
      setCol('確度ランク', d.rankLabel);
      setCol('売上予定月', d.expectedMonth);
      setCol('メモ', d.memo);
      setCol('会社名', d.companyName);
      setCol('顧客ID', d.customerId);
      setCol('フェーズ', d.phase);
      setCol('理由', d.reason);
      setCol('入金ステータス', d.paymentStatus);
      setCol('計上会社', d.billingCompany);
      if (d.bUnitSales  !== undefined) setCol('B売上単価', Number(d.bUnitSales) || 0);
      if (d.bUnitCost   !== undefined) setCol('B費用単価', Number(d.bUnitCost)  || 0);
      if (d.bQty        !== undefined) setCol('B件数',     Number(d.bQty)       || 0);
      if (d.productCode !== undefined) setCol('商材コード', d.productCode || '');
      if (d.recurring   !== undefined) setCol('継続課金', !!d.recurring);
      if (d.recurringEndMonth !== undefined) setCol('継続終了月', d.recurringEndMonth);
      if (d.nextActionDate !== undefined) setCol('次回アクション日', d.nextActionDate);

      const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

      // ダッシュボードモーダルからの粗利直接更新
      if (d.grossProfit !== undefined) {
        setCol('粗利', Number(d.grossProfit));
        setCol('最終更新日', today);
        return json({ success: true });
      }

      // フォームからの単価・コース数・件数・月数更新（全再計算）
      if (d.unitSales !== undefined || d.unitCost !== undefined ||
          d.courses !== undefined || d.qty !== undefined || d.months !== undefined ||
          d.bUnitSales !== undefined || d.bUnitCost !== undefined || d.bQty !== undefined) {
        const unitSales = Number(d.unitSales !== undefined ? d.unitSales : vals[i][DEAL_HEADERS.indexOf('売上（単価）')]);
        const unitCost  = Number(d.unitCost  !== undefined ? d.unitCost  : vals[i][DEAL_HEADERS.indexOf('費用（単価）')]);
        const courses   = Math.max(1, Number(d.courses !== undefined ? d.courses : vals[i][DEAL_HEADERS.indexOf('コース数')]) || 1);
        const qty       = Math.max(1, Number(d.qty     !== undefined ? d.qty     : vals[i][DEAL_HEADERS.indexOf('件数')])   || 1);
        const months    = Math.max(1, Math.min(24, Number(d.months !== undefined ? d.months : vals[i][DEAL_HEADERS.indexOf('月数')]) || 1));
        const monthlyGP = (unitSales - unitCost) * courses * qty;
        // B行（複合商材：社労士顧問・リスキリング・嘱託産業医など。月数に関わらず一回きりの金額）
        const bUnitSales = Number(d.bUnitSales !== undefined ? d.bUnitSales : vals[i][DEAL_HEADERS.indexOf('B売上単価')]) || 0;
        const bUnitCost  = Number(d.bUnitCost  !== undefined ? d.bUnitCost  : vals[i][DEAL_HEADERS.indexOf('B費用単価')]) || 0;
        const bQty       = Number(d.bQty       !== undefined ? d.bQty       : vals[i][DEAL_HEADERS.indexOf('B件数')])     || 0;
        const bLumpGP    = (bUnitSales - bUnitCost) * bQty;
        const pDetail = getProductDetail(String(vals[i][DEAL_HEADERS.indexOf('商材名')]));
        const incentiveRate = pDetail ? pDetail.incentiveRate : 0;
        const incentiveFixedAmount = pDetail ? pDetail.incentiveFixedAmount : 0;

        // 顧問契約・リスキリング分割払いなど複数月に跨る案件を「決定」→「売上」へ更新する場合、
        // 当月分（1ヶ月分）のみを本レコードで売上確定し、残りの月数は新規レコードとして
        // 「決定」ランクのまま翌月以降に繰り越す（従来は月数分の全額が当月の売上になっていた）。
        // 継続課金（顧問契約の継続分・終了月未定）フラグの案件は別ロジック（継続終了月まで毎月計上）で
        // 扱われているため対象外。
        const isRecurring = !!vals[i][DEAL_HEADERS.indexOf('継続課金')];
        const becomingSold = String(d.rankLabel || '').trim() === '売上' && prevRank !== '売上';
        // ストック案件のインセンティブは月割りではなく、契約全期間分（例: 12ヶ月×レート）を
        // 初回の「決定→売上」確定時に一括計上する。繰り越した「決定」レコードを翌月以降また
        // 「売上」に確定する運用（毎月このsplit処理を通る）でも二重計上しないよう、
        // 一度計上したら「インセンティブ計上済み」フラグを立てて以降は0にする。
        const incentivePaidAlready = !!vals[i][DEAL_HEADERS.indexOf('インセンティブ計上済み')];
        if (becomingSold && months > 1 && !isRecurring) {
          // B行（一回きりの金額）は当月分の確定時にまとめて計上する
          const monthlySales = unitSales * courses * qty + bUnitSales * bQty;
          const monthlyCost  = unitCost  * courses * qty + bUnitCost  * bQty;
          // 初回のみ、この時点の月数（＝繰り越されていない契約全期間、または残りの繰り越し月数）
          // 分をまとめて計上。2回目以降（計上済み）は0。
          // インセンティブはA行（月次経常GP）のみを対象とし、B行の一回きり金額は含めない
          const soldIncentive = incentivePaidAlready
            ? 0
            : calcIncentive(monthlyGP, months, incentiveRate, incentiveFixedAmount, qty);

          // 当月分：本レコードを月数=1として売上確定
          setCol('売上（単価）', unitSales);
          setCol('費用（単価）', unitCost);
          setCol('コース数', courses);
          setCol('件数', qty);
          setCol('月数', 1);
          setCol('売上予定額', monthlySales);
          setCol('費用（合計）', monthlyCost);
          setCol('粗利', monthlyGP + bLumpGP);
          setCol('インセンティブ', soldIncentive);
          setCol('インセンティブ計上済み', true);
          setCol('最終更新日', today);

          // 残り月数分：新規レコードとして「決定」ランク・翌月以降に繰り越す
          const curExpMonth  = String(d.expectedMonth || vals[i][DEAL_HEADERS.indexOf('売上予定月')] || '').slice(0, 7);
          const nextExpMonth = addMonthsToYM_(curExpMonth, 1);
          const remMonths    = months - 1;
          const remTotalSales   = unitSales * courses * qty * remMonths;
          const remTotalCost    = unitCost  * courses * qty * remMonths;
          const remGrossProfit  = monthlyGP * remMonths;
          // インセンティブは当月分（soldIncentive）で払い済みのため、繰り越し分は常に0
          const remIncentive    = 0;

          const newRow = DEAL_HEADERS.map((h, hi) => vals[i][hi]);
          const set = (key, val) => { newRow[DEAL_HEADERS.indexOf(key)] = val; };
          set('案件ID', generateDealId());
          set('登録日', today);
          set('確度ランク', '決定');
          if (d.companyName !== undefined) set('会社名', d.companyName);
          if (d.customerId  !== undefined) set('顧客ID', d.customerId);
          if (d.phase       !== undefined) set('フェーズ', d.phase);
          set('売上（単価）', unitSales);
          set('費用（単価）', unitCost);
          set('コース数', courses);
          set('件数', qty);
          set('月数', remMonths);
          set('売上予定額', remTotalSales);
          set('費用（合計）', remTotalCost);
          set('粗利', remGrossProfit);
          set('インセンティブ', remIncentive);
          set('売上予定月', nextExpMonth);
          set('入金ステータス', '未入金');
          set('入金確認日', '');
          if (d.memo   !== undefined) set('メモ', d.memo);
          set('引継営業名', '');
          set('引継日', '');
          if (d.reason !== undefined) set('理由', d.reason);
          set('最終更新日', today);
          if (d.billingCompany !== undefined) set('計上会社', d.billingCompany);
          // B行は当月分の一回計上のため、繰り越しレコードには引き継がない
          set('B売上単価', 0);
          set('B費用単価', 0);
          set('B件数', 0);
          if (d.productCode !== undefined) set('商材コード', d.productCode);
          set('継続課金', false);
          set('継続終了月', '');
          if (d.nextActionDate !== undefined) set('次回アクション日', d.nextActionDate);
          // 繰り越しレコードも「計上済み」扱いにする（翌月以降また売上に確定しても再計上しないため）
          set('インセンティブ計上済み', true);
          sheet.appendRow(newRow);

          return json({ success: true, grossProfit: monthlyGP + bLumpGP, incentive: soldIncentive, splitDealId: newRow[DEAL_HEADERS.indexOf('案件ID')] });
        }

        const totalSales  = unitSales * courses * qty * months + bUnitSales * bQty;
        const totalCost   = unitCost  * courses * qty * months + bUnitCost  * bQty;
        const grossProfit = monthlyGP * months + bLumpGP;
        // 計上済みなら金額を再計算せず維持する（ダッシュボードの編集モーダルは保存の度に
        // 月数・単価を含む全項目を送ってくるため、計上済みインセンティブが月割りに
        // 縮んでしまうのを防ぐ）
        const incentive = incentivePaidAlready
          ? (Number(vals[i][DEAL_HEADERS.indexOf('インセンティブ')]) || 0)
          : calcIncentive(monthlyGP, months, incentiveRate, incentiveFixedAmount, qty);
        setCol('売上（単価）', unitSales);
        setCol('費用（単価）', unitCost);
        setCol('コース数', courses);
        setCol('件数', qty);
        setCol('月数', months);
        setCol('売上予定額', totalSales);
        setCol('費用（合計）', totalCost);
        setCol('粗利', grossProfit);
        setCol('インセンティブ', incentive);
        const nowSold = String(d.rankLabel !== undefined ? d.rankLabel : prevRank).trim() === '売上';
        if (!incentivePaidAlready && nowSold && incentive > 0) {
          setCol('インセンティブ計上済み', true);
        }
        setCol('最終更新日', today);
        return json({ success: true, grossProfit, incentive });
      }

      setCol('最終更新日', today);
      return json({ success: true });
    }
  }
  return json({ success: false, error: '見つかりません: ' + d.id });
}

// ============================================================
// 案件削除
// ============================================================
function deleteDeal(id) {
  if (!id) return json({ success: false, error: 'IDが空です' });
  const sheet = getOrCreateDealSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return json({ success: false, error: '見つかりません: ' + id });
  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 2);
      return json({ success: true });
    }
  }
  return json({ success: false, error: '見つかりません: ' + id });
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
      sheet.getRange(row, DEAL_HEADERS.indexOf('引継営業名') + 1).setValue(newPerson);
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

  // DEAL_HEADERS（31列）に合わせた列順:
  // A案件ID B登録日 C個人コード D営業名 E顧客ID F会社名 G商材名 Hフェーズ I確度ランク
  // J売上単価 K費用単価 Lコース数 M件数 N月数 O売上予定額 P費用合計 Q粗利
  // Rインセンティブ S売上予定月 T入金ステータス U入金確認日 Vメモ
  // W引継営業名 X引継日 Y理由 Z最終更新日
  // AA計上会社 AB B売上単価 AC B費用単価 AD B件数 AE商材コード
  const rows = [
    // TEST-001: 正常な売上確定（ストック12ヶ月・インセンティブあり）
    ['TEST-001', today, '', 'テスト担当者', '', '株式会社LOTY', 'HubCast_直販',
     '売上', '売上', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, Math.floor((35000-13000)*12*0.05),
     thisYM, '入金済み', today, '[TEST]001_正常売上_ストック12',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-002: 表記揺れ検証（' loTy ' → normalizeCompanyで株式会社LOTYと同一扱いになるか）
    ['TEST-002', today, '', 'テスト担当者', '', ' loTy ', 'HubCast_直販',
     'ヒアリング中', 'A', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, 0,
     nextYM, '未入金', '', '[TEST]002_表記揺れ(loTy)',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-003: ストック按分テスト（件数2・月数12で月次粗利が半分になるか）
    ['TEST-003', today, '', 'テスト担当者', '', 'ストックテスト株式会社', 'HubCast_直販',
     'ヒアリング中', 'B', 35000, 13000, 1, 2, 12,
     35000*2*12, 13000*2*12, (35000-13000)*2*12, 0,
     thisYM, '未入金', '', '[TEST]003_ストック按分（件数2）',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-004: スポット商材（月数1）
    ['TEST-004', today, '', 'テスト担当者', '', 'スポットテスト株式会社', 'IT支援研修事業',
     'クロージング中', 'A', 1000000, 0, 1, 1, 1,
     1000000, 0, 1000000, Math.floor(1000000*0.03),
     nextYM, '未入金', '', '[TEST]004_スポット商材（月数1）',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-005: 売上-費用=粗利 計算整合性テスト
    ['TEST-005', today, '', 'テスト担当者', '', '計算検証株式会社', 'リスキリング補助_直販',
     '提案中', 'B', 200000, 0, 1, 1, 1,
     200000, 0, 200000, 0,
     nextYM, '未入金', '', '[TEST]005_計算整合性チェック',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-006: B行テスト（計上会社・B売上単価・B費用単価・B件数あり）
    ['TEST-006', today, '', 'テスト担当者', '', 'B行テスト株式会社', 'HubCast_代理店',
     'クロージング中', 'A', 35000, 27000, 1, 3, 12,
     35000*3*12, 27000*3*12, (35000-27000)*3*12, 0,
     nextYM, '未入金', '', '[TEST]006_B行テスト',
     '', '', '', today, 'エー・ファクトリー株式会社', 40000, 30000, 2, ''],

    // TEST-007: インセンティブ計算テスト（決定ランク）
    ['TEST-007', today, '', 'テスト担当者', '', 'インセンティブテスト社', 'HubCast_直販',
     '売上', '決定', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, Math.floor((35000-13000)*12*0.05),
     thisYM, '未入金', '', '[TEST]007_インセンティブ計算',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-008: 複数コース数×件数のテスト
    ['TEST-008', today, '', 'テスト担当者', '', '複数件数テスト社', 'HubCast_直販',
     '提案中', 'C', 35000, 13000, 2, 3, 12,
     35000*2*3*12, 13000*2*3*12, (35000-13000)*2*3*12, 0,
     nextYM, '未入金', '', '[TEST]008_複数コース×件数',
     '', '', '', today, '', 0, 0, 0, ''],

    // TEST-009: 失注案件（パイプライン集計から除外されるか）
    ['TEST-009', today, '', 'テスト担当者', '', '失注テスト株式会社', 'HubCast_直販',
     '失注', '失注', 35000, 13000, 1, 1, 12,
     35000*12, 13000*12, (35000-13000)*12, 0,
     nextYM, '未入金', '', '[TEST]009_失注案件',
     '', '', '商談が長引いた', today, '', 0, 0, 0, ''],

    // TEST-010: 大文字・空白混じり表記揺れ（株式会社　LOTY → lotyと同一か）
    ['TEST-010', today, '', 'テスト担当者', '', '株式会社　LOTY', '現地長期補助_直販',
     'ヒアリング中', 'C', 600000, 0, 1, 1, 1,
     600000, 0, 600000, 0,
     nextYM, '未入金', '', '[TEST]010_表記揺れ（全角スペース入り）',
     '', '', '', today, '', 0, 0, 0, ''],
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
// 案件マスタのヘッダー行をDEAL_HEADERSに再同期（データ行には触れない）
// 列追加後にヘッダーラベルが追従していない場合の安全な修復用
// ============================================================
function syncDealHeaders() {
  const sheet = getOrCreateDealSheet();
  sheet.getRange(1, 1, 1, DEAL_HEADERS.length).setValues([DEAL_HEADERS]);
  return json({ success: true, headers: DEAL_HEADERS });
}

// ============================================================
// インセンティブ計算
// ============================================================
function calcIncentive(grossProfitMonthly, months, incentiveRate, fixedAmount, qty) {
  // 固定額商材（例: HubCast）は「固定額 × 件数」で計算し、率・月数は無視する
  const fixed = Number(fixedAmount) || 0;
  if (fixed > 0) return Math.floor(fixed * (Number(qty) || 1));
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
// 'yyyy-MM' 文字列に n ヶ月加算した 'yyyy-MM' 文字列を返す
// ============================================================
function addMonthsToYM_(ym, n) {
  const m = String(ym || '').match(/^(\d{4})-(\d{2})$/);
  const now = new Date();
  const y  = m ? Number(m[1]) : now.getFullYear();
  const mo = m ? Number(m[2]) : (now.getMonth() + 1);
  const d = new Date(y, mo - 1 + Number(n), 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
}

// ============================================================
// 商材詳細1件取得（ヘルパー）
// ============================================================
// nameOrCode: 商材コード(A列) または 商材名(B列)（どちらでも引ける、コード優先）
function getProductDetail(nameOrCode) {
  if (!nameOrCode) return null;
  const sheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const vals = sheet.getRange(2, 1, lastRow - 1, PRODUCT_HEADERS_V3.length).getValues();
  const query = String(nameOrCode).trim();
  // A列(index0)=商材コード優先、B列(index1)=商材名でフォールバック
  const row = vals.find(r => String(r[0]).trim() && String(r[0]).trim() === query)
           || vals.find(r => String(r[1]).trim() === query);
  if (!row) return null;
  const kind = String(row[2] || 'スポット').trim();
  return {
    code:          String(row[0] || '').trim(),   // A: 商材コード
    name:          String(row[1]).trim(),           // B: 商材名
    kind:          kind,                            // C: 種別
    months:        kind === 'ストック' ? 12 : 1,
    unitPrice:     Number(row[3]) || 0,            // D: 売上単価
    cost:          Number(row[4]) || 0,            // E: 費用
    incentiveRate: Number(row[5]) || 0,            // F: インセンティブ率
    priceType:     String(row[6]).trim() || (Number(row[3]) > 0 ? '固定' : '都度見積もり'), // G
    bUnitPrice:    Number(row[7]) || 0,            // H: B単価
    bCost:         Number(row[8]) || 0,            // I: B費用
    incentiveFixedAmount: Number(row[9]) || 0,     // J: インセンティブ固定額
  };
}

// ============================================================
// 商材マスタの単価変更 → 案件マスタへの自動反映
// ============================================================
// 設定_商材シートのD/E/H/I列（売上単価・費用・B単価・B費用）が編集された際に
// installableトリガー経由で呼ばれる。GASエディタ or setupProductPriceTrigger() で
// 一度だけトリガー登録が必要（clasp pushだけではトリガーは登録されない）。
function onProductPriceEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_PRODUCTS) return;
  if (e.range.getRow() === 1) return; // ヘッダー行

  const TARGET_COLS = { 4: '売上単価', 5: '費用', 8: 'B単価', 9: 'B費用' }; // D,E,H,I
  const field = TARGET_COLS[e.range.getColumn()];
  if (!field) return;

  const productCode = String(sheet.getRange(e.range.getRow(), 1).getValue() || '').trim(); // A列
  if (!productCode) return;

  const oldValue = Number(e.oldValue) || 0;
  const newValue = Number(e.value) || 0;
  if (oldValue === newValue) return;

  const result = syncDealsForProductPriceChange_(productCode, field, oldValue, newValue);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${result.updated}件反映 / ${result.skipped}件は個別単価のため対象外`,
      '商材単価の一括反映', 8
    );
  } catch (err) {}
}

// 商材マスタの単価変更を案件マスタに反映する中核ロジック（onEdit・手動実行の両方から呼べる）
// field: '売上単価' | '費用' | 'B単価' | 'B費用'
function syncDealsForProductPriceChange_(productCode, field, oldValue, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dealSheet = ss.getSheetByName(SHEET_DEALS);
  if (!dealSheet || dealSheet.getLastRow() <= 1) return { updated: 0, skipped: 0 };

  const idx = h => DEAL_HEADERS.indexOf(h);
  const COL_MAP = {
    '売上単価': { priceCol: idx('売上（単価）'), isB: false },
    '費用':     { priceCol: idx('費用（単価）'), isB: false },
    'B単価':    { priceCol: idx('B売上単価'),   isB: true  },
    'B費用':    { priceCol: idx('B費用単価'),   isB: true  },
  };
  const target = COL_MAP[field];
  if (!target) return { updated: 0, skipped: 0 };

  const range = dealSheet.getRange(2, 1, dealSheet.getLastRow() - 1, DEAL_HEADERS.length);
  const vals = range.getValues();
  const pDetail = getProductDetail(productCode); // 更新後の最新マスタ値（インセンティブ率など）

  let updated = 0, skipped = 0;
  const updatedIds = [];

  vals.forEach(row => {
    if (String(row[idx('商材コード')]).trim() !== productCode) return;
    // B系の変更はB行を持つ案件（B件数 > 0）のみ対象
    if (target.isB && (Number(row[idx('B件数')]) || 0) <= 0) return;

    const currentPrice = Number(row[target.priceCol]) || 0;
    // ガード: 現在値がマスタの「変更前の値」と一致する案件だけを自動反映する。
    // 値引き等で既にマスタと異なる単価にカスタマイズ済みの案件は対象外（スキップ）。
    if (currentPrice !== oldValue) { skipped++; return; }

    row[target.priceCol] = newValue;

    if (!target.isB) {
      // A行の単価変更は売上予定額・費用合計・粗利・インセンティブを再計算
      // （現行仕様どおりB行の金額はこの合計には含めない）
      const unitSales = Number(row[idx('売上（単価）')]) || 0;
      const unitCost  = Number(row[idx('費用（単価）')]) || 0;
      const courses   = Math.max(1, Number(row[idx('コース数')]) || 1);
      const qty       = Math.max(1, Number(row[idx('件数')])     || 1);
      const months    = Math.max(1, Number(row[idx('月数')])     || 1);
      const monthlyGP = (unitSales - unitCost) * courses * qty;

      row[idx('売上予定額')]   = unitSales * courses * qty * months;
      row[idx('費用（合計）')] = unitCost  * courses * qty * months;
      row[idx('粗利')]         = monthlyGP * months;
      // 計上済み（インセンティブ計上済み=true）の案件は、単価変更があっても
      // 月割りの再計算はしない（既に確定支給された金額を縮小させないため）
      const alreadyPaid = !!row[idx('インセンティブ計上済み')];
      if (!alreadyPaid) {
        row[idx('インセンティブ')] = calcIncentive(
          monthlyGP, months,
          pDetail ? pDetail.incentiveRate : 0,
          pDetail ? pDetail.incentiveFixedAmount : 0,
          qty
        );
      }
    }
    row[idx('最終更新日')] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    updatedIds.push(row[idx('案件ID')]);
    updated++;
  });

  if (updated) {
    range.setValues(vals);
    appendPriceChangeHistory_(productCode, pDetail ? pDetail.name : '', field, oldValue, newValue, updated, skipped, updatedIds);
    invalidateAllDataCache_();
  }
  return { updated, skipped };
}

function appendPriceChangeHistory_(productCode, productName, field, oldValue, newValue, updated, skipped, updatedIds) {
  const sheet = getOrCreateMasterSheet(SHEET_PRICE_HISTORY, PRICE_HISTORY_HEADERS);
  let executor = '';
  try { executor = Session.getActiveUser().getEmail(); } catch (e) {}
  sheet.appendRow([
    new Date(), productCode, productName, field, oldValue, newValue,
    updated, skipped, updatedIds.join(','), executor
  ]);
}

// 設定_商材シートの単価編集を検知するinstallableトリガーを登録する。
// GASエディタから一度だけ手動実行すること（clasp pushだけではトリガーは登録されない）。
function setupProductPriceTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onProductPriceEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onProductPriceEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log('onProductPriceEdit トリガーを登録しました');
}

// ============================================================
// 粗利整合性チェック（株式会社オアシス社労士顧問で発覚した事例への対応）
// ============================================================
// 単価（売上単価・費用単価・コース数・件数・月数・B行）だけ後から修正され、
// 粗利列が再計算されないまま古い値が残っている案件を洗い出す診断関数。
// updateDeal() が単価系フィールドを受け取ったときに再計算する式（monthlyGP*months+bLumpGP）と
// 同じ式で「本来あるべき粗利」を計算し、シートに保存済みの粗利と比較する。
// 読み取り専用（案件マスタは変更しない）。GASエディタの関数選択から手動実行すること。
function auditGrossProfitMismatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const rows = sheet.getDataRange().getValues();
  const idx = name => DEAL_HEADERS.indexOf(name);

  const results = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rank = row[idx('確度ランク')];
    if (!row[idx('案件ID')] || rank === '失注') continue;

    const unitSales  = Number(row[idx('売上（単価）')]) || 0;
    const unitCost   = Number(row[idx('費用（単価）')]) || 0;
    const courses    = Number(row[idx('コース数')]) || 1;
    const qty        = Number(row[idx('件数')]) || 1;
    const months     = Number(row[idx('月数')]) || 1;
    const bUnitSales = Number(row[idx('B売上単価')]) || 0;
    const bUnitCost  = Number(row[idx('B費用単価')]) || 0;
    const bQty       = Number(row[idx('B件数')]) || 0;
    const storedGP   = Number(row[idx('粗利')]) || 0;

    const expectedGP = (unitSales - unitCost) * courses * qty * months + (bUnitSales - bUnitCost) * bQty;
    const diff = storedGP - expectedGP;

    if (Math.abs(diff) >= 1) {
      results.push([
        row[idx('案件ID')], row[idx('会社名')], row[idx('商材名')], rank,
        storedGP, expectedGP, diff, i + 1
      ]);
    }
  }

  let outSheet = ss.getSheetByName('粗利整合性チェック');
  if (outSheet) outSheet.clearContents(); else outSheet = ss.insertSheet('粗利整合性チェック');
  outSheet.appendRow(['案件ID', '会社名', '商材名', '確度ランク', '保存済み粗利', '単価から計算した粗利', '差額', '案件マスタの行番号']);
  results.forEach(r => outSheet.appendRow(r));

  Logger.log(results.length + '件のズレを検出しました（詳細は「粗利整合性チェック」シート参照）');
  return results.length;
}

// auditGrossProfitMismatches() で洗い出した行の粗利列を、単価側を正として一括で書き換える。
// dryRun=true（デフォルト）はシートを書き換えず件数の確認のみ。内容を確認したうえで
// fixGrossProfitMismatches(false) を実行すると実際に案件マスタの粗利列を上書きする。
// updateDeal() の再計算式と完全に同じ式を使うため、ダッシュボードの編集モーダルで
// 該当案件を1件ずつ「保存」し直すのと同じ結果になる（それを一括で行うだけ）。
function fixGrossProfitMismatches(dryRun) {
  if (dryRun === undefined) dryRun = true;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const rows = sheet.getDataRange().getValues();
  const idx = name => DEAL_HEADERS.indexOf(name);
  const gpCol = idx('粗利') + 1;

  let fixed = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rank = row[idx('確度ランク')];
    if (!row[idx('案件ID')] || rank === '失注') continue;

    const unitSales  = Number(row[idx('売上（単価）')]) || 0;
    const unitCost   = Number(row[idx('費用（単価）')]) || 0;
    const courses    = Number(row[idx('コース数')]) || 1;
    const qty        = Number(row[idx('件数')]) || 1;
    const months     = Number(row[idx('月数')]) || 1;
    const bUnitSales = Number(row[idx('B売上単価')]) || 0;
    const bUnitCost  = Number(row[idx('B費用単価')]) || 0;
    const bQty       = Number(row[idx('B件数')]) || 0;
    const storedGP   = Number(row[idx('粗利')]) || 0;

    const expectedGP = (unitSales - unitCost) * courses * qty * months + (bUnitSales - bUnitCost) * bQty;
    if (Math.abs(storedGP - expectedGP) < 1) continue;

    fixed++;
    if (!dryRun) sheet.getRange(i + 1, gpCol).setValue(expectedGP);
  }

  Logger.log((dryRun ? '[dryRun] ' : '') + fixed + '件の粗利を' + (dryRun ? '修正対象として検出' : '修正しました'));
  if (!dryRun) invalidateAllDataCache_();
  return fixed;
}

// ============================================================
// 日報データ登録
// ============================================================
function addEntry(d) {
  if (!d.person || !String(d.person).trim()) return json({ success: false, error: '営業担当（個人コード）が空です' });
  if (!d.date   || !String(d.date).trim())   return json({ success: false, error: '日付が空です' });

  const sheet = getOrCreateResultSheet();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  const entryCodeMap = {};
  getPersonDetails().forEach(p => { if (p.code) entryCodeMap[p.code] = p.name; });
  const entryPersonName = entryCodeMap[String(d.person || '').trim()] || String(d.person || '').trim();

  if (d.activity !== undefined) {
    const act = d.activity || {};
    sheet.appendRow([
      now, d.date||'', d.person||'', entryPersonName, '【活動】',
      n(act.selfApo), n(act.telApo), n(act.refApo), n(act.agentRef),
      n(act.bizent), n(act.selfAgent),
      n(act.hearing), n(act.proposal), n(act.closing), 0,
      act.memo||'', '', '', '', '', ''
    ]);
    const deals = d.deals || [];
    deals.forEach(function(deal) {
      sheet.appendRow([
        now, d.date||'', d.person||'', entryPersonName, deal.product||'',
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
    now, d.date||'', d.person||'', entryPersonName, d.product||'',
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
  sheet.getRange(nextRow, 1, 1, 4).setValues([[code, name, role, '在籍中']]);
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
  const vals = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1]).trim() === name) {
      sheet.getRange(i + 2, 4).setValue(status);
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
  const vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1]).trim() === name) {
      sheet.deleteRow(i + 2);
      return json({ success: true, name: name });
    }
  }
  return json({ success: false, error: '見つかりません: ' + name });
}

// ============================================================
// 商材追加（新列対応）
// ============================================================
// code: 商材コード（A列）を先頭引数に
function addProductToSheet(code, name, kind, unitPrice, cost, incentiveRate, priceType, incentiveFixedAmount) {
  if (!name || !name.trim()) return json({ success: false, error: '商材名が空です' });
  name = name.trim();
  code = String(code || '').trim();
  const sheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  // 重複チェック：B列（商材名）で確認
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
    if (names.includes(name)) return json({ success: false, error: '既に存在します: ' + name });
  }
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, 9).setValues([[
    code,                        // A: 商材コード
    name,                        // B: 商材名
    kind          || 'スポット', // C: 種別
    Number(unitPrice)     || 0,  // D: 売上単価
    Number(cost)          || 0,  // E: 費用
    Number(incentiveRate) || 0,  // F: インセンティブ率
    priceType     || '固定',     // G: 価格タイプ
    0,                            // H: B単価
    0,                            // I: B費用
    Number(incentiveFixedAmount) || 0, // J: インセンティブ固定額
  ]]);
  return json({ success: true, name: name, code: code });
}

// ============================================================
// 商材更新（商材名またはコードで特定、フィールドを部分更新）
// ============================================================
function updateProduct(d) {
  const nameOrCode = String(d.name || d.code || '').trim();
  if (!nameOrCode) return json({ success: false, error: '商材名またはコードが必要です' });
  const sheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][0]).trim();
    const name = String(rows[i][1]).trim();
    if (name.toLowerCase() !== nameOrCode.toLowerCase() && code.toLowerCase() !== nameOrCode.toLowerCase()) continue;
    if (d.kind         !== undefined) sheet.getRange(i+1, 3).setValue(d.kind);
    if (d.unitPrice    !== undefined) sheet.getRange(i+1, 4).setValue(Number(d.unitPrice)||0);
    if (d.cost         !== undefined) sheet.getRange(i+1, 5).setValue(Number(d.cost)||0);
    if (d.incentiveRate!== undefined) sheet.getRange(i+1, 6).setValue(Number(d.incentiveRate)||0);
    if (d.priceType    !== undefined) sheet.getRange(i+1, 7).setValue(d.priceType);
    if (d.bUnitPrice   !== undefined) sheet.getRange(i+1, 8).setValue(Number(d.bUnitPrice)||0);
    if (d.bCost        !== undefined) sheet.getRange(i+1, 9).setValue(Number(d.bCost)||0);
    if (d.incentiveFixedAmount !== undefined) sheet.getRange(i+1, 10).setValue(Number(d.incentiveFixedAmount)||0);
    return json({ success: true, name, code });
  }
  return json({ success: false, error: '商材が見つかりません: ' + nameOrCode });
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

    // 今月末までのカレンダー予定（アポ・商談等）。チーム全体の場合は在籍中の全営業を横断集計
    const monthEndDate = new Date(today.getFullYear(), today.getMonth()+1, 0, 23, 59, 59);
    const calPersons = getPersonDetails().filter(p => p.email && (isTeam || p.name === person));
    let calEvents = [];
    calPersons.forEach(p => {
      getCalendarEventsInRange_(p.email, today, monthEndDate).forEach(e => {
        calEvents.push({ person: p.name, title: e.title, start: e.start, startRaw: e.startRaw });
      });
    });
    calEvents.sort((a,b) => a.startRaw - b.startRaw);
    const CAL_LIST_MAX = 12;
    const calSection = calEvents.length > 0
      ? `\n【今月末までの予定（Googleカレンダー・アポ/商談等）】${calEvents.length}件\n` +
        calEvents.slice(0, CAL_LIST_MAX).map(e => `- ${e.start}${isTeam ? '　'+e.person : ''}　${e.title}`).join('\n') +
        (calEvents.length > CAL_LIST_MAX ? `\n…他${calEvents.length - CAL_LIST_MAX}件` : '')
      : `\n【今月末までの予定（Googleカレンダー）】アポ・商談等の予定が見当たりません（未入力またはカレンダー未連携の可能性）`;

    const elapsed = Math.max(1, today.getMonth()>=7 ? today.getMonth()-7 : today.getMonth()+5);
    const prompt = `あなたはプロの営業マネジャーです。以下の営業データ（実績・案件・今後のカレンダー予定）を総合的に分析し、今月の改善アドバイスを日本語で3〜4点、箇条書きで具体的に提示してください。数字を必ず使ってください。カレンダーの予定件数や内容にも触れ、予定が少ない・偏っている場合はその点も指摘してください。

【対象】${isTeam?'チーム全体':'担当者: '+person}
【期間】FY${currentFY}（8月〜翌7月）経過${elapsed}ヶ月
【今期累計粗利】${actualTotal}万円 ／ 年間目標 ${annualTarget>0?annualTarget+'万円':'未設定'}
【今月粗利実績】${thisMonthGP}万円 ／ 月間目標 ${monthlyTarget>0?monthlyTarget+'万円':'未設定'}
【今月KPI】有効面談${meeting}件、自アポ${selfApo}件、テレアポ${telApo}件、紹介${refCount}件
【アクティブ案件】${activeDeals.length}件 ／ ヨミ合計 ${yomiTotal}万円
【確度内訳】決定:${activeDeals.filter(d=>d['確度ランク']==='決定').length}件 A:${activeDeals.filter(d=>d['確度ランク']==='A').length}件 B:${activeDeals.filter(d=>d['確度ランク']==='B').length}件 C:${activeDeals.filter(d=>d['確度ランク']==='C').length}件
${calSection}

アドバイスは実践的かつ前向きなトーンで。`;

    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key='+apiKey,
      { method:'POST', headers:{'Content-Type':'application/json'},
        payload:JSON.stringify({contents:[{parts:[{text:prompt}]}],
          generationConfig:{temperature:0.7,maxOutputTokens:4096,thinkingConfig:{thinkingBudget:0}}}),
        muteHttpExceptions:true }
    );
    const data = JSON.parse(res.getContentText());
    if (data.error) return json({ success:false, error:data.error.message });
    const parts = (data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) || [];
    const text = parts.filter(p => p && p.text).map(p => p.text).join('') || '応答がありません';
    return json({ success:true, text, person:person||'チーム全体', month:todayYM });
  } catch(err) {
    return json({ success:false, error:err.message });
  }
}

// ============================================================
// 次回アクション日のAI提案
// ルールベース（確度ランクごとの標準フォロー間隔）をデフォルトとし、
// メモ欄に「来週」「月末までに」等の時期の言及があればGeminiで解析して上書きする。
// 営業はこの提案をそのまま使うか、手動で書き換えてから保存するかを選べる（自動保存はしない）。
// ============================================================
const NEXT_ACTION_OFFSET_DAYS = { '決定': 30, 'A': 3, 'B': 7, 'C': 14 };

function suggestNextActionDate(params) {
  try {
    const rank = String((params && params.rankLabel) || '').trim();
    const memo = String((params && params.memo) || '').trim();
    const today = new Date();

    const offsetDays = NEXT_ACTION_OFFSET_DAYS[rank];
    if (offsetDays === undefined) {
      // 売上確定・失注は次回アクション不要
      return json({ success: true, date: '', reason: '確度ランク「' + (rank || '未設定') + '」は次回アクション日の対象外です。', source: 'rule' });
    }

    const ruleDate = new Date(today.getTime() + offsetDays * 86400000);
    let result = {
      date: Utilities.formatDate(ruleDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      reason: '確度ランク「' + rank + '」の標準フォロー間隔（' + offsetDays + '日後）による提案です。',
      source: 'rule',
    };

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey && memo) {
      try {
        const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
        const prompt = 'あなたは営業支援AIです。以下の商談メモから「次に営業担当がアクションすべき日付」を推測してください。\n' +
          '今日の日付: ' + todayStr + '\n' +
          '確度ランク: ' + rank + '\n' +
          '商談メモ: 「' + memo + '」\n\n' +
          'メモ内に「来週」「月末までに」「9月上旬」のような時期の言及があれば、具体的な日付（yyyy-MM-dd）に変換してください。\n' +
          '時期の言及が無ければ date は null にしてください。\n' +
          '必ず以下のJSON形式のみで回答してください（説明文やコードブロック記法は不要）:\n' +
          '{"date": "yyyy-MM-dd または null", "reason": "30文字以内の根拠"}';

        const res = UrlFetchApp.fetch(
          'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
          { method: 'POST', headers: { 'Content-Type': 'application/json' },
            payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }],
              generationConfig: { temperature: 0.2, maxOutputTokens: 1024, thinkingConfig: { thinkingBudget: 0 } } }),
            muteHttpExceptions: true }
        );
        const data = JSON.parse(res.getContentText());
        const text = data && data.candidates && data.candidates[0] && data.candidates[0].content &&
          data.candidates[0].content.parts && data.candidates[0].content.parts[0] &&
          data.candidates[0].content.parts[0].text || '';
        const m = text.match(/\{[\s\S]*\}/);
        if (m) {
          const parsed = JSON.parse(m[0]);
          if (parsed.date && /^\d{4}-\d{2}-\d{2}$/.test(parsed.date)) {
            result = { date: parsed.date, reason: parsed.reason || 'メモの記載内容から推測しました。', source: 'gemini' };
          }
        }
      } catch (e) {
        // Gemini解析に失敗してもルールベースの提案はそのまま返す
      }
    }
    return json({ success: true, date: result.date, reason: result.reason, source: result.source });
  } catch (err) {
    return json({ success: false, error: err.message });
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
function getAllData(force) {
  try {
    // ── GAS CacheService チェック（チャンク分割対応）──
    // force=true（手動リフレッシュ時）はキャッシュを読まず必ずシートから再取得する
    let gasCache = null;
    try { gasCache = CacheService.getScriptCache(); } catch(e) {}
    if (gasCache && !force) {
      const hit = getChunkedCache_(gasCache, GAS_ALL_CACHE_KEY);
      if (hit) {
        return ContentService.createTextOutput(hit)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 各データを内部で直接取得（HTTPコール不要）
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // master
    const masterResult = JSON.parse(getMaster().getContent());
    // ダッシュボードはcustomers（顧客DB全件ダンプ）を使わないため除外
    // → mode=master（deal_form.html等）はgetMaster()を直接呼ぶので影響なし
    delete masterResult.customers;

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
    // ダッシュボードの本番集計（月別推移・資金繰り・KPI・インセンティブ等）に
    // E2Eテストデータ（TEST-*／テスト担当者）が混入しないよう、mode=all の
    // レスポンスからは必ず除外する（テストデータの参照は mode=deals 経由に限定）
    const dealsResultRaw = JSON.parse(getDeals(null).getContent());
    const cleanDeals = (dealsResultRaw.deals || []).filter(function(d) { return !isTestDeal_(d); });
    const dealsResult = {
      success: dealsResultRaw.success,
      deals: cleanDeals,
      count: cleanDeals.length
    };

    // パイプライン by 担当者（確度ランク × 担当者別 粗利集計）
    const pipelineByPerson = buildPipelineByPerson_(cleanDeals);

    // 確度マッピング（確度ランク → 着地確率）
    const confidenceMap = getConfidenceMapping_();

    // 全担当者の今週行動KPI集計
    const weeklyKpi = getWeeklyKPISummaryAll_(getThisWeekMonday_());

    // doGet の mode=master ハンドラと同等のフィールドを付与
    masterResult.companies   = getCompanies();
    masterResult.topProducts = getTopProducts();

    const payload = JSON.stringify({
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

    // ── CacheService に保存（TTL 5分 / チャンク分割で100KB制限を回避）──
    if (gasCache) {
      putChunkedCache_(gasCache, GAS_ALL_CACHE_KEY, payload, 5 * 60);
    }

    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
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

  // シートを1回だけ読み込む（担当者ループ内での複数回読み込みを解消）
  const gRows = (goalSheet && goalSheet.getLastRow() > 1)
    ? goalSheet.getDataRange().getValues().slice(1) : [];
  const aRows = (actSheet && actSheet.getLastRow() > 1)
    ? actSheet.getDataRange().getValues().slice(1) : [];

  // 設定_週次目標: A=個人コード B=営業名 C=週開始日 D=架電目標 E=面談目標 F=紹介目標
  // 日次活動:     A=個人コード B=営業名 C=日付     D=架電数   E=面談数   F=紹介数
  return persons.map(function(p) {
    const personCode = p.code;
    const personName = p.name;

    // 週次目標を取得（A列=個人コードで照合）
    var target = { calls: 0, meetings: 0, referrals: 0 };
    const gRow = gRows.find(function(r) {
      return String(r[0]).trim() === personCode && toDateStr(r[2]) === weekStart;
    });
    if (gRow) {
      target = { calls: Number(gRow[3])||0, meetings: Number(gRow[4])||0, referrals: Number(gRow[5])||0 };
    }

    // 週次実績を集計（A列=個人コードで照合）
    var actual = { calls: 0, meetings: 0, referrals: 0 };
    aRows.forEach(function(r) {
      if (String(r[0]).trim() === personCode && weekDates.includes(toDateStr(r[2]))) {
        actual.calls     += Number(r[3]) || 0;
        actual.meetings  += Number(r[4]) || 0;
        actual.referrals += Number(r[5]) || 0;
      }
    });

    return {
      person: personName,
      code:   personCode,
      role:   p.role,
      target: {
        calls: target.calls, meetings: target.meetings, referrals: target.referrals,
        call:    target.calls,
        meeting: target.meetings,
        referral: target.referrals
      },
      actual: {
        calls: actual.calls, meetings: actual.meetings, referrals: actual.referrals,
        // フロントエンドの kpiSumFor が参照する旧キー（営業実績ヘッダー準拠）
        'ヒアリング':   actual.meetings,
        'テレアポ経由': actual.calls,
        '紹介会社':    actual.referrals,
        // 文字列キーエイリアス（フロントエンド側の参照方式に依存しない保証）
        'meetings':  actual.meetings,
        'calls':     actual.calls,
        'referrals': actual.referrals
      }
    };
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
function getMaster() {
  const personDetails = getPersonDetails();
  const persons = personDetails.filter(p => p.status === '在籍中').map(p => p.name);

  const productSheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const pVals = productSheet.getDataRange().getValues();
  const productDetails = pVals.slice(1).filter(r => String(r[1]).trim()).map(r => {
    // A(0)=商材コード B(1)=商材名 C(2)=種別 D(3)=売上単価 E(4)=費用
    // F(5)=インセンティブ率 G(6)=価格タイプ H(7)=B単価 I(8)=B費用 J(9)=インセンティブ固定額
    const kind = String(r[2] || 'スポット').trim();
    return {
      code:          String(r[0] || '').trim(),
      name:          String(r[1]).trim(),
      kind:          kind,
      months:        kind === 'ストック' ? 12 : 1,
      unitPrice:     Number(r[3]) || 0,
      cost:          Number(r[4]) || 0,
      incentiveRate: Number(r[5]) || 0,  // 0.10% → GASは0.001として読む
      priceType:     String(r[6]).trim() || (Number(r[3]) > 0 ? '固定' : '都度見積もり'),
      bUnitPrice:    Number(r[7]) || 0,
      bCost:         Number(r[8]) || 0,
      incentiveFixedAmount: Number(r[9]) || 0,
    };
  });
  const products = productDetails.map(p => p.name);

  const customers = getCustomersForDealForm_();
  return json({ success: true, persons, personDetails, products, productDetails, customers });
}

// ============================================================
// 案件登録フォーム用の顧客候補一覧（顧客マスタが正）
// customer_master.js の getAllCustomers_() を利用
// ============================================================
function getCustomersForDealForm_() {
  return getAllCustomers_().map(c => ({
    '顧客ID': c.id,
    '企業名': c.company,
    'AFC担当者': c.afcStaff,
  }));
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
// 実行単位（リクエスト単位）でシート読み込みをメモ化。
// GASは1リクエスト＝1実行なのでモジュールスコープ変数でOK。
let _personDetailsCache = null;

function getPersonDetails() {
  if (_personDetailsCache) return _personDetailsCache;
  const sheet = getOrCreatePersonSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return (_personDetailsCache = []);
  // A:個人コード B:営業名 C:役職 D:ステータス E:メールアドレス
  const vals = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  _personDetailsCache = vals
    .filter(r => String(r[1]).trim())
    .map(r => ({
      code:   String(r[0] || '').trim(),
      name:   String(r[1]).trim(),
      role:   String(r[2]).trim() || 'スタッフ',
      status: String(r[3]).trim() || '在籍中',
      email:  String(r[4] || '').trim()
    }));
  return _personDetailsCache;
}

// ============================================================
// 顧客DBの担当営業変更
// ============================================================
function changeCompanyPerson(code, newPerson) {
  try {
    if (!code || !newPerson) return json({ success: false, error: 'コードまたは営業コードが空です' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CUSTOMERS);
    if (!sheet) return json({ success: false, error: '顧客DBが見つかりません' });

    // 個人コード・営業名の列番号をヘッダーから動的解決
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const personCodeCol = headerRow.indexOf('個人コード') + 1; // 1-based, 0なら未存在
    const personNameCol = headerRow.indexOf('営業名')   + 1;

    // 個人コード→営業名を解決
    const persons = getPersonDetails();
    const matched = persons.find(p => p.code === newPerson) || persons.find(p => p.name === newPerson);
    const resolvedCode = matched ? matched.code : newPerson;
    const resolvedName = matched ? matched.name : newPerson;

    const histCol = ensureHistoryColumn(sheet);
    const lastRow = sheet.getLastRow();
    const allCodes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    for (let i = 0; i < allCodes.length; i++) {
      if (String(allCodes[i][0]).trim() === code.trim()) {
        const row = i + 2;
        const currentCode = personCodeCol > 0 ? String(sheet.getRange(row, personCodeCol).getValue()).trim() : '';
        if (currentCode && currentCode !== resolvedCode) {
          const histCell = sheet.getRange(row, histCol);
          const existing = String(histCell.getValue()).trim();
          const entry = currentCode + '(〜' + today + ')';
          histCell.setValue(existing ? existing + ', ' + entry : entry);
        }
        if (personCodeCol > 0) sheet.getRange(row, personCodeCol).setValue(resolvedCode);
        if (personNameCol > 0) sheet.getRange(row, personNameCol).setValue(resolvedName);
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
  // B列（商材名）を読む（A列は商材コードのため）
  const _prodSheet = getOrCreateMasterSheet(SHEET_PRODUCTS, PRODUCT_HEADERS_V3);
  const _prodLastRow = _prodSheet.getLastRow();
  const products = _prodLastRow > 1
    ? _prodSheet.getRange(2, 2, _prodLastRow - 1, 1).getValues().map(r => String(r[0]).trim()).filter(v => v)
    : [];
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
  // 旧シート名 '設定_担当者' へのフォールバック（担当者→営業 移行期対応）
  if ((!sheet || sheet.getLastRow() <= 1) && SHEET_PERSONS !== '設定_担当者') {
    const legacy = ss.getSheetByName('設定_担当者');
    if (legacy && legacy.getLastRow() > 1) return legacy;
  }
  if (!sheet) sheet = ss.insertSheet(SHEET_PERSONS);
  const h1 = String(sheet.getRange(1,1).getValue()).trim();
  if (!h1) {
    const headers = ['個人コード','営業名','役職','ステータス'];
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
// 費用計上明細シート（案件ごとの費用を複数月へ分散計上する明細）
// ============================================================
function getOrCreateCostScheduleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_COST_SCHEDULE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COST_SCHEDULE);
    const hr = sheet.getRange(1, 1, 1, COST_SCHEDULE_HEADERS.length);
    hr.setValues([COST_SCHEDULE_HEADERS]);
    hr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200); // 案件ID
    sheet.setColumnWidth(2, 100); // 計上月
    sheet.setColumnWidth(3, 110); // 金額
  }
  return sheet;
}

// 案件IDの費用計上明細をすべて置き換える（月・金額ともに正しいもののみ保存）
function saveCostSchedule(d) {
  if (!d.id) return json({ success: false, error: 'IDが空です' });
  const sheet = getOrCreateCostScheduleSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]).trim() === String(d.id).trim()) sheet.deleteRow(i + 2);
    }
  }
  const schedule = Array.isArray(d.schedule) ? d.schedule : [];
  schedule.forEach(s => {
    const month = String(s.month || '').trim();
    const amount = Number(s.amount) || 0;
    if (!/^\d{4}-\d{2}$/.test(month) || amount <= 0) return;
    sheet.appendRow([d.id, month, amount]);
  });
  return json({ success: true });
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
    .addItem('案件マスタの顧客IDを自動補完する', 'fillMissingCustomerIds')
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
// person には必ず個人コード（例: "S001"）を渡すこと。営業名を渡すとシートとの照合が失敗する。
function getWeeklyData(person, weekStart) {
  if (!person || !weekStart) return { success: false, error: 'person と weekStart は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 名前で渡された場合はコードに変換（addActivity と同じ方式）
  const wdCodeMap = {};
  const wdNameToCode = {};
  getPersonDetails().forEach(p => {
    if (p.code) {
      wdCodeMap[p.code] = p.name;
      if (p.name) wdNameToCode[p.name] = p.code;
    }
  });
  const personCode = wdCodeMap[person] !== undefined ? person : (wdNameToCode[person] || person);

  // 週次目標を取得
  // 設定_週次目標: A=個人コード B=営業名 C=週開始日 D=架電目標 E=面談目標 F=紹介目標
  const goalSheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  let target = null;
  if (goalSheet && goalSheet.getLastRow() > 1) {
    const rows = goalSheet.getDataRange().getValues().slice(1);
    const goalRow = rows.find(r => String(r[0]) === personCode && toDateStr(r[2]) === weekStart);
    if (goalRow) target = { calls: Number(goalRow[3]) || 0, meetings: Number(goalRow[4]) || 0, referrals: Number(goalRow[5]) || 0 };
  }

  // 当週の日次活動を取得（月〜金）
  // 日次活動: A=個人コード B=営業名 C=日付 D=架電数 E=面談数 F=紹介数 G=コメント
  const actSheet = ss.getSheetByName(SHEET_ACTIVITIES);
  const daily = [];
  if (actSheet && actSheet.getLastRow() > 1) {
    const rows = actSheet.getDataRange().getValues().slice(1);
    const weekDates = getWeekDateRange(weekStart);
    rows.forEach(r => {
      const rowPerson = String(r[0]);
      const rowDate   = toDateStr(r[2]);
      if (rowPerson === personCode && weekDates.includes(rowDate)) {
        daily.push({ date: rowDate, calls: Number(r[3]) || 0, meetings: Number(r[4]) || 0, referrals: Number(r[5]) || 0 });
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
  const actCodeMap = {};
  const actNameToCode = {};
  getPersonDetails().forEach(p => {
    if (p.code) {
      actCodeMap[p.code] = p.name;
      if (p.name) actNameToCode[p.name] = p.code;
    }
  });
  // d.person は個人コードを期待。名前で渡された場合はコードに変換
  const personCode    = actCodeMap[person] !== undefined ? person : (actNameToCode[person] || person);
  const actPersonName = actCodeMap[personCode] || personCode;

  let sheet = ss.getSheetByName(SHEET_ACTIVITIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVITIES);
    sheet.appendRow(['個人コード', '営業名', '日付', '架電数', '面談数', '紹介数', 'コメント']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 同日のレコードがあれば更新、なければ追記
  // A=個人コード B=営業名 C=日付 D=架電数 E=面談数 F=紹介数 G=コメント
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const vals = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === personCode && String(vals[i][2]).slice(0,10) === date) {
        sheet.getRange(i + 2, 4, 1, 4).setValues([[calls, meetings, referrals, comment]]);
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([personCode, actPersonName, date, calls, meetings, referrals, comment]);
  return { success: true, updated: false };
}

// ============================================================
// 週次目標保存（setWeeklyTarget）
// ============================================================
function setWeeklyTarget(d) {
  const person         = String(d.person    || '').trim();
  const weekStart      = String(d.weekStart || '').trim();
  const callTarget     = Number(d.callTarget)     || 0;
  const meetingTarget  = Number(d.meetingTarget)  || 0;
  const referralTarget = Number(d.referralTarget) || 0;
  if (!person || !weekStart) return { success: false, error: 'person と weekStart は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wgCodeMap = {};
  const wgNameToCode = {};
  getPersonDetails().forEach(p => {
    if (p.code) {
      wgCodeMap[p.code] = p.name;
      if (p.name) wgNameToCode[p.name] = p.code;
    }
  });
  // d.person は個人コードを期待。名前で渡された場合はコードに変換
  const personCode   = wgCodeMap[person] !== undefined ? person : (wgNameToCode[person] || person);
  const wgPersonName = wgCodeMap[personCode] || personCode;

  let sheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY_GOALS);
    sheet.appendRow(['個人コード', '営業名', '週開始日', '架電目標', '面談目標', '紹介目標']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 同担当者・同週があれば更新
  // A=個人コード B=営業名 C=週開始日 D=架電目標 E=面談目標 F=紹介目標
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const vals = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === personCode && String(vals[i][2]).slice(0,10) === weekStart) {
        sheet.getRange(i + 2, 4, 1, 3).setValues([[callTarget, meetingTarget, referralTarget]]);
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([personCode, wgPersonName, weekStart, callTarget, meetingTarget, referralTarget]);
  return { success: true, updated: false };
}

// ============================================================
// 月次KPI目標保存（setMonthlyKPITarget）
// ============================================================
function setMonthlyKPITarget(d) {
  const person  = String(d.person || '').trim();
  const month   = String(d.month  || '').trim(); // "2026-05"
  const weeks   = d.weeks || [];
  if (!person || !month) return { success: false, error: 'person と month は必須です' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mkpiCodeMap = {};
  const mkpiNameToCode = {};
  getPersonDetails().forEach(p => {
    if (p.code) {
      mkpiCodeMap[p.code] = p.name;
      if (p.name) mkpiNameToCode[p.name] = p.code;
    }
  });
  const personCode     = mkpiCodeMap[person] !== undefined ? person : (mkpiNameToCode[person] || person);
  const mkpiPersonName = mkpiCodeMap[personCode] || personCode;

  let sheet = ss.getSheetByName(SHEET_WEEKLY_GOALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY_GOALS);
    sheet.appendRow(['個人コード', '営業名', '週開始日', '架電目標', '面談目標', '紹介目標']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f4f8');
  }

  // 既存データを一括読み込みし、対象担当者・同月行をメモリ上でフィルタ除外
  // → deleteRow ループによる行インデックスズレを完全に回避
  const all = sheet.getDataRange().getValues();
  const kept = all.slice(1).filter(r =>
    !(String(r[0]).trim() === personCode && toDateStr(r[2]).slice(0, 7) === month)
  );

  // 新しい週別目標行をメモリ上で追記
  weeks.forEach(w => {
    const weekStart = String(w.weekStart || '').slice(0, 10);
    if (!weekStart) return;
    kept.push([personCode, mkpiPersonName, weekStart, w.calls || 0, w.meetings || 0, w.referrals || 0]);
  });

  // ヘッダー以降を一括クリアして setValues で原子的に書き込む
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (kept.length > 0) sheet.getRange(2, 1, kept.length, 6).setValues(kept);

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
const CALENDAR_EVENT_KEYWORDS = /アポ|架電|商談|面談|訪問|MTG|ミーティング|打ち合わせ/i;

// メールアドレス1件分のカレンダーから、期間内の営業活動系イベントを抽出する共通ヘルパー
function getCalendarEventsInRange_(email, startDate, endDate) {
  try {
    const cal = CalendarApp.getCalendarById(email);
    if (!cal) return [];
    return cal.getEvents(startDate, endDate)
      .filter(e => CALENDAR_EVENT_KEYWORDS.test(e.getTitle()))
      .map(e => ({
        title:    e.getTitle(),
        startRaw: e.getStartTime(),
        start:    Utilities.formatDate(e.getStartTime(), 'Asia/Tokyo', 'M/d(E) HH:mm'),
        end:      Utilities.formatDate(e.getEndTime(),   'Asia/Tokyo', 'HH:mm'),
      }));
  } catch (e) {
    return [];
  }
}

function getCalendarEventsForWeek(personName, weekStart) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PERSONS);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues().slice(1);
  const personRow = rows.find(r => String(r[1]).trim() === personName);
  if (!personRow || !personRow[4]) return []; // E列 = メールアドレス

  const startDate = new Date(weekStart);
  const endDate   = new Date(weekStart);
  endDate.setDate(endDate.getDate() + 5); // 月〜金

  return getCalendarEventsInRange_(String(personRow[4]).trim(), startDate, endDate);
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
          generationConfig: { temperature: 0.7, maxOutputTokens: 2048, thinkingConfig: { thinkingBudget: 0 } } }),
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
// 案件マスタの空欄を一括補完
//   ① 案件IDが空 → DL-YYYYMMDDHHmmss-XXXX 形式で自動生成
//   ② 顧客IDが空 → 顧客DBで会社名を照合、未登録なら新規追加
// ============================================================
function fillMissingCustomerIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── 顧客DB ロード ─────────────────────────────────────────
  const custSheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!custSheet) { ui.alert('「顧客DB」シートが見つかりません。'); return; }

  const custVals    = custSheet.getDataRange().getValues();
  const custHeaders = custVals[0].map(h => String(h).trim());
  const codeIdx     = custHeaders.indexOf('顧客コード');
  const companyIdx  = custHeaders.indexOf('企業名');
  const afcIdx      = custHeaders.indexOf('AFC営業担当');
  if (codeIdx < 0 || companyIdx < 0) { ui.alert('顧客DBに「顧客コード」または「企業名」列が見つかりません。'); return; }

  const companyToCode = {};
  const existingCodes = new Set();
  custVals.slice(1).forEach(r => {
    const name = String(r[companyIdx] || '').trim();
    const code = String(r[codeIdx]    || '').trim();
    if (code) existingCodes.add(code);
    if (name && code) companyToCode[name] = code;
  });

  // ── 案件マスタ ロード ────────────────────────────────────
  const dealSheet = getOrCreateDealSheet();
  const lastRow   = dealSheet.getLastRow();
  if (lastRow <= 1) { ui.alert('案件マスタにデータがありません。'); return; }

  const vals          = dealSheet.getRange(1, 1, lastRow, DEAL_HEADERS.length).getValues();
  const headers       = vals[0].map(h => String(h).trim());
  const dealIdCol     = headers.indexOf('案件ID');    // index 0
  const customerIdCol = headers.indexOf('顧客ID');    // index 4
  const companyCol    = headers.indexOf('会社名');    // index 5
  const personNameCol = headers.indexOf('営業名');    // index 3
  if (customerIdCol < 0 || companyCol < 0) { ui.alert('案件マスタのヘッダーが想定外です。'); return; }

  // ── ① 未登録会社を先に顧客DBへ追加 ─────────────────────
  //    （案件IDの有無を問わず会社名があればすべて対象）
  const newCompanies = {};
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const companyName = String(row[companyCol] || '').trim();
    if (!companyName) continue;
    if (String(row[customerIdCol]).trim()) continue;       // 顧客ID既入力はスキップ
    if (!companyToCode[companyName] && !newCompanies[companyName]) {
      newCompanies[companyName] = String(row[personNameCol] || '').trim();
    }
  }

  const numCols = custHeaders.length;
  let registered = 0;
  Object.entries(newCompanies).forEach(([name, personName]) => {
    let code;
    do { code = generateCustomerCode(); } while (existingCodes.has(code));
    existingCodes.add(code);
    companyToCode[name] = code;

    const newRow = new Array(numCols).fill('');
    newRow[codeIdx]    = code;
    newRow[companyIdx] = name;
    if (afcIdx >= 0) newRow[afcIdx] = personName;
    custSheet.appendRow(newRow);
    registered++;
  });

  // ── ② 案件ID・顧客ID を1行ずつ補完 ─────────────────────
  let filledDealId     = 0;
  let filledCustomerId = 0;

  for (let i = 1; i < vals.length; i++) {
    const row         = vals[i];
    const companyName = String(row[companyCol] || '').trim();

    // 会社名もなく案件IDもない行は完全空行とみなしてスキップ
    if (!companyName && !String(row[dealIdCol] || '').trim()) continue;

    // ① 案件ID が空なら生成
    if (!String(row[dealIdCol] || '').trim()) {
      const newId = generateDealId();
      dealSheet.getRange(i + 1, dealIdCol + 1).setValue(newId);
      filledDealId++;
    }

    // ② 顧客ID が空で会社名があれば補完
    if (companyName && !String(row[customerIdCol] || '').trim()) {
      const code = companyToCode[companyName];
      if (code) {
        dealSheet.getRange(i + 1, customerIdCol + 1).setValue(code);
        filledCustomerId++;
      }
    }
  }

  ui.alert(
    `完了しました。\n` +
    `・案件IDを自動生成: ${filledDealId}件\n` +
    `・顧客DBに新規登録: ${registered}社\n` +
    `・顧客IDを補完: ${filledCustomerId}件`
  );
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
  // [name, kind, unitPrice, cost, incentiveRate, priceType]
  // kind: 'ストック'=12ヶ月継続 / 'スポット'=単発
  // incentiveRate: 小数（0.1 = 10%、0.001 = 0.1%）
  const products = [
    ['リスキリング研修_直販',   'スポット', 200000,       0, 0.1,   '固定'],
    ['リスキリング研修_代理店', 'スポット', 200000,  150000, 0.1,   '固定'],
    ['HubCast_直販',            'ストック',  35000,   13000, 0.1,   '固定'],
    ['HubCast_代理店',          'ストック',  35000,   27000, 0.1,   '固定'],
    ['IT導入補助金',            'スポット', 1000000,      0, 0.001, '固定'],
    ['社長紹介_直販',           'スポット', 600000,       0, 0.1,   '固定'],
    ['社長紹介_代理店',         'スポット', 600000,  200000, 0.1,   '固定'],
    ['健康診断事務代理',        'スポット',      0,       0, 0.1,   '都度見積もり'],
    ['保健師電話健康相談',      'スポット',      0,       0, 0.1,   '都度見積もり'],
    ['社労士案件',              'スポット',      0,       0, 0.1,   '都度見積もり'],
  ];
  // addProductToSheet(code, name, kind, unitPrice, cost, incentiveRate, priceType)
  // ※ この関数は初回セットアップ用。現在の設定_商材シートを優先すること。
  products.forEach(([name, kind, unitPrice, cost, incentiveRate, priceType]) => {
    addProductToSheet('', name, kind, unitPrice, cost, incentiveRate, priceType);
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

  // 商材マスタをルックアップマップに変換（インセンティブ・月数計算用）
  const prodByName = {}, prodByCode = {};
  const _prodSheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (_prodSheet) {
    _prodSheet.getDataRange().getValues().slice(1).forEach(r => {
      const code = String(r[0]||'').trim();
      const name = String(r[1]||'').trim();
      if (!name) return;
      const obj = {
        code, name,
        kind:          String(r[2]||'スポット').trim(),
        months:        String(r[2]).trim() === 'ストック' ? 12 : 1,
        incentiveRate: Number(r[5]) || 0,
        incentiveFixedAmount: Number(r[9]) || 0,
      };
      prodByName[name.toLowerCase()] = obj;
      if (code) prodByCode[code] = obj;
    });
  }

  // 営業名→個人コードの変換マップを構築
  const personNameToCode = {};
  getPersonDetails().forEach(p => { if (p.name && p.code) personNameToCode[p.name] = p.code; });

  const VALID_RANKS = ['売上', '決定', 'A', 'B', 'C', '失注'];
  let added = 0, skipped = 0, errors = [];

  // 重複チェック用に既存データを一度だけ読み込む
  const iPersonIdx  = DEAL_HEADERS.indexOf('個人コード');
  const iCompanyIdx = DEAL_HEADERS.indexOf('会社名');
  const iMonthIdx   = DEAL_HEADERS.indexOf('売上予定月');
  const existingRows = dest.getLastRow() > 1
    ? dest.getDataRange().getValues().slice(1)
    : [];

  // 顧客マスタをルックアップマップに構築（インポート内での重複新規登録防止）
  const custSheet = ss.getSheetByName('顧客マスタ');
  const custNormMap = {}; // 正規化企業名 → 顧客ID
  if (custSheet) {
    custSheet.getDataRange().getValues().forEach((r, i) => {
      if (i === 0 || !r[0]) return;
      custNormMap[normalizeCompany_(String(r[1]))] = String(r[0]);
    });
  }

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

    // 顧客IDルックアップ / 未登録は顧客マスタに自動追加
    let customerId = '';
    if (custSheet) {
      const norm = normalizeCompany_(company);
      if (custNormMap[norm]) {
        customerId = custNormMap[norm];
      } else {
        const newId = getNextId_(custSheet, 'CUS');
        custSheet.appendRow([
          newId, company, '', '', '', person, personName, '取引中',
          new Date(), new Date(), ''
        ]);
        custNormMap[norm] = newId;
        customerId = newId;
      }
    }

    // 案件ID生成
    const now = new Date();
    const id = 'DEAL-IMP-' + Utilities.formatDate(now,'Asia/Tokyo','yyyyMMddHHmmss') + '-' + (added+1);
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    const phase = (rank === '売上' || rank === '決定') ? '完了' : 'ヒアリング中';
    const payStatus = rank === '売上' ? '入金済み' : '未入金';

    // 商材マスタから月数・インセンティブを計算
    const pDetail      = prodByCode[product] || prodByName[product.toLowerCase()];
    const importMonths = pDetail ? pDetail.months : 1;
    const monthlyGP    = (sales - cost) * 1 * 1; // コース数=1・件数=1 固定のため
    const incentive    = (pDetail && (pDetail.incentiveRate || pDetail.incentiveFixedAmount))
      ? calcIncentive(monthlyGP, importMonths, pDetail.incentiveRate, pDetail.incentiveFixedAmount, 1)
      : 0;

    const rowMap = {
      '案件ID': id, '登録日': today, '個人コード': person, '営業名': personName, '顧客ID': customerId,
      '会社名': company, '商材名': product, 'フェーズ': phase, '確度ランク': rank,
      '売上（単価）': sales, '費用（単価）': cost, 'コース数': 1, '件数': 1, '月数': importMonths,
      '売上予定額': sales, '費用（合計）': cost, '粗利': gp,
      'インセンティブ': incentive, '売上予定月': expMonth,
      '入金ステータス': payStatus, '入金確認日': '', 'メモ': memo,
      '引継営業名': '', '引継日': '', '理由': '', '最終更新日': today,
      '計上会社': billingCo, 'B売上単価': bUnitSales, 'B費用単価': bUnitCost, 'B件数': bQty,
      'インセンティブ計上済み': rank === '売上' && incentive > 0
    };
    const newRow = DEAL_HEADERS.map(h => rowMap[h] !== undefined ? rowMap[h] : '');
    dest.appendRow(newRow);
    existingRows.push(newRow); // 同一インポート内での重複防止
    added++;
  });

  Logger.log('【インポート完了】 追加:'+added+'件 / スキップ:'+skipped+'件');
  if (errors.length) Logger.log('【エラー詳細】\n'+errors.join('\n'));
  invalidateAllDataCache_();
}

// ============================================================
// 担当者マスタ初期設定（GASエディタから一度だけ実行）
// 鈴木純平/af0003 を 設定_営業 シートに追加する
// ============================================================
function setupPersonSuzuki() {
  const sheet = getOrCreatePersonSheet();
  const vals = sheet.getDataRange().getValues();
  const exists = vals.some(r => String(r[0]).trim() === 'af0003');
  if (exists) {
    Logger.log('af0003（鈴木純平）は既に登録済みです');
    return;
  }
  sheet.appendRow(['af0003', '鈴木純平', 'スタッフ', '在籍中', '']);
  Logger.log('af0003（鈴木純平）を設定_営業シートに追加しました');
}

// ============================================================
// インポート時の異常インセンティブ値を修正（一度だけ実行）
// DEAL-IMP- で始まる案件のうちインセンティブが100万超のものを0にリセット
// ============================================================
function fixImportedIncentives() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx  = headers.indexOf('案件ID');
  const incIdx = headers.indexOf('インセンティブ');
  let fixed = 0;
  data.slice(1).forEach((row, i) => {
    const id  = String(row[idIdx]  || '');
    const inc = Number(row[incIdx]) || 0;
    if (!id.startsWith('DEAL-IMP-') || inc < 1000000) return;
    sheet.getRange(i + 2, incIdx + 1).setValue(0);
    fixed++;
  });
  invalidateAllDataCache_();
  Logger.log('fixImportedIncentives: ' + fixed + '件修正、キャッシュクリア済み');
}

// ============================================================
// インセンティブ率を正しい値に設定し、案件マスタ全件を再計算して保存
// 社労士・助成金 → 0.1%（0.001）/ それ以外 → 10%（0.10）
// GASエディタから一度だけ実行する
// ============================================================
function fixIncentivesWithRates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Step 1: 設定_商材 F列（インセンティブ率）を一括更新 ---
  const prodSheet = ss.getSheetByName('設定_商材');
  const prodVals  = prodSheet.getDataRange().getValues();
  const rateUpdates = [];
  for (let i = 1; i < prodVals.length; i++) {
    const name = String(prodVals[i][1] || ''); // B列: 商材名
    const isSharoushi = name.includes('社労士') || name.includes('助成金');
    rateUpdates.push([isSharoushi ? 0.001 : 0.10]);
  }
  if (rateUpdates.length) {
    prodSheet.getRange(2, 6, rateUpdates.length, 1).setValues(rateUpdates);
  }

  // 更新後のマスタを再読み込みしてルックアップマップ構築
  const prodVals2 = prodSheet.getDataRange().getValues();
  const prodByCode = {}, prodByName = {};
  for (let i = 1; i < prodVals2.length; i++) {
    const code = String(prodVals2[i][0] || '').trim();
    const name = String(prodVals2[i][1] || '').trim().toLowerCase();
    const rate = Number(prodVals2[i][5]) || 0;
    if (code) prodByCode[code] = rate;
    if (name) prodByName[name] = rate;
  }

  // --- Step 2: 案件マスタ全件のインセンティブを再計算してバッチ書き込み ---
  const dealSheet = ss.getSheetByName('案件マスタ');
  const dealVals  = dealSheet.getDataRange().getValues();
  const dh = dealVals[0];
  const idx = col => dh.indexOf(col);

  const iSales    = idx('売上（単価）');
  const iCost     = idx('費用（単価）');
  const iCourses  = idx('コース数');
  const iQty      = idx('件数');
  const iMonths   = idx('月数');
  const iInc      = idx('インセンティブ');
  const iProdCode = idx('商材コード');
  const iProdName = idx('商材名');

  const incValues = [];
  let updated = 0, skipped = 0;

  for (let i = 1; i < dealVals.length; i++) {
    const row = dealVals[i];
    const code = String(row[iProdCode] || '').trim();
    const name = String(row[iProdName] || '').trim().toLowerCase();
    const rate = (code && prodByCode[code] !== undefined) ? prodByCode[code]
               : (prodByName[name] !== undefined)         ? prodByName[name]
               : null;

    if (rate === null) {
      incValues.push([row[iInc]]); // 商材未登録は変更しない
      skipped++;
      continue;
    }

    const unitSales = Number(row[iSales])   || 0;
    const unitCost  = Number(row[iCost])    || 0;
    const courses   = Math.max(1, Number(row[iCourses]) || 1);
    const qty       = Math.max(1, Number(row[iQty])     || 1);
    const months    = Math.max(1, Number(row[iMonths])  || 1);
    const monthlyGP = (unitSales - unitCost) * courses * qty;
    incValues.push([Math.floor(monthlyGP * months * rate)]);
    updated++;
  }

  if (incValues.length) {
    dealSheet.getRange(2, iInc + 1, incValues.length, 1).setValues(incValues);
  }

  invalidateAllDataCache_();
  Logger.log('設定_商材 レート更新完了 / 案件マスタ再計算: '
    + updated + '件更新 / ' + skipped + '件スキップ（商材未登録） / キャッシュクリア済み');
}

// ============================================================
// Excelインポート時のDate型ゴミ値（費用が-2兆等）を修正
// 費用（単価）が負値 or インセンティブが100万超の行を0にリセット
// GASエディタから一度だけ実行する
// ============================================================
function fixBadCostData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const iCost = h.indexOf('費用（単価）');
  const iInc  = h.indexOf('インセンティブ');

  const iSales = h.indexOf('売上（単価）');
  const costCol = [], incCol = [], salesCol = [];
  let fixed = 0;

  for (let i = 1; i < data.length; i++) {
    const sales = Number(data[i][iSales]) || 0;
    const cost  = Number(data[i][iCost])  || 0;
    const inc   = Number(data[i][iInc])   || 0;
    // 売上・費用が負値または1億超（Date型ゴミ値）、インセンティブが絶対値100万超は異常値
    const badSales = sales < 0 || sales > 100000000;
    const badCost  = cost  < 0 || cost  > 100000000;
    const badInc   = Math.abs(inc) > 1000000;
    salesCol.push([badSales ? 0 : data[i][iSales]]);
    costCol.push( [badCost  ? 0 : data[i][iCost]]);
    incCol.push(  [badInc   ? 0 : data[i][iInc]]);
    if (badSales || badCost || badInc) fixed++;
  }

  if (costCol.length) {
    sheet.getRange(2, iSales + 1, salesCol.length, 1).setValues(salesCol);
    sheet.getRange(2, iCost  + 1, costCol.length,  1).setValues(costCol);
    sheet.getRange(2, iInc   + 1, incCol.length,   1).setValues(incCol);
  }

  invalidateAllDataCache_();
  Logger.log('fixBadCostData: ' + fixed + '件修正（費用0・インセンティブ0にリセット） / キャッシュクリア済み');
}
