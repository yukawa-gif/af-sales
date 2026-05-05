// ============================================================
// 顧客マスタ・担当者マスタ関連 GAS関数
// gas_v3.js に追記して使用する
// ============================================================

const CUSTOMER_SHEET = '顧客マスタ';
const CONTACT_SHEET  = '担当者マスタ';

// ── ID採番ユーティリティ ──────────────────────────────────────
function getNextId_(sheet, prefix) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    if (id.startsWith(prefix + '-')) {
      const n = parseInt(id.split('-')[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  return prefix + '-' + String(maxNum + 1).padStart(4, '0');
}

// ============================================================
// 顧客マスタ CRUD
// 列: A:顧客ID B:企業名 C:都道府県 D:住所 E:業種 F:AFC担当者
//     G:顧客ステータス H:登録日 I:最終更新日 J:メモ
// ============================================================

function getAllCustomers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMER_SHEET);
  if (!sheet) return [];
  const [, ...rows] = sheet.getDataRange().getValues();
  return rows.filter(r => r[0]).map(r => ({
    id:       r[0],
    company:  r[1],
    pref:     r[2],
    address:  r[3],
    industry: r[4],
    afcStaff: r[5],
    status:   r[6],
    memo:     r[9]
  }));
}

function getCustomers(d) {
  const page     = parseInt((d && d.page != null) ? d.page : 0, 10);
  const pageSize = parseInt((d && d.pageSize)     ? d.pageSize : 50, 10);
  const query    = d && d.query    ? String(d.query).toLowerCase()    : '';
  const status   = d && d.status   ? String(d.status)                 : '';
  const afcStaff = d && d.afcStaff ? String(d.afcStaff)              : '';

  const all = getAllCustomers_();
  const filtered = all.filter(c => {
    if (status   && c.status   !== status)   return false;
    if (afcStaff && c.afcStaff !== afcStaff) return false;
    if (query) {
      const co = String(c.company  || '').toLowerCase();
      const st = String(c.afcStaff || '').toLowerCase();
      if (!co.includes(query) && !st.includes(query)) return false;
    }
    return true;
  });

  const total = filtered.length;
  const data  = filtered.slice(page * pageSize, (page + 1) * pageSize);
  return { success: true, data, total, page, pageSize };
}

function getCustomerStats() {
  const all      = getAllCustomers_();
  const total    = all.length;
  const active   = all.filter(c => c.status === '取引中' || c.status === '既存').length;
  const prospect = all.filter(c => c.status === '見込み').length;
  const inactive = all.filter(c => c.status === '休眠'   || c.status === '失注').length;
  return { success: true, total, active, prospect, inactive };
}

function addCustomer(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMER_SHEET);
  if (!sheet) throw new Error('顧客マスタシートが見つかりません');
  const id  = getNextId_(sheet, 'CUS');
  const now = new Date();
  sheet.appendRow([id, d.company||'', d.pref||'', d.address||'', d.industry||'',
                   d.afcStaff||'', d.status||'見込み', now, now, d.memo||'']);
  return { success: true, id };
}

function updateCustomer(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMER_SHEET);
  if (!sheet) throw new Error('顧客マスタシートが見つかりません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== d.id) continue;
    const r = i + 1;
    if (d.company  !== undefined) sheet.getRange(r, 2).setValue(d.company);
    if (d.pref     !== undefined) sheet.getRange(r, 3).setValue(d.pref);
    if (d.address  !== undefined) sheet.getRange(r, 4).setValue(d.address);
    if (d.industry !== undefined) sheet.getRange(r, 5).setValue(d.industry);
    if (d.afcStaff !== undefined) sheet.getRange(r, 6).setValue(d.afcStaff);
    if (d.status   !== undefined) sheet.getRange(r, 7).setValue(d.status);
    if (d.memo     !== undefined) sheet.getRange(r, 10).setValue(d.memo);
    sheet.getRange(r, 9).setValue(new Date());
    return { success: true };
  }
  throw new Error('顧客ID ' + d.id + ' が見つかりません');
}

// ============================================================
// 担当者マスタ CRUD
// 列: A:担当者ID B:顧客ID C:企業名 D:担当者名 E:役職
//     F:メールアドレス G:電話番号 H:メール配信可否 I:登録日 J:メモ
//     K:主連絡手段 L:SNS_ID
// ============================================================

function getContacts(customerId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTACT_SHEET);
  if (!sheet) return [];
  const [, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter(r => r[0] && (!customerId || r[1] === customerId))
    .map(r => ({
      id:           r[0],
      customerId:   r[1],
      company:      r[2],
      name:         r[3],
      title:        r[4],
      email:        r[5],
      phone:        r[6],
      mailOk:       r[7],
      createdAt:    r[8] ? Utilities.formatDate(new Date(r[8]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      memo:         r[9],
      primaryChannel: r[10],
      snsId:        r[11]
    }));
}

function addContact(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTACT_SHEET);
  if (!sheet) throw new Error('担当者マスタシートが見つかりません');

  // 顧客マスタから企業名を補完
  let company = d.company || '';
  if (!company && d.customerId) {
    const found = getAllCustomers_().find(c => c.id === d.customerId);
    if (found) company = found.company;
  }

  const id  = getNextId_(sheet, 'CON');
  const now = new Date();
  sheet.appendRow([
    id, d.customerId||'', company, d.name||'', d.title||'',
    d.email||'', d.phone||'', d.mailOk !== undefined ? d.mailOk : '○',
    now, d.memo||'',
    d.primaryChannel||'メール', d.snsId||''
  ]);
  return { success: true, id };
}

function updateContact(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTACT_SHEET);
  if (!sheet) throw new Error('担当者マスタシートが見つかりません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== d.id) continue;
    const r = i + 1;
    if (d.name           !== undefined) sheet.getRange(r,  4).setValue(d.name);
    if (d.title          !== undefined) sheet.getRange(r,  5).setValue(d.title);
    if (d.email          !== undefined) sheet.getRange(r,  6).setValue(d.email);
    if (d.phone          !== undefined) sheet.getRange(r,  7).setValue(d.phone);
    if (d.mailOk         !== undefined) sheet.getRange(r,  8).setValue(d.mailOk);
    if (d.memo           !== undefined) sheet.getRange(r, 10).setValue(d.memo);
    if (d.primaryChannel !== undefined) sheet.getRange(r, 11).setValue(d.primaryChannel);
    if (d.snsId          !== undefined) sheet.getRange(r, 12).setValue(d.snsId);
    return { success: true };
  }
  throw new Error('担当者ID ' + d.id + ' が見つかりません');
}

// ============================================================
// 旧「顧客DB」→ 新「顧客マスタ」「担当者マスタ」自動移行
// 旧列: A:顧客コード B:AFC営業担当 C:企業名 D:担当者名
//       E:担当者電話番号 F:担当者メールアドレス G:都道府県 H:住所
// ============================================================
function migrateOldCustomerDB() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet    = ss.getSheetByName('顧客DB');
  const cusSheet    = ss.getSheetByName(CUSTOMER_SHEET);
  const conSheet    = ss.getSheetByName(CONTACT_SHEET);

  if (!oldSheet) throw new Error('顧客DBシートが見つかりません');
  if (!cusSheet) throw new Error('顧客マスタシートが見つかりません');
  if (!conSheet) throw new Error('担当者マスタシートが見つかりません');

  const [, ...rows]  = oldSheet.getDataRange().getValues();
  const existingCompanies = getAllCustomers_().map(c => c.company);
  let count = 0;

  for (const r of rows) {
    const company = String(r[2] || '').trim();
    if (!company) continue;
    if (existingCompanies.includes(company)) continue;

    const cusId = getNextId_(cusSheet, 'CUS');
    const now   = new Date();

    cusSheet.appendRow([
      cusId,
      company,
      r[6] || '',   // 都道府県
      r[7] || '',   // 住所
      '',           // 業種（未設定・後で手動 or Gemini判定）
      r[1] || '',   // AFC担当者
      '既存',
      now,
      now,
      ''
    ]);
    existingCompanies.push(company);

    // 担当者情報があれば担当者マスタにも追加
    const hasContact = String(r[3] || '').trim() || String(r[5] || '').trim();
    if (hasContact) {
      const conId = getNextId_(conSheet, 'CON');
      conSheet.appendRow([
        conId,
        cusId,
        company,
        r[3] || '',  // 担当者名
        '',          // 役職
        r[5] || '',  // メールアドレス
        r[4] || '',  // 電話番号
        '○',
        now,
        ''
      ]);
    }
    count++;
  }

  return { success: true, migrated: count };
}

// ============================================================
// Gemini: 企業名 → 業種自動判定
// ============================================================
function detectIndustry(companyName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url    = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const INDUSTRIES = [
    'A 農業・林業','B 漁業','C 鉱業','D 建設業','E 製造業',
    'F 電気・ガス・水道','G 情報通信業','H 運輸・郵便業',
    'I 卸売・小売業','J 金融・保険業','K 不動産業',
    'L 学術・専門サービス業','M 宿泊・飲食業','N 生活サービス・娯楽',
    'O 教育・学習支援','P 医療・福祉','Q 複合サービス',
    'R その他サービス業','S 公務','T 分類不能'
  ];

  const prompt =
    `企業名から日本標準産業分類（大分類）を1つ推定してください。\n` +
    `企業名: ${companyName}\n\n` +
    `選択肢:\n${INDUSTRIES.join('\n')}\n\n` +
    `JSONで回答: {"industry":"選択した分類（例: G 情報通信業）","confidence":"high/medium/low"}`;

  const res  = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json' }
    })
  });
  const body = JSON.parse(res.getContentText());
  return JSON.parse(body.candidates[0].content.parts[0].text);
}

// ============================================================
// 顧客・担当者 削除
// ============================================================

function deleteCustomer(id) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cusSheet = ss.getSheetByName(CUSTOMER_SHEET);
  const conSheet = ss.getSheetByName(CONTACT_SHEET);
  if (!cusSheet) throw new Error('顧客マスタシートが見つかりません');

  // 顧客行を削除
  const cusData = cusSheet.getDataRange().getValues();
  for (let i = 1; i < cusData.length; i++) {
    if (cusData[i][0] === id) { cusSheet.deleteRow(i + 1); break; }
  }

  // 紐づく担当者を全削除
  if (conSheet) {
    const conData = conSheet.getDataRange().getValues();
    for (let i = conData.length - 1; i >= 1; i--) {
      if (conData[i][1] === id) conSheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

function deleteContact(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTACT_SHEET);
  if (!sheet) throw new Error('担当者マスタシートが見つかりません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  throw new Error('担当者ID ' + id + ' が見つかりません');
}

// ============================================================
// 既存 doGet / doPost の switch-case に以下を追加してください
// ============================================================
//
//   case 'getCustomers':    return ok(getCustomers());
//   case 'addCustomer':     return ok(addCustomer(d));
//   case 'updateCustomer':  return ok(updateCustomer(d));
//   case 'deleteCustomer':  return ok(deleteCustomer(d.id));
//   case 'getContactMaster':return ok(getContacts(d.customerId));
//   case 'addContact':      return ok(addContact(d));
//   case 'updateContact':   return ok(updateContact(d));
//   case 'deleteContact':   return ok(deleteContact(d.id));
//   case 'detectIndustry':  return ok(detectIndustry(d.companyName));
//   case 'migrateOldDB':    return ok(migrateOldCustomerDB());
