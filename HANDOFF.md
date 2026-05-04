# AF Sales SFA システム 引き継ぎ文書

## プロジェクト概要

エー・ファクトリー株式会社の営業管理SFAシステム。

| 項目 | 内容 |
|------|------|
| リポジトリ | https://github.com/yukawa-gif/af-sales.git |
| 本番URL | https://af-sales.vercel.app |
| デプロイ | GitHub main push → Vercel 自動反映 |
| GAS URL | `https://script.google.com/macros/s/AKfycbw_Hx1yajz8zFkpjoTs3YdAR3AdSPHG8VGDzg5h5xiZUid80rC1gSq-uRPX6BAZJOM/exec` |
| スプレッドシート | https://docs.google.com/spreadsheets/d/1eW1tYD_FYih6w4HYpgayXk2JO0r3UEfXFQNo1jcfm28/edit |

## ファイル構成

| ファイル | 内容 |
|----------|------|
| `dashboard.html` | メインダッシュボード |
| `deal_form.html` | 案件登録フォーム |
| `daily_report.html` | 日報入力フォーム |
| `customer_master.html` | 顧客管理画面 |
| `gas_v3.js` | GASバックエンド（GASエディタで管理・再デプロイ必要） |
| `gas_customer_master.js` | 顧客マスタGAS関数 |

## 技術的な注意事項

- GAS を変更したら必ず「デプロイを管理→新しいバージョン→デプロイ」で再デプロイ
- ローカル動作確認：`python -m http.server 8080` → `http://localhost:8080/`
- FY（会計年度）は8月始まり（FY2025 = 2025/8〜2026/7）
- 担当者コード P（湯川悦英）は目標・インセンティブ計算の対象外
- Google Workspace 契約済み（カレンダー予約スケジュール機能が使える）

---

## 次回セッションでやること：MA機能実装

### 背景

顧客管理画面（`customer_master.html`）にMA（マーケティングオートメーション）機能を追加する。
業種・都道府県単位で顧客担当者にマーケティングメールを送りたい。

目的：
1. 担当者引き継ぎ管理
2. 顧客ステータス管理
3. **MAメール配信**（今回実装）

### 実装フロー

1. **絞り込み** — 業種・都道府県・ステータスで複合フィルタ、「メール配信○」の担当者のみ抽出
2. **ターゲット選択** — チェックボックスで送信先を選ぶ
3. **メール生成** — Geminiに業種・目的・担当者名を渡してメール文章を自動生成
4. **予約リンク挿入** — AFC担当者のGoogleカレンダー予約URLをメール本文に自動挿入
5. **送信** — GmailAppで下書き一括作成（誤送信防止のため送信ではなく下書き）

### Step 1：スプレッドシートの準備（手動作業・先にやる）

`設定_担当者` シートに **「予約ページURL」列を追加**し、各担当者のGoogleカレンダー予約URLを入力する。

予約URLの取得方法：
1. Googleカレンダーを開く
2. 左メニュー「他のカレンダー」→「＋」→「予約スケジュールを作成」
3. 設定後に発行される共有URLをコピー
4. 担当者ごとにスプレッドシートへ記入

### Step 2：GAS修正（`gas_v3.js` → GASエディタで反映・再デプロイ）

以下の関数を追加し、doPost にルーティングを追加する：

```javascript
// 担当者の予約URLを取得（設定_担当者シートから）
function getPersonCalendarUrl(personName) { ... }

// Geminiでマーケティングメール文章を生成
// 既存の draftFollowUp 関数（gemini-2.5-flash 使用）を参考に実装
function generateMaEmail({ industry, pref, purpose, senderName, calendarUrl }) { ... }

// Gmail下書きを一括作成
// GmailApp.createDraft(to, subject, body) を使う
function createMaDrafts({ contacts, subject, body }) { ... }
```

doPost に追加するアクション：
- `action: 'generateMaEmail'`
- `action: 'createMaDrafts'`

### Step 3：`customer_master.html` のUI修正

#### 3-1. 絞り込みパネルの強化

現在の検索バー横に以下を追加：
- 業種ドロップダウン
- 「メール配信○のみ」チェックボックス
（都道府県・ステータスは既存フィルタを活用）

#### 3-2. チェックボックスによる一括選択

顧客一覧テーブルの左端にチェックボックス列を追加。
「全選択」ボタンも用意。
選択するのは「顧客」ではなく、その顧客に紐づく「担当者（メール配信○）」。

#### 3-3. MAパネル（チェック後に表示）

選択後、画面下部にパネルを表示：

```
[選択中: 12件の担当者]
送信AFC担当者: [ドロップダウン]
目的: [新規商品紹介 / フォロー / キャンペーン / 自由入力]
[✨ Geminiでメール生成]
─────────────────────
件名: [編集可能テキスト]
本文: [編集可能テキスト（{{予約リンク}} が自動挿入）]
─────────────────────
[📧 Gmailで下書き作成]
```

### 既存コードの参考箇所

- Gemini API 呼び出し例：`gas_v3.js` の `draftFollowUp()` 関数（gemini-2.5-flash 使用）
- 担当者データ取得：`customer_master.html` の `loadAllContacts()` 関数
- メール配信可否フィールド：担当者モーダルの `conMailOk`（○/×）

---

## その他の残作業

- **データ統一作業**：全担当者のExcelデータを9項目に統一 → `importFromSheet()` で一括インポート
- **新・顧客マスタの構築**：案件マスタを中心に顧客データを再設計（顧客ID紐づけ）
