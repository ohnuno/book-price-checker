# 古本買取価格調査システム

**最終更新日**: 2025年12月

---

## 📚 目次

1. [Google Spreadsheetの使い方](#google-spreadsheetの使い方)
2. [Google Apps Scriptのメンテナンス](#google-apps-scriptのメンテナンス)
3. [Google Cloud Platformのメンテナンス](#google-cloud-platformのメンテナンス)
4. [システム管理者向け情報](#システム管理者向け情報)

---

## Google Spreadsheetの使い方

### シートの説明

#### 📊 ダッシュボード

システム全体の状況を可視化するシート。手動で編集する必要はありません。

**表示内容:**
- 現在の状況（登録書籍数、更新状況、価格統計）
- 買取実績（総冊数、総利益、平均利益）
- 今月の実績
- 価格変動アラート（上昇TOP5、下落TOP5、0円書籍）
- 高利益書籍ランキング（TOP10）
- 実施中のキャンペーン情報

**更新方法:**
- メニュー「古本買取システム」→「ダッシュボードを更新」で手動更新
- 買取完了シートへの移行時に自動更新

---

#### 📖 ISBNリスト

買取前の書籍を管理するメインシート。

**列構成:**
- A列: ISBN（13桁または10桁）
- B列: 書籍名（自動取得）
- C列: 著者（自動取得）
- D列: 初回見積価格（自動記録）
- E列: 最新見積価格（毎日自動更新）
- F列: 価格更新日時（自動記録）
- G列: 価格増減（自動計算）
- H列: チェックボックス（買取完了への移行用）

**使い方:**
1. A列にISBNを入力
2. 自動的に書籍情報が取得される（Google Books API）
3. 毎日深夜0時に価格が自動更新される（GCP）
4. 買取完了時にH列のチェックボックスにチェック

---

#### 📈 価格履歴

各書籍の価格変動履歴を記録するシート。

**列構成:**
- A列: ISBN
- B列: 書籍名
- C列: 更新日時
- D列: 価格
- E列: 変動額

**特徴:**
- 価格が変動した場合のみ記録される
- 買取完了時に該当ISBNの履歴は自動削除される

---

#### ⚠️ エラーログ

価格更新処理の実行結果を記録するシート。

**列構成:**
- A列: 実行日時
- B列: 処理件数
- C列: 成功件数
- D列: 失敗件数
- E列: 成功率
- F列: 失敗ISBN

**確認方法:**
- 毎日1行ずつ増えていることを確認
- 成功率が低い場合はシステム異常の可能性

---

#### ✅ 買取完了シート（複数）

実際に買取した書籍を記録するシート。日付ごとに自動作成されます。

**シート名形式:** `買取完了_YYYY-MM-DD`

**列構成:**
- A列: ISBN
- B列: タイトル
- C列: 著者
- D列: 出版社
- E列: 最新見積価格
- F列: 売却価格（手動入力）
- G列: 利益（自動計算: F列 - E列）
- H列: 登録日

**使い方:**
1. ISBNリストから移行時に自動作成
2. F列（売却価格）を手動で入力
3. G列（利益）が自動計算される

---

### 作業フロー

#### 1. ISBNの入力

```
1. ISBNリストシートのA列に13桁または10桁のISBNを入力
2. 数秒待つと自動的に書籍情報（タイトル、著者、出版社）が取得される
3. 翌日の深夜0時に買取価格が自動更新される
```

**手動で書籍情報を取得する場合:**
- メニュー「書籍情報」→「選択範囲のISBNを処理」
- 複数行を選択して一括処理可能

---

#### 2. チェックボックスの追加

ISBNを入力すると自動的にH列にチェックボックスが追加されます。

**手動で追加する場合:**
```
メニュー「書籍情報」→「全行にチェックボックス設定」
```

---

#### 3. ISBNリストから買取完了シートへの移行

```
1. ISBNリストシートのH列で買取完了した書籍にチェックを入れる
2. メニュー「古本買取システム」→「買取完了に移行」をクリック
3. ダイアログで記録先シートを選択:
   - 既存のシート（直近1ヶ月以内）を選ぶ
   - または新しいシート（当日の日付）を作成
4. 確認して実行
```

**実行結果:**
- チェックした書籍が選択したシートに移動
- ISBNリストから該当行が削除される
- 価格履歴から該当ISBNの履歴が削除される
- ダッシュボードが自動更新される

---

#### 4. 買取価格の記録

```
1. 買取完了シートを開く
2. F列（売却価格）に実際の買取価格を入力
3. G列（利益）が自動計算される
```

**例:**
- E列（見積価格）: 150円
- F列（売却価格）: 180円 ← 手動入力
- G列（利益）: 30円 ← 自動計算

---

### 特殊メニューの説明

#### 📚 書籍情報メニュー

##### 📝 全ISBNの情報を取得

**機能:** ISBNリスト全体に対してGoogle Books APIで書籍情報を取得

**使用タイミング:**
- ISBNを大量に入力した後
- 自動取得に失敗した書籍がある場合

**実行方法:**
```
1. メニュー「書籍情報」→「全ISBNの情報を取得」
2. 確認ダイアログで「はい」をクリック
3. 処理完了まで待機（10件ごとに進捗表示）
```

**注意事項:**
- 既に情報がある行はスキップされる
- 大量のISBN（100件以上）がある場合は時間がかかる

---

##### 🔄 選択範囲のISBNを処理

**機能:** 選択した範囲のISBNのみ処理

**使用タイミング:**
- 特定の書籍だけ情報を取得したい場合
- 一部のISBNで取得失敗があった場合

**実行方法:**
```
1. ISBNリストシートでA列の処理したい範囲を選択
2. メニュー「書籍情報」→「選択範囲のISBNを処理」
3. 処理完了まで待機
```

---

##### ✅ 全行にチェックボックス設定

**機能:** ISBNがある全行のH列にチェックボックスを追加

**使用タイミング:**
- チェックボックスが消えてしまった場合
- 古いデータを移行した場合

**実行方法:**
```
1. メニュー「書籍情報」→「全行にチェックボックス設定」
2. 完了通知を確認
```

---

#### 📚 古本買取システムメニュー

##### ✅ 買取完了に移行

**機能:** チェックした書籍を買取完了シートに移動

**詳細:** [3. ISBNリストから買取完了シートへの移行](#3-isbnリストから買取完了シートへの移行) を参照

---

##### 📊 ダッシュボードをセットアップ

**機能:** ダッシュボードシートを初期化・再構築

**使用タイミング:**
- 初回セットアップ時
- ダッシュボードが破損した場合
- レイアウトを最新版に更新する場合

**実行方法:**
```
1. メニュー「古本買取システム」→「ダッシュボードをセットアップ」
2. 処理完了まで待機（数秒）
3. ダッシュボードシートを確認
```

**注意事項:**
- 既存のダッシュボードシートは上書きされる
- データ自体は削除されない（ISBNリスト、買取完了シートは無影響）

---

##### 🔄 ダッシュボードを更新

**機能:** ダッシュボードの全データを強制再計算

**使用タイミング:**
- ダッシュボードの数値が正しく更新されない場合
- 買取完了シートを追加・編集した後
- 統計情報を最新化したい場合

**実行方法:**
```
1. メニュー「古本買取システム」→「ダッシュボードを更新」
2. 処理完了まで待機（数秒）
```

**技術詳細:**
- 全カスタム関数のキャッシュをクリア
- 約100個のセルの数式を再設定
- 複数の買取完了シートがある場合に特に有効

---

##### 🎁 キャンペーン情報を更新

**機能:** charibon.jpから最新のキャンペーン情報を取得してダッシュボードに表示

**使用タイミング:**
- 新しいキャンペーンが始まった時
- キャンペーン情報を手動で確認したい時

**実行方法:**
```
1. メニュー「古本買取システム」→「キャンペーン情報を更新」
2. 処理完了まで待機
3. ダッシュボードのRow 28以降を確認
```

**表示内容:**
- キャンペーンタイトル
- バナー画像
- 内容、期間、対象

**注意事項:**
- 期間外のキャンペーンは表示されない
- charibon.jpのHTML構造が変更されると動作しなくなる可能性あり

---

## Google Apps Scriptのメンテナンス

### シートの追加・修正・削除時の注意事項

#### シート名を変更する場合

```javascript
// Config.gs の SHEET_NAMES を修正
const CONFIG = {
  SHEET_NAMES: {
    ISBN_LIST: 'ISBNリスト',        // ← ここを変更
    COMPLETED: '買取完了',
    PRICE_HISTORY: '価格履歴',
    DASHBOARD: 'ダッシュボード',
    ERROR_LOG: 'エラーログ'
  },
  // ...
}
```

**影響範囲:**
- すべての`.gs`ファイルでこの定数を参照している
- 変更後は全機能をテストすること

---

#### 列を追加・削除する場合

**ISBNリストシートの列変更:**

```javascript
// Config.gs の ISBN_LIST_COLUMNS を修正
ISBN_LIST_COLUMNS: {
  ISBN: 1,           // A列
  TITLE: 2,          // B列
  AUTHOR: 3,         // C列
  PUBLISHER: 4,      // D列
  PRICE: 5,          // E列
  UPDATED: 6,        // F列
  PRICE_CHANGE: 7,   // G列
  CHECKBOX: 8        // H列（新しい列を追加する場合は9以降）
}
```

**重要:** GCPのPythonコード（book_price_fetcher.py）も同時に修正が必要

---

#### 買取完了シートの列変更:

```javascript
// Config.gs の COMPLETED_COLUMNS を修正
COMPLETED_COLUMNS: {
  ISBN: 1,           // A列
  TITLE: 2,          // B列
  AUTHOR: 3,         // C列
  PUBLISHER: 4,      // D列
  ESTIMATE: 5,       // E列
  ACTUAL: 6,         // F列
  DIFFERENCE: 7,     // G列（新しい列を追加する場合は8以降）
  DATE: 8            // H列
}
```

---

### トリガー一覧

#### 設定済みトリガー

システムで使用しているトリガー（要手動設定）:

| トリガー名 | 関数名 | イベント | 説明 |
|---|---|---|---|
| スプレッドシート起動時 | `onOpen` | スプレッドシートを開いたとき | カスタムメニューを追加 |

#### 推奨トリガー（オプション）

必要に応じて設定:

| トリガー名 | 関数名 | イベント | 頻度 | 説明 |
|---|---|---|---|---|
| 毎日キャンペーン更新 | `dailyCampaignUpdate` | 時間主導型 | 日次（午前6時） | キャンペーン情報を毎日自動更新 |

---

### トリガーの設定方法

```
1. Apps Scriptエディタを開く
2. 左側メニューの「トリガー」（時計アイコン）をクリック
3. 右下の「トリガーを追加」をクリック
4. 以下を設定:
   - 実行する関数: 該当する関数名を選択
   - イベントのソース: 「スプレッドシートから」または「時間主導型」
   - イベントの種類: 「起動時」または「日タイマー」など
   - 時刻: 必要に応じて設定
5. 「保存」をクリック
6. 権限の承認を求められたら「承認」
```

---

### スクリプトプロパティの設定

#### Google Books API キーの設定

```
1. Apps Scriptエディタを開く
2. 左側メニューの「プロジェクトの設定」（歯車アイコン）をクリック
3. 「スクリプト プロパティ」セクションで「スクリプト プロパティを追加」
4. 以下を入力:
   - プロパティ: GOOGLE_BOOKS_API_KEY
   - 値: （取得したAPIキー）
5. 「スクリプト プロパティを保存」をクリック
```

#### 通知先メールアドレスの設定

```
1. 同じく「スクリプト プロパティを追加」
2. 以下を入力:
   - プロパティ: NOTIFICATION_EMAIL
   - 値: your-email@example.com
3. 「スクリプト プロパティを保存」をクリック
```

---

### コードのバックアップ

**推奨頻度:** 月1回または大きな変更後

```
1. Apps Scriptエディタを開く
2. 各.gsファイルの内容をコピー
3. ローカルのテキストエディタに保存
4. ファイル名: YYYY-MM-DD_ファイル名.gs
```

**または:**

```
1. Apps Scriptエディタで「プロジェクトの設定」
2. 「Google Apps Script API」を有効化
3. claspツールでローカルに管理（上級者向け）
```

---

## Google Cloud Platformのメンテナンス

### 環境構築方法

#### 前提条件

- Googleアカウント
- クレジットカード（無料枠内で使用可能だが登録必須）
- 基本的なターミナル操作の知識

---

#### 1. GCPプロジェクトの作成

```
1. https://console.cloud.google.com/ にアクセス
2. 右上の「プロジェクトを作成」をクリック
3. プロジェクト名: book-price-checker（任意）
4. 「作成」をクリック
5. 作成したプロジェクトを選択
```

---

#### 2. 必要なAPIの有効化

```
1. 左側メニュー「APIとサービス」→「ライブラリ」
2. 以下のAPIを検索して有効化:
   - Cloud Functions API
   - Cloud Build API
   - Cloud Logging API
3. それぞれ「有効にする」をクリック
```

---

#### 3. サービスアカウントの作成

```
1. 左側メニュー「IAMと管理」→「サービスアカウント」
2. 「サービスアカウントを作成」をクリック
3. 以下を入力:
   - サービスアカウント名: book-price-scraper
   - ID: 自動生成されたものを使用
4. 「作成して続行」をクリック
5. ロールを選択:
   - Cloud Functions 起動者
   - Logs 書き込み
6. 「完了」をクリック
```

---

#### 4. 認証情報ファイルの作成

```
1. 作成したサービスアカウントをクリック
2. 「キー」タブを選択
3. 「鍵を追加」→「新しい鍵を作成」
4. キーのタイプ: JSON
5. 「作成」をクリック
6. ダウンロードされたJSONファイルを保存
7. ファイル名を credentials.json に変更
```

**重要:** このファイルは絶対に公開しないこと

---

#### 5. Google Cloud SDKのインストール

**Windows:**
```powershell
# インストーラーをダウンロード
# https://cloud.google.com/sdk/docs/install

# インストール後、PowerShellを再起動
gcloud --version
```

**Mac/Linux:**
```bash
# Homebrewでインストール（Mac）
brew install google-cloud-sdk

# または公式スクリプト
curl https://sdk.cloud.google.com | bash

# 設定を反映
exec -l $SHELL

# バージョン確認
gcloud --version
```

---

#### 6. gcloudの初期化

```bash
# 初期化
gcloud init

# ログイン（ブラウザが開く）
# → Googleアカウントでログイン

# プロジェクトを選択
# → 作成したプロジェクト（book-price-checker）を選択

# デフォルトのリージョンを設定
gcloud config set functions/region asia-northeast1
```

---

#### 7. プロジェクトディレクトリの準備

```bash
# プロジェクトディレクトリを作成
mkdir ~/book-price-checker-gcp
cd ~/book-price-checker-gcp

# 必要なファイルを配置
# - book_price_fetcher.py
# - main.py
# - requirements.txt
# - credentials.json
```

**ディレクトリ構造:**
```
~/book-price-checker-gcp/
├── book_price_fetcher.py
├── main.py
├── requirements.txt
└── credentials.json
```

---

#### 8. 初回デプロイ

```bash
# プロジェクトディレクトリに移動
cd ~/book-price-checker-gcp

# デプロイ（初回は10分程度かかる）
gcloud functions deploy update_prices \
  --gen2 \
  --runtime python311 \
  --trigger-http \
  --allow-unauthenticated \
  --entry-point update_prices \
  --source . \
  --region asia-northeast1 \
  --memory 2GB \
  --timeout 540s \
  --set-env-vars SPREADSHEET_ID=あなたのスプレッドシートID
```

**スプレッドシートIDの確認方法:**
```
https://docs.google.com/spreadsheets/d/【このID部分】/edit
```

---

#### 9. Cloud Schedulerの設定（毎日自動実行）

```bash
# Cloud Scheduler APIを有効化
gcloud services enable cloudscheduler.googleapis.com

# スケジュールを作成（毎日午前0時に実行）
gcloud scheduler jobs create http daily-price-update \
  --location asia-northeast1 \
  --schedule="0 0 * * *" \
  --uri="https://asia-northeast1-あなたのプロジェクトID.cloudfunctions.net/update_prices" \
  --http-method GET \
  --time-zone="Asia/Tokyo"
```

**関数のURLの確認方法:**
```bash
gcloud functions describe update_prices \
  --region asia-northeast1 \
  --gen2 \
  --format="value(serviceConfig.uri)"
```

---

### コードの修正方法

#### 1. ローカルでファイルを編集

```bash
# プロジェクトディレクトリに移動
cd ~/book-price-checker-gcp

# ファイルを編集（任意のエディタ）
nano book_price_fetcher.py
# または
code book_price_fetcher.py  # VS Code
```

---

#### 2. 修正内容の確認

**主な修正ポイント:**

##### 列番号の変更（ISBNリストシートの構造変更時）

```python
# book_price_fetcher.py の update_spreadsheet() 内

# 例: 価格増減の列を変更
sheet.update_cell(i, 7, change)  # Column G (価格増減)
                 ↑
                 この数字を変更
```

**列番号の対応:**
- 1 = A列
- 2 = B列
- 3 = C列
- 以下同様

---

##### 処理件数の変更

```python
# book_price_fetcher.py の update_spreadsheet() 内

MAX_PROCESS_COUNT = 10  # ← この値を変更（1日あたりの処理件数）
```

**注意:** 大きくしすぎるとメモリエラーやタイムアウトの可能性

---

##### タイムアウト時間の変更

```python
# 現在は不要（デプロイ時のパラメータで設定）
```

デプロイコマンドの `--timeout 540s` を変更

---

#### 3. 再デプロイ

```bash
# プロジェクトディレクトリに移動
cd ~/book-price-checker-gcp

# 変更をデプロイ
gcloud functions deploy update_prices \
  --gen2 \
  --runtime python311 \
  --trigger-http \
  --allow-unauthenticated \
  --entry-point update_prices \
  --source . \
  --region asia-northeast1 \
  --memory 2GB \
  --timeout 540s \
  --set-env-vars SPREADSHEET_ID=あなたのスプレッドシートID
```

**注意:** 
- `SPREADSHEET_ID`は毎回指定が必要
- 2回目以降は3-5分で完了

---

#### 4. 動作確認

##### 手動実行テスト

```bash
# 関数を手動実行
gcloud functions call update_prices \
  --region asia-northeast1 \
  --gen2
```

##### ログの確認

```bash
# 最新のログを表示
gcloud functions logs read update_prices \
  --region asia-northeast1 \
  --gen2 \
  --limit 100
```

**または:** Cloud Consoleで確認
```
1. https://console.cloud.google.com/ にアクセス
2. 左側メニュー「Cloud Functions」
3. update_prices をクリック
4. 「ログ」タブを選択
```

---

### トラブルシューティング

#### デプロイエラー

**エラー:** `Permission denied`

**解決方法:**
```bash
# プロジェクトIDを確認
gcloud config get-value project

# 権限を確認
gcloud projects get-iam-policy プロジェクトID
```

---

**エラー:** `API not enabled`

**解決方法:**
```bash
# 必要なAPIを有効化
gcloud services enable cloudfunctions.googleapis.com
gcloud services enable cloudbuild.googleapis.com
```

---

#### 実行時エラー

**エラー:** `SPREADSHEET_ID environment variable not set`

**解決方法:**
```bash
# 環境変数を再設定してデプロイ
gcloud functions deploy update_prices \
  --gen2 \
  --runtime python311 \
  --trigger-http \
  --allow-unauthenticated \
  --entry-point update_prices \
  --source . \
  --region asia-northeast1 \
  --memory 2GB \
  --timeout 540s \
  --set-env-vars SPREADSHEET_ID=正しいスプレッドシートID
```

---

**エラー:** `Memory limit exceeded`

**解決方法:**
```bash
# メモリを増やしてデプロイ（最大8GB）
--memory 4GB  # 2GB → 4GB に変更
```

---

**エラー:** `Function execution took too long`

**解決方法:**
```bash
# タイムアウトを延長（最大540秒）
--timeout 540s
```

または処理件数を減らす:
```python
MAX_PROCESS_COUNT = 5  # 10 → 5 に変更
```

---

### バックアップ

**コードのバックアップ（推奨頻度: 月1回）:**

```bash
# バックアップディレクトリを作成
mkdir -p ~/backups/book-price-checker

# 日付付きでバックアップ
cp -r ~/book-price-checker-gcp ~/backups/book-price-checker/backup-$(date +%Y%m%d)
```

---

## システム管理者向け情報

### GAS（Google Apps Script）ファイル構成

#### Config.gs

**役割:** システム全体の設定と定数定義

**主要な関数:**
- `getSpreadsheet()` - アクティブなスプレッドシートを取得
- `getSheet(sheetName)` - 指定したシート名のシートを取得
- `getScriptProperty(key)` - スクリプトプロパティを取得
- `setScriptProperty(key, value)` - スクリプトプロパティを設定
- `isValidISBN(isbn)` - ISBNの形式をチェック
- `formatCurrency(value)` - 数値を円表記にフォーマット
- `formatDateTime(date)` - 日時を文字列にフォーマット

**重要な定数:**
- `CONFIG.SHEET_NAMES` - 各シートの名前
- `CONFIG.ISBN_LIST_COLUMNS` - ISBNリストシートの列番号
- `CONFIG.COMPLETED_COLUMNS` - 買取完了シートの列番号
- `CONFIG.SCRAPING` - スクレイピング設定

**修正時の注意:**
- シート名や列番号を変更する場合は全体への影響を確認
- 他のファイルがこの定数を参照している

---

#### BuyCompleted.gs

**役割:** 買取完了機能とダッシュボード更新

**主要な関数:**
- `onOpen()` - カスタムメニューを追加（トリガー: 起動時）
- `refreshDashboard()` - ダッシュボードを強制更新
- `moveToBuyCompleted()` - チェックした書籍を買取完了シートに移行
- `selectTargetSheet(ss, bookCount)` - 記録先シートを選択するダイアログ
- `createNewBuySheet(ss, sheetName)` - 新しい買取完了シートを作成
- `deletePriceHistory(ss, isbn)` - 価格履歴から該当ISBNの行を削除
- `adjustColumnWidths(sheet)` - シートの列幅を自動調整

**処理フロー:**
```
1. ユーザーがチェックボックスをチェック
2. moveToBuyCompleted() 実行
3. selectTargetSheet() でシート選択
4. createNewBuySheet() で必要ならシート作成
5. データを移行
6. deletePriceHistory() で価格履歴削除
7. ISBNリストから行削除
```

---

#### Dashboard.gs

**役割:** ダッシュボード用のカスタム関数

**主要な関数:**
- `getTotalBuyCount()` - 全買取完了シートの総冊数を取得
- `getTotalProfit()` - 全買取完了シートの総利益を取得
- `getMaxProfit()` - 全買取完了シートの最高利益を取得
- `getMonthlyBuyCount()` - 今月の買取冊数を取得
- `getMonthlyProfit()` - 今月の利益を取得
- `getPriceIncreasesTop5()` - 価格上昇TOP5を取得
- `getPriceDecreasesTop5()` - 価格下落TOP5を取得
- `getZeroPriceBooks()` - 0円になった書籍を取得
- `getTopProfitBooks()` - 高利益書籍TOP10を取得

**カスタム関数の特徴:**
- スプレッドシートのセルから `=getTotalBuyCount()` のように呼び出せる
- 引数なしで動作（内部で全シートを検索）
- キャッシュされるため、`refreshDashboard()` で強制更新が必要

---

#### DashboardSetup.gs

**役割:** ダッシュボードシートの初期化と構築

**主要な関数:**
- `setupDashboardSheet()` - ダッシュボード全体をセットアップ
- `setupLayout(sheet)` - レイアウト（テキスト配置）を設定
- `setupFormulas(sheet)` - 数式を設定
- `setupFormatting(sheet)` - 書式（色、フォント、条件付き書式）を設定

**実行タイミング:**
- 初回セットアップ時
- ダッシュボードが破損した時
- レイアウトを変更した時

**修正時の注意:**
- `setupLayout()` でセルの配置を変更
- `setupFormulas()` で数式を変更
- `setupFormatting()` で見た目を変更
- 3つの関数は密接に関連しているため、整合性を保つこと

---

#### CampaignFetcher.gs

**役割:** charibon.jpからキャンペーン情報を取得

**主要な関数:**
- `fetchCampaignInfo()` - キャンペーン情報を取得してダッシュボードに書き込み
- `findLatestCampaignSection(html)` - ニュースページから最新のキャンペーン記事を検出
- `fetchCampaignDetailsFromArticle(url)` - 記事ページから詳細情報を取得
- `extractBannerImageFromArticle(html)` - バナー画像URLを抽出
- `extractCampaignInfoFromArticle(html)` - 内容・期間・対象を抽出
- `checkIfCampaignIsActive(periodText)` - 期間を判定（当日が期間内か）
- `writeCampaignToDashboard(campaign)` - ダッシュボードに書き込み
- `updateCampaignInfo()` - 手動更新用（メニューから呼び出し）
- `dailyCampaignUpdate()` - 自動更新用（時間主導型トリガー）

**Webスクレイピングの仕組み:**
```
1. UrlFetchApp.fetch() でHTMLを取得
2. 正規表現でHTML要素を抽出
3. <strong>タグの中身を取り出す
4. 日付パターンを解析
5. 当日が期間内かチェック
6. ダッシュボードに書き込み
```

**修正が必要になる可能性:**
- charibon.jpのHTML構造が変更された場合
- セレクタ（クラス名、タグ名）を更新する必要あり

---

#### GoogleBooksAPI.gs

**役割:** Google Books APIから書籍情報を取得

**主要な関数:**
- `fetchBookInfoFromGoogleBooks(isbn)` - ISBNから書籍情報を取得
- `createBookInfoMenu()` - 書籍情報メニューを作成（トリガー: 起動時）
- `processAllISBNs()` - 全ISBNに対して書籍情報を取得
- `processSelectedISBNs()` - 選択範囲のISBNを処理
- `addCheckboxesToAll()` - 全行にチェックボックスを設定
- `onEdit(e)` - ISBN入力時の自動処理（トリガー: 編集時）
- `ensureCheckbox(sheet, row)` - チェックボックスが未設定なら設定

**API呼び出しの流れ:**
```
1. ISBNから検索URL生成
2. UrlFetchApp.fetch() でAPI呼び出し
3. JSONレスポンスをパース
4. タイトル、著者、出版社を取得
5. スプレッドシートに書き込み
```

**API Keyの設定:**
- スクリプトプロパティ `GOOGLE_BOOKS_API_KEY` に設定
- 未設定でも動作するが、レート制限が厳しい

---

#### Utils.gs

**役割:** 汎用ユーティリティ関数

**主要な関数:**
- `fetchWithRetry(url, options)` - HTTPリクエストをリトライ機能付きで実行
- `extractTextFromHtml(html, pattern, groupIndex)` - HTMLからテキストを抽出
- `decodeHtmlEntities(text)` - HTMLエンティティをデコード
- `getSheetData(sheet)` - シートの全データを取得
- `getColumnValues(sheet, columnIndex)` - 指定列の値を取得
- `findRow(sheet, columnIndex, value)` - 条件に合致する行を検索
- `appendRow(sheet, data)` - 行を追加
- `getCheckedRows(sheet, checkboxColumn)` - チェックされている行を取得
- `getDaysAgo(days)` - 指定日数前の日付を取得
- `uniqueArray(array)` - 配列の重複を除去
- `sum(array)` - 数値配列の合計を計算
- `average(array)` - 数値配列の平均を計算

**用途:**
- 他のファイルから共通処理として呼び出される
- コードの重複を避けるためのヘルパー関数群

---

### GCP（Google Cloud Platform）ファイル構成

#### book_price_fetcher.py

**役割:** ValueBooks.jpから買取価格を取得してスプレッドシートを更新

**主要なクラス:**
- `ValueBooksScraper` - メインのスクレイパークラス

**主要なメソッド:**
- `__init__(credentials_file, headless)` - 初期化
- `_setup_driver(headless)` - Seleniumドライバーをセットアップ
- `_setup_google_sheets(credentials_file)` - Google Sheets APIをセットアップ
- `_get_jst_now()` - 現在の日本時間を取得
- `search_isbn_estimate(isbn)` - ISBNで買取見積を検索
- `_extract_estimate_result(isbn)` - 検索結果から書籍情報を抽出
- `update_spreadsheet(spreadsheet_id)` - スプレッドシートを更新（メイン処理）
- `_add_price_history(spreadsheet, book_info, previous_price)` - 価格履歴に記録
- `_write_execution_summary(spreadsheet, success_count, error_count, failed_isbns)` - エラーログに実行サマリを書き込み
- `close()` - リソースをクリーンアップ

**処理フロー:**
```
1. update_spreadsheet() 実行
2. ISBNリストシートから全レコードを取得
3. 当日未更新のISBNをフィルタリング（最大10件）
4. 各ISBNに対して:
   a. search_isbn_estimate() でValueBooks.jpから価格取得
   b. スプレッドシートのE列（最新見積価格）を更新
   c. F列（価格更新日時）を更新
   d. G列（価格増減）を計算して更新
   e. _add_price_history() で価格履歴に記録
5. _write_execution_summary() で実行結果をエラーログに記録
6. close() でブラウザを閉じる
```

**スクレイピングの仕組み:**
```
1. Selenium + ChromeDriverでブラウザを操作
2. ValueBooks.jpの見積ページにアクセス
3. ISBN入力フォームを検索（複数セレクタを試行）
4. ISBNを1文字ずつ入力（人間らしい操作）
5. Enterキーで検索実行
6. 結果ページから価格を抽出（buy-priceクラス）
7. 「該当する商品は見つかりませんでした」の場合は0円を返す
```

**Bot検出回避:**
- User-Agentを設定
- `--disable-blink-features=AutomationControlled`
- `webdriver`プロパティを未定義に設定
- 1文字ずつ入力（time.sleep(0.1)）
- ページクリーンアップ（localStorage、sessionStorage、cookies）

**メモリ対策:**
- 1件ずつ処理
- 各ISBN処理後にページをクリーンアップ
- MAX_PROCESS_COUNT = 10 に制限

---

#### main.py

**役割:** Cloud Functions のエントリーポイント

**主要な関数:**
- `update_prices(request)` - HTTPトリガーで価格更新を実行

**処理フロー:**
```
1. 環境変数 SPREADSHEET_ID を取得
2. ValueBooksScraper を初期化
3. update_spreadsheet() を実行
4. 成功: ('Success: ...', 200) を返す
5. 失敗: ('Error: ...', 500) を返す
6. finally: scraper.close() でリソース解放
```

**Cloud Functions特有の処理:**
- `functions_framework.http` デコレータでHTTPトリガーを設定
- 環境変数から設定を取得（SPREADSHEET_ID）
- エラーハンドリングでステータスコードを返す

---

#### requirements.txt

**役割:** Pythonパッケージの依存関係を定義

**インストールされるパッケージ:**
```
selenium>=4.15.0              # ブラウザ自動操作
gspread>=5.12.0              # Google Sheets API
oauth2client>=4.1.3          # Google認証
functions-framework>=3.0.0   # Cloud Functions
pytz                         # タイムゾーン処理
google-cloud-logging>=3.5.0  # Cloud Logging
```

**修正時の注意:**
- バージョンを変更する場合は互換性を確認
- 新しいパッケージを追加した場合は再デプロイが必要

---

### 列番号対応表（重要）

#### ISBNリストシート

| 列番号 | 列名 | Config.gs | Python | 説明 |
|---|---|---|---|---|
| 1 | A | `ISBN: 1` | `1` | ISBN |
| 2 | B | `TITLE: 2` | `2` | 書籍名 |
| 3 | C | `AUTHOR: 3` | - | 著者 |
| 4 | D | `PUBLISHER: 4` | - | 出版社 |
| 5 | E | `PRICE: 5` | `5` | 最新見積価格 |
| 6 | F | `UPDATED: 6` | `6` | 価格更新日時 |
| 7 | G | `PRICE_CHANGE: 7` | `7` | 価格増減 |
| 8 | H | `CHECKBOX: 8` | - | チェックボックス |

**重要:** 
- 列を追加・削除した場合は両方を修正すること
- Pythonの `sheet.update_cell(i, 列番号, 値)` を確認

---

#### 買取完了シート

| 列番号 | 列名 | Config.gs | 説明 |
|---|---|---|---|
| 1 | A | `ISBN: 1` | ISBN |
| 2 | B | `TITLE: 2` | タイトル |
| 3 | C | `AUTHOR: 3` | 著者 |
| 4 | D | `PUBLISHER: 4` | 出版社 |
| 5 | E | `ESTIMATE: 5` | 最新見積価格 |
| 6 | F | `ACTUAL: 6` | 売却価格（手動入力） |
| 7 | G | `DIFFERENCE: 7` | 利益（=F-E） |
| 8 | H | `DATE: 8` | 登録日 |

---

### 緊急時の対応

#### GCPの処理が止まった場合

```bash
# 1. ログを確認
gcloud functions logs read update_prices \
  --region asia-northeast1 \
  --gen2 \
  --limit 100

# 2. 手動実行して動作確認
gcloud functions call update_prices \
  --region asia-northeast1 \
  --gen2

# 3. エラーが続く場合は一時的に無効化
gcloud scheduler jobs pause daily-price-update \
  --location asia-northeast1

# 4. 修正後に再開
gcloud scheduler jobs resume daily-price-update \
  --location asia-northeast1
```

---

#### GASのトリガーが動作しない場合

```
1. Apps Scriptエディタを開く
2. 左側メニュー「トリガー」
3. 該当するトリガーを削除
4. 新しくトリガーを追加
5. 権限を再承認
```

---

#### データが壊れた場合

```
1. スプレッドシートの「バージョン履歴」から復元
   - ファイル → バージョン履歴 → バージョン履歴を表示
   - 正常だった時点を選択
   - 「このバージョンを復元」

2. ダッシュボードのみ壊れた場合
   - メニュー「古本買取システム」→「ダッシュボードをセットアップ」
```

---

**以上**
