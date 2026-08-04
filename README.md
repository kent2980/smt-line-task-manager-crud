# Excel処理スクリプト

.xlsファイルを.xlsxに変換し、Excelデータを読み取ってJSONに変換し、APIに送信するPowerShellスクリプトです。

## ディレクトリ構造

```
.
├── Main.ps1                 # メイン処理スクリプト
├── Config.ps1               # 設定ファイル
├── modules/                 # モジュールディレクトリ
│   ├── ExcelConverter.psm1  # .xls→.xlsx変換モジュール
│   ├── ExcelReader.psm1     # Excel読み取りモジュール（テンプレート）
│   ├── JsonConverter.psm1  # JSON変換モジュール
│   ├── ApiClient.psm1      # API送信モジュール
│   ├── Logger.psm1         # ログ出力モジュール
│   └── EmailSender.psm1    # メール送信モジュール
└── README.md               # このファイル
```

## 処理フロー

1. **COMオブジェクトで.xls→.xlsxに変換**

   - `ExcelConverter.psm1`の`Convert-XlsToXlsx`関数を使用
2. **ImportExcelモジュールで読み取り**

   - `ExcelReader.psm1`の`Read-ExcelData`関数を使用
   - **注意**: この関数はテンプレートです。実際の読み取り処理を実装してください
3. **読み取ったデータをJSON変換**

   - `JsonConverter.psm1`の`ConvertTo-JsonData`関数を使用
4. **API送信**

   - `ApiClient.psm1`の`Send-ApiRequest`関数を使用
5. **失敗時の処理**

   - ログファイルに出力（`Logger.psm1`）
   - メール送信（`EmailSender.psm1`）

## セットアップ

### 1. 必要なモジュールのインストール

```powershell
# ImportExcelモジュールをインストール
Install-Module -Name ImportExcel -Scope CurrentUser

# Microsoft Graph PowerShellモジュールをインストール（メール送信に必要）
Install-Module -Name Microsoft.Graph -Scope CurrentUser
```

### 2. 設定ファイルの編集

#### Config.ps1の編集

`Config.ps1`を編集して、以下の設定を行ってください：

- **データディレクトリ**: Excelファイルが格納されているディレクトリ
- **ログディレクトリ**: ログファイルを保存するディレクトリ
- **ファイル名パターン**: 処理するファイル名のパターン
- **API設定**: APIのエンドポイントURL、メソッド、ヘッダーなど

#### .envファイルの作成（機密情報用）

機密情報（メール認証情報など）は`.env`ファイルから読み込みます。
プロジェクトルートに`.env`ファイルを作成し、以下の形式で設定してください：

```env
# ディレクトリ設定
DATA_DIRECTORY=C:\Users\your-username\Documents\生産計画\data\
DATA_DIRECTORY_XLSX=C:\Users\your-username\Documents\生産計画\data\xlsx\
LOG_DIRECTORY=C:\Users\your-username\Documents\生産計画\data\log\
TIMESTAMP_FILE=C:\Users\your-username\Documents\生産計画\data\timestamp.txt

# Microsoft 365 メール設定（Microsoft Graph PowerShell使用）
EMAIL_FROM=your-email@yourdomain.com
EMAIL_TO=recipient@yourdomain.com

# アプリケーション認証を使用する場合（オプション）
# 対話的認証を使用する場合は、これらの設定は不要です
EMAIL_TENANT_ID=your-tenant-id
EMAIL_CLIENT_ID=your-client-id
EMAIL_CLIENT_SECRET=your-client-secret

# キントーンAPI設定
KINTONE_API_URI=https://your-subdomain.cybozu.com/k/v1/records.json
KINTONE_API_TOKEN=your-api-token
KINTONE_APP_ID=86
```

**重要**: `.env`ファイルは`.gitignore`に含まれているため、Gitにコミットされません。
機密情報を安全に管理できます。

**Microsoft Graph PowerShellの認証について**:

メール送信には以下の2つの認証方法があります：

1. **対話的認証（推奨）**: 初回実行時にブラウザが開き、Microsoft 365アカウントでログインします。`.env`ファイルに`EMAIL_TENANT_ID`、`EMAIL_CLIENT_ID`、`EMAIL_CLIENT_SECRET`の設定は不要です。
2. **アプリケーション認証**: Azure ADアプリケーションを登録し、クライアントIDとシークレットを使用します。`.env`ファイルに`EMAIL_TENANT_ID`、`EMAIL_CLIENT_ID`、`EMAIL_CLIENT_SECRET`を設定してください。

**注意**: Microsoft Graph PowerShellを使用する場合、アプリパスワードは不要です。SMTPを使用する場合のみアプリパスワードが必要です（現在はMicrosoft Graph PowerShellを使用しています）。

### 3. ExcelReaderモジュールの実装

`modules/ExcelReader.psm1`の`Read-ExcelData`関数内のTODOコメント部分を実装してください。

例：

```powershell
if ($WorksheetName) {
    $data = Import-Excel -Path $XlsxPath -WorksheetName $WorksheetName
} else {
    $data = Import-Excel -Path $XlsxPath -WorksheetName 1
}
```

## 使用方法

```powershell
# メインスクリプトを実行
.\Main.ps1

# 詳細ログを表示
.\Main.ps1 -Verbose
```

### 定期実行の設定

Windowsで定期実行する方法については、[Windowsで定期実行する方法](docs/WINDOWS_SCHEDULED_TASK.md)を参照してください。

主な方法：

- **タスクスケジューラ**（推奨）: Windows標準のタスクスケジューラを使用
- **PowerShellスケジュールジョブ**: PowerShellの`Register-ScheduledJob`を使用
- **バッチファイル + タスクスケジューラ**: バッチファイル経由で実行

## 各モジュールの説明

### ExcelConverter.psm1

- **関数**: `Convert-XlsToXlsx`
- **機能**: .xlsファイルを.xlsx形式に変換
- **パラメータ**:
  - `XlsPath`: 変換する.xlsファイルのパス
  - `XlsxPath`: 変換後の.xlsxファイルのパス（省略可）

### ExcelReader.psm1

- **関数**: `Read-ExcelData`
- **機能**: Excelファイルを読み取り（テンプレート）
- **パラメータ**:
  - `XlsxPath`: 読み取る.xlsxファイルのパス
  - `WorksheetName`: ワークシート名（省略可）

### JsonConverter.psm1

- **関数**: `ConvertTo-JsonData`
- **機能**: データをJSON形式に変換
- **パラメータ**:
  - `InputObject`: 変換するオブジェクト
  - `Depth`: JSONの深さ（デフォルト: 10）
  - `Compress`: 圧縮形式で出力するか

### ApiClient.psm1

- **関数**: `Send-ApiRequest`
- **機能**: APIにリクエストを送信
- **パラメータ**:
  - `Uri`: APIのエンドポイントURL
  - `Method`: HTTPメソッド（デフォルト: POST）
  - `Body`: 送信するJSONデータ
  - `Headers`: 追加のHTTPヘッダー
  - `ContentType`: コンテンツタイプ（デフォルト: application/json）
  - `TimeoutSec`: タイムアウト秒数（デフォルト: 30）

### Logger.psm1

- **関数**: `Write-Log`, `Initialize-Logger`
- **機能**: ログファイルへの書き込み
- **パラメータ**:
  - `Message`: ログメッセージ
  - `LogPath`: ログファイルのパス
  - `LogLevel`: ログレベル（INFO, WARNING, ERROR）

### EmailSender.psm1

- **関数**: `Send-ErrorEmail`
- **機能**: エラーメールを送信
- **パラメータ**:
  - `Subject`: メールの件名
  - `Body`: メールの本文
  - `To`: 送信先メールアドレス
  - `From`: 送信元メールアドレス
  - `SmtpServer`: SMTPサーバー
  - `SmtpPort`: SMTPポート（デフォルト: 587）
  - `Username`: SMTP認証ユーザー名
  - `Password`: SMTP認証パスワード
  - `UseSsl`: SSL/TLSを使用するか（デフォルト: $true）

## エラーハンドリング

- 各処理でエラーが発生した場合、ログファイルに記録されます
- エラーが発生した場合、自動的にメールが送信されます
- 致命的なエラーが発生した場合も、メールが送信されます

## 注意事項

- Excelがインストールされている必要があります（COMオブジェクトを使用するため）
- ImportExcelモジュールがインストールされている必要があります
- メール送信にはSMTPサーバーの設定が必要です
- パスワードは平文で保存しないことを推奨します（環境変数や暗号化されたファイルから読み込むことを推奨）

## トラブルシューティング

### モジュールが見つからない

- `modules`ディレクトリが正しい場所にあるか確認してください
- モジュールファイルのパスが正しいか確認してください

### Excel変換エラー

- Excelがインストールされているか確認してください
- ファイルが他のプロセスで開かれていないか確認してください

### API送信エラー

- APIのエンドポイントURLが正しいか確認してください
- ネットワーク接続を確認してください
- APIの認証情報が正しいか確認してください

### メール送信エラー

- SMTPサーバーの設定が正しいか確認してください
- 認証情報が正しいか確認してください
- ファイアウォールの設定を確認してください
