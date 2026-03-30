# Test-Email.ps1
# メール送信テストスクリプト

# UTF-8エンコーディングを設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# スクリプトのディレクトリを取得
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulesPath = Join-Path $scriptDirectory "modules"

# 設定ファイルを読み込み
$configPath = Join-Path $scriptDirectory "Config.ps1"
if (Test-Path $configPath) {
    $Config = . $configPath
}
else {
    Write-Error "設定ファイルが見つかりません: $configPath"
    exit 1
}

# EmailSenderモジュールを読み込み
$emailSenderPath = Join-Path $modulesPath "EmailSender.psm1"
if (Test-Path $emailSenderPath) {
    Import-Module $emailSenderPath -Force
    Write-Host "EmailSenderモジュールを読み込みました" -ForegroundColor Green
}
else {
    Write-Error "EmailSenderモジュールが見つかりません: $emailSenderPath"
    exit 1
}

# メール設定の確認
Write-Host "`n=== メール設定確認 ===" -ForegroundColor Cyan
Write-Host "送信元: $($Config.Email.From)"
Write-Host "送信先: $($Config.Email.To)"
if ($Config.Email.TenantId) {
    Write-Host "テナントID: $($Config.Email.TenantId)"
    Write-Host "クライアントID: $($Config.Email.ClientId)"
    Write-Host "クライアントシークレット: $($Config.Email.ClientSecret -replace '.', '*')"  # シークレットをマスク
    Write-Host "認証方法: アプリケーション認証（OAuth2クライアント認証情報フロー）"
}
else {
    Write-Host "認証方法: アプリケーション認証（必須）"
}
Write-Host ""

# 設定の検証
$errors = @()
if ([string]::IsNullOrWhiteSpace($Config.Email.From)) {
    $errors += "送信元メールアドレスが設定されていません"
}
if ([string]::IsNullOrWhiteSpace($Config.Email.To)) {
    $errors += "送信先メールアドレスが設定されていません"
}
if ([string]::IsNullOrWhiteSpace($Config.Email.TenantId)) {
    $errors += "テナントIDが設定されていません（アプリケーション認証に必須）"
}
if ([string]::IsNullOrWhiteSpace($Config.Email.ClientId)) {
    $errors += "クライアントIDが設定されていません（アプリケーション認証に必須）"
}
if ([string]::IsNullOrWhiteSpace($Config.Email.ClientSecret)) {
    $errors += "クライアントシークレットが設定されていません（アプリケーション認証に必須）"
}

if ($errors.Count -gt 0) {
    Write-Host "エラー: 以下の設定が不足しています:" -ForegroundColor Red
    foreach ($error in $errors) {
        Write-Host "  - $error" -ForegroundColor Red
    }
    Write-Host "`n.envファイルに必要な設定を追加してください。" -ForegroundColor Yellow
    Write-Host "必須設定:" -ForegroundColor Yellow
    Write-Host "  EMAIL_FROM=your-email@yourdomain.com" -ForegroundColor Yellow
    Write-Host "  EMAIL_TO=recipient@yourdomain.com" -ForegroundColor Yellow
    Write-Host "  EMAIL_TENANT_ID=your-tenant-id" -ForegroundColor Yellow
    Write-Host "  EMAIL_CLIENT_ID=your-client-id" -ForegroundColor Yellow
    Write-Host "  EMAIL_CLIENT_SECRET=your-client-secret" -ForegroundColor Yellow
    exit 1
}

# メール送信テスト
Write-Host "=== メール送信テスト開始 ===" -ForegroundColor Cyan
Write-Host ""

try {
    $testSubject = "【テスト】メール送信テスト - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $testBody = @"
これはメール送信のテストメールです。

送信日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
スクリプト: Test-Email.ps1
認証方法: OAuth2クライアント認証情報フロー

このメールが正常に受信できていれば、メール送信機能は正常に動作しています。
"@

    Write-Host "メール送信中..." -ForegroundColor Yellow
    
    # Microsoft Graph APIを使用してメール送信（OAuth2クライアント認証情報フロー）
    $emailParams = @{
        Subject      = $testSubject
        Body         = $testBody
        To           = $Config.Email.To
        From         = $Config.Email.From
        TenantId     = $Config.Email.TenantId
        ClientId     = $Config.Email.ClientId
        ClientSecret = $Config.Email.ClientSecret
    }
    
    Send-ErrorEmail @emailParams -Verbose
    
    Write-Host "`n✓ メール送信が正常に完了しました！" -ForegroundColor Green
    Write-Host "送信先: $($Config.Email.To)" -ForegroundColor Green
    Write-Host "件名: $testSubject" -ForegroundColor Green
    Write-Host "`n受信トレイを確認してください。" -ForegroundColor Cyan
}
catch {
    Write-Host "`n✗ メール送信に失敗しました:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # エラーの詳細を表示
    if ($_.Exception.InnerException) {
        Write-Host "`n詳細:" -ForegroundColor Yellow
        Write-Host $_.Exception.InnerException.Message -ForegroundColor Yellow
    }
    
    Write-Host "`nトラブルシューティング:" -ForegroundColor Yellow
    Write-Host "1. .envファイルの設定を確認してください" -ForegroundColor Yellow
    Write-Host "2. テナントID、クライアントID、クライアントシークレットが正しいか確認してください" -ForegroundColor Yellow
    Write-Host "3. Azure ADアプリケーションにMail.Send権限が付与されているか確認してください" -ForegroundColor Yellow
    Write-Host "4. クライアントシークレットの有効期限が切れていないか確認してください" -ForegroundColor Yellow
    Write-Host "5. ネットワーク接続を確認してください" -ForegroundColor Yellow
    
    exit 1
}

