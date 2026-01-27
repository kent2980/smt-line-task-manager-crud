# EmailSender.psm1
# メール送信を行うモジュール（Microsoft Graph PowerShell対応）

function Send-ErrorEmail {
    <#
    .SYNOPSIS
    エラーメールを送信します（Microsoft Graph PowerShell使用）。
    
    .DESCRIPTION
    Microsoft Graph PowerShellを使用してメールを送信します。
    Microsoft.Graph.Usersモジュールが必要です。
    
    .PARAMETER Subject
    メールの件名
    
    .PARAMETER Body
    メールの本文
    
    .PARAMETER To
    送信先メールアドレス
    
    .PARAMETER From
    送信元メールアドレス（ユーザーIDまたはメールアドレス）
    
    .PARAMETER TenantId
    Azure ADテナントID（オプション）
    
    .PARAMETER ClientId
    Azure ADアプリケーション（クライアント）ID（オプション）
    
    .PARAMETER ClientSecret
    Azure ADアプリケーションのクライアントシークレット（オプション）
    
    .EXAMPLE
    Send-ErrorEmail -Subject "エラー通知" -Body "エラーが発生しました" -To "user@example.com" -From "sender@example.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Mandatory = $true)]
        [string]$To,
        
        [Parameter(Mandatory = $true)]
        [string]$From,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )
    
    try {
        Write-Verbose "メール送信中: $To"
        
        # アプリケーション認証のパラメータチェック
        if (-not ($TenantId -and $ClientId -and $ClientSecret)) {
            throw "アプリケーション認証にはTenantId、ClientId、ClientSecretが必要です"
        }
        
        # OAuth2トークンエンドポイントからアクセストークンを取得
        Write-Verbose "アクセストークンを取得中..."
        $tokenBody = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $ClientId
            client_secret = $ClientSecret
        }
        
        $tokenResponse = Invoke-RestMethod `
            -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $tokenBody `
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop
        
        $accessToken = $tokenResponse.access_token
        
        if (-not $accessToken) {
            throw "アクセストークンの取得に失敗しました"
        }
        
        Write-Verbose "アクセストークンを取得しました"
        
        # メールメッセージの作成
        $mailBody = @{
            message = @{
                subject = $Subject
                body    = @{
                    contentType = "Text"
                    content     = $Body
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $To
                        }
                    }
                )
            }
        } | ConvertTo-Json -Depth 5
        
        # Microsoft Graph APIを使用してメール送信
        Write-Verbose "Microsoft Graph APIを使用してメール送信: $From -> $To"
        
        Invoke-RestMethod `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/users/$From/sendMail" `
            -Headers @{
                Authorization = "Bearer $accessToken"
            } `
            -Body $mailBody `
            -ContentType "application/json" `
            -ErrorAction Stop
        
        Write-Verbose "メール送信成功"
    }
    catch {
        $errorMessage = "メール送信エラー: $_"
        Write-Error $errorMessage
        
        # エラーの詳細を表示
        if ($_.Exception.InnerException) {
            Write-Error "詳細: $($_.Exception.InnerException.Message)"
        }
        
        throw $errorMessage
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Send-ErrorEmail

