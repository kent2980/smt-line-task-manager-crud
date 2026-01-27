# Exchange Online PowerShell モジュールのインストール

`Get-CASMailbox`などのExchange Onlineコマンドを使用するには、Exchange Online PowerShellモジュールが必要です。

## インストール方法

### Exchange Online PowerShell V3 モジュール（推奨）

```powershell
# Exchange Online PowerShell V3モジュールをインストール
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# モジュールをインポート
Import-Module ExchangeOnlineManagement

# Exchange Onlineに接続
Connect-ExchangeOnline -UserPrincipalName "your-email@yourdomain.com"
```

### 接続後のコマンド例

```powershell
# メールボックスの設定を確認
Get-CASMailbox -Identity "your-email@yourdomain.com" | Select-Object PrimarySmtpAddress, SmtpClientAuthenticationDisabled

# 基本認証の状態を確認
Get-CASMailbox -Identity "your-email@yourdomain.com" | Format-List SmtpClientAuthenticationDisabled
```

## 注意事項

- Exchange Online PowerShellモジュールは、メールボックスの設定を確認・変更するために使用します
- アプリパスワードの生成自体は、通常Webブラウザから行う方が簡単です
- 管理者権限が必要な場合があります

## 参考リンク

- [Exchange Online PowerShell V3 モジュール](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
