# Config.ps1
# 設定ファイル

# .envファイルから機密情報を読み込み
$envVars = @{}
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$envFilePath = Join-Path $scriptDirectory ".env"

if (Test-Path $envFilePath) {
    # UtilsモジュールのLoad-EnvFile関数を使用
    $utilsPath = Join-Path $scriptDirectory "modules\Utils.psm1"
    if (Test-Path $utilsPath) {
        Import-Module $utilsPath -Force
        $envVars = Load-EnvFile -EnvFilePath $envFilePath
    }
    else {
        # Utilsモジュールが読み込めない場合は手動で読み込み
        $lines = Get-Content -Path $envFilePath -Encoding UTF8 -ErrorAction SilentlyContinue
        foreach ($line in $lines) {
            $trimmedLine = $line.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmedLine) -and -not $trimmedLine.StartsWith('#')) {
                if ($trimmedLine -match '^([^=]+)=(.*)$') {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    if (($value.StartsWith('"') -and $value.EndsWith('"')) -or 
                        ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                        $value = $value.Substring(1, $value.Length - 2)
                    }
                    $envVars[$key] = $value
                }
            }
        }
    }
}

# ファイルパス設定
$Config = @{

    # キントーンアプリID
    # .envファイルから読み込む（必須）
    AppId               = if ($envVars['KINTONE_APP_ID']) { [int]$envVars['KINTONE_APP_ID'] } else { throw "KINTONE_APP_IDが.envファイルに設定されていません" }

    # データディレクトリ
    # .envファイルから読み込む（必須）
    DataDirectory       = if ($envVars['DATA_DIRECTORY']) { $envVars['DATA_DIRECTORY'] } else { throw "DATA_DIRECTORYが.envファイルに設定されていません" }

    # 一時ディレクトリ
    # .envファイルから読み込む（必須）
    DataDirectoryTemp   = if ($envVars['DATA_DIRECTORY_TEMP']) { $envVars['DATA_DIRECTORY_TEMP'] } else { throw "DATA_DIRECTORY_TEMPが.envファイルに設定されていません" }

    # 変換後のディレクトリ
    # .envファイルから読み込む（必須）
    DataDirectoryXlsx   = if ($envVars['DATA_DIRECTORY_XLSX']) { $envVars['DATA_DIRECTORY_XLSX'] } else { throw "DATA_DIRECTORY_XLSXが.envファイルに設定されていません" }
    
    # ログディレクトリ
    # .envファイルから読み込む（必須）
    LogDirectory        = if ($envVars['LOG_DIRECTORY']) { $envVars['LOG_DIRECTORY'] } else { throw "LOG_DIRECTORYが.envファイルに設定されていません" }

    # 各エクセルファイルのタイムスタンプ永続化ファイル
    # .envファイルから読み込む（必須）
    TimestampFile       = if ($envVars['TIMESTAMP_FILE']) { $envVars['TIMESTAMP_FILE'] } else { throw "TIMESTAMP_FILEが.envファイルに設定されていません" }

    # ファイル名パターン
    FileNamePattern     = "GC0{0}.xls"  # {0} が数字に置き換えられます

    # 変換後のファイル名パターン
    FileNamePatternXlsx = "GC0{0}.xlsx"  # {0} が数字に置き換えられます
    
    # 処理するファイル番号の範囲
    FileNumberStart     = 1
    FileNumberEnd       = 9

    # API設定（キントーン）
    # .envファイルから読み込む（必須）
    Api                 = @{
        Uri        = if ($envVars['KINTONE_API_URI']) { $envVars['KINTONE_API_URI'] } else { throw "KINTONE_API_URIが.envファイルに設定されていません" }
        Method     = @{
            Get    = "GET"
            Post   = "POST"
            Put    = "PUT"
            Delete = "DELETE"
        }
        Headers    = @{
            "X-Cybozu-API-Token" = if ($envVars['KINTONE_API_TOKEN']) { $envVars['KINTONE_API_TOKEN'] } else { throw "KINTONE_API_TOKENが.envファイルに設定されていません" }
            "Content-Type"       = "application/json"
        }
        TimeoutSec = 30
    }
    
    # メール設定（Microsoft Graph PowerShell）
    # .envファイルから読み込む（必須）
    Email               = @{
        From         = if ($envVars['EMAIL_FROM']) { $envVars['EMAIL_FROM'] } else { throw "EMAIL_FROMが.envファイルに設定されていません" }
        To           = if ($envVars['EMAIL_TO']) { $envVars['EMAIL_TO'] } else { throw "EMAIL_TOが.envファイルに設定されていません" }
        # アプリケーション認証を使用する場合（オプション）
        TenantId     = if ($envVars['EMAIL_TENANT_ID']) { $envVars['EMAIL_TENANT_ID'] } else { "" }
        ClientId     = if ($envVars['EMAIL_CLIENT_ID']) { $envVars['EMAIL_CLIENT_ID'] } else { "" }
        ClientSecret = if ($envVars['EMAIL_CLIENT_SECRET']) { $envVars['EMAIL_CLIENT_SECRET'] } else { "" }
    }
}

# 設定をエクスポート
$Config

