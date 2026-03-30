# Logger.psm1
# ログ出力を行うモジュール

function Write-Log {
    <#
    .SYNOPSIS
    ログファイルにメッセージを書き込みます。
    
    .DESCRIPTION
    タイムスタンプ付きでログファイルにメッセージを書き込みます。
    
    .PARAMETER Message
    ログメッセージ
    
    .PARAMETER LogPath
    ログファイルのパス
    
    .PARAMETER LogLevel
    ログレベル（INFO, WARNING, ERROR）
    
    .EXAMPLE
    Write-Log -Message "処理開始" -LogPath "C:\logs\log.txt"
    
    .EXAMPLE
    Write-Log -Message "エラー発生" -LogPath "C:\logs\log.txt" -LogLevel "ERROR"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$LogLevel = "INFO"
    )
    
    try {
        # ログディレクトリが存在しない場合は作成
        $logDirectory = Split-Path -Parent $LogPath
        if (-not (Test-Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
        
        # タイムスタンプを取得
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # ログメッセージをフォーマット
        $logMessage = "[$timestamp] [$LogLevel] $Message"
        
        # ログファイルに書き込み
        Add-Content -Path $LogPath -Value $logMessage -Encoding UTF8
        
        Write-Verbose $logMessage
    }
    catch {
        Write-Error "ログ書き込みエラー: $_"
    }
}

function Initialize-Logger {
    <#
    .SYNOPSIS
    ロガーを初期化します。
    
    .DESCRIPTION
    ログディレクトリを作成し、ログファイルのパスを返します。
    
    .PARAMETER LogDirectory
    ログディレクトリのパス
    
    .PARAMETER LogFileName
    ログファイル名（デフォルト: log.txt）
    
    .EXAMPLE
    $logPath = Initialize-Logger -LogDirectory "C:\logs"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogDirectory,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFileName = "log.txt"
    )
    
    try {
        # ログディレクトリが存在しない場合は作成
        if (-not (Test-Path $LogDirectory)) {
            New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
            Write-Verbose "ログディレクトリを作成しました: $LogDirectory"
        }
        
        $logPath = Join-Path $LogDirectory $LogFileName
        return $logPath
    }
    catch {
        throw "ロガー初期化エラー: $_"
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Write-Log, Initialize-Logger

