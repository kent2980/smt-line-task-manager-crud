# Utils.psm1
# ユーティリティ関数モジュール

function ConvertTo-DateString {
    <#
    .SYNOPSIS
    日付値を文字列形式に変換します。
    
    .DESCRIPTION
    Excelから取得した日付値（DateTime、数値、文字列）を、指定された形式の日付文字列に変換します。
    
    .PARAMETER DateValue
    変換する日付値（DateTime、数値、文字列）
    
    .PARAMETER Format
    日付文字列の形式（デフォルト: "yyyy-MM-dd"）
    
    .EXAMPLE
    $dateString = ConvertTo-DateString -DateValue $dateValue
    
    .EXAMPLE
    $dateString = ConvertTo-DateString -DateValue $dateValue -Format "yyyy/MM/dd"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $DateValue,
        
        [Parameter(Mandatory = $false)]
        [string]$Format = "yyyy-MM-dd"
    )
    
    if ($null -eq $DateValue) {
        return $null
    }
    
    try {
        $dateTime = $null
        
        # 既にDateTime型の場合はそのまま使用
        if ($DateValue -is [DateTime]) {
            $dateTime = $DateValue
        }
        # 数値（Excelシリアル値）の場合は変換
        elseif ($DateValue -is [double] -or $DateValue -is [int] -or $DateValue -is [long]) {
            $dateTime = [DateTime]::FromOADate($DateValue)
        }
        # 文字列の場合はパースを試みる
        elseif ($DateValue -is [string] -and -not [string]::IsNullOrWhiteSpace($DateValue)) {
            $dateTime = [DateTime]::Parse($DateValue)
        }
        
        # DateTime型に変換できた場合は文字列形式に変換
        if ($null -ne $dateTime) {
            return $dateTime.ToString($Format)
        }
        else {
            # 変換できない場合は元の値を文字列として使用
            Write-Warning "日付として認識できませんでした。元の値を文字列として返します（値: $DateValue）"
            return $DateValue.ToString()
        }
    }
    catch {
        # 変換に失敗した場合は元の値を文字列として使用
        Write-Warning "日付の変換に失敗しました（値: $DateValue）: $_"
        return if ($null -ne $DateValue) { $DateValue.ToString() } else { $null }
    }
}

function Load-EnvFile {
    <#
    .SYNOPSIS
    .envファイルから環境変数を読み込みます。
    
    .DESCRIPTION
    .envファイルの各行を読み込み、KEY=VALUE形式の行を環境変数として設定します。
    コメント行（#で始まる行）と空行は無視されます。
    
    .PARAMETER EnvFilePath
    .envファイルのパス（デフォルト: スクリプトディレクトリの.env）
    
    .EXAMPLE
    Load-EnvFile -EnvFilePath "C:\path\to\.env"
    
    .EXAMPLE
    Load-EnvFile  # デフォルトパスを使用
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$EnvFilePath
    )
    
    # デフォルトパスの設定
    if ([string]::IsNullOrEmpty($EnvFilePath)) {
        $scriptDirectory = Split-Path -Parent $PSScriptRoot
        $EnvFilePath = Join-Path $scriptDirectory ".env"
    }
    
    # .envファイルが存在しない場合は警告を出して終了
    if (-not (Test-Path $EnvFilePath)) {
        Write-Warning ".envファイルが見つかりません: $EnvFilePath"
        return @{}
    }
    
    $envVars = @{}
    
    try {
        Write-Verbose ".envファイルを読み込み中: $EnvFilePath"
        
        # .envファイルの各行を読み込み
        $lines = Get-Content -Path $EnvFilePath -Encoding UTF8
        
        foreach ($line in $lines) {
            # 空行とコメント行をスキップ
            $trimmedLine = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) {
                continue
            }
            
            # KEY=VALUE形式を解析
            if ($trimmedLine -match '^([^=]+)=(.*)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                
                # 値の引用符を削除（"value" または 'value'）
                if (($value.StartsWith('"') -and $value.EndsWith('"')) -or 
                    ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                    $value = $value.Substring(1, $value.Length - 2)
                }
                
                $envVars[$key] = $value
                Write-Verbose "環境変数を読み込み: $key"
            }
        }
        
        Write-Verbose ".envファイルの読み込み完了: $($envVars.Count) 件"
        return $envVars
    }
    catch {
        Write-Error ".envファイルの読み込みエラー: $_"
        return @{}
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function ConvertTo-DateString, Load-EnvFile

