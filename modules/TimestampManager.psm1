# TimestampManager.psm1
# タイムスタンプ管理モジュール

function Get-FileTimestamp {
    <#
    .SYNOPSIS
    ファイルの最終更新日時を取得します。
    
    .DESCRIPTION
    指定されたファイルの最終更新日時を取得します。
    
    .PARAMETER FilePath
    ファイルのパス
    
    .EXAMPLE
    $timestamp = Get-FileTimestamp -FilePath "C:\data\file.xls"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "ファイルが存在しません: $FilePath"
    }
    
    $fileInfo = Get-Item $FilePath
    return $fileInfo.LastWriteTime
}

function Load-Timestamps {
    <#
    .SYNOPSIS
    タイムスタンプファイルからタイムスタンプ情報を読み込みます。
    
    .DESCRIPTION
    永続化されたタイムスタンプファイルから、ファイルパスとタイムスタンプのマッピングを読み込みます。
    
    .PARAMETER TimestampFilePath
    タイムスタンプファイルのパス
    
    .EXAMPLE
    $timestamps = Load-Timestamps -TimestampFilePath "C:\data\timestamp.txt"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TimestampFilePath
    )
    
    $timestamps = @{}
    
    # タイムスタンプファイルが存在する場合は読み込む
    if (Test-Path $TimestampFilePath) {
        try {
            $lines = Get-Content $TimestampFilePath -Encoding UTF8 -ErrorAction Stop
            
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) {
                    continue
                }
                
                # 形式: ファイルパス|タイムスタンプ
                $parts = $line -split '\|', 2
                if ($parts.Length -eq 2) {
                    $filePath = $parts[0].Trim()
                    $timestampStr = $parts[1].Trim()
                    
                    # パスを正規化（大文字小文字を統一、パス区切り文字を統一）
                    $filePath = [System.IO.Path]::GetFullPath($filePath)
                    
                    # タイムスタンプをDateTimeに変換
                    try {
                        # 形式: yyyy-MM-dd HH:mm:ss でパースを試みる
                        $timestamp = [DateTime]::ParseExact($timestampStr, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None)
                        $timestamps[$filePath] = $timestamp
                        Write-Verbose "タイムスタンプを読み込み: $filePath -> $timestamp"
                    }
                    catch {
                        # ParseExactが失敗した場合、TryParseを試みる
                        $timestamp = $null
                        if ([DateTime]::TryParse($timestampStr, [ref]$timestamp)) {
                            $timestamps[$filePath] = $timestamp
                            Write-Verbose "タイムスタンプを読み込み（TryParse）: $filePath -> $timestamp"
                        }
                        else {
                            Write-Warning "タイムスタンプの解析に失敗: $filePath -> $timestampStr"
                        }
                    }
                }
            }
            
            Write-Verbose "タイムスタンプを読み込みました: $($timestamps.Count) 件"
        }
        catch {
            Write-Warning "タイムスタンプファイルの読み込みエラー: $_"
        }
    }
    else {
        Write-Verbose "タイムスタンプファイルが存在しません。新規作成します。"
    }
    
    return $timestamps
}

function Save-Timestamps {
    <#
    .SYNOPSIS
    タイムスタンプ情報をファイルに保存します。
    
    .DESCRIPTION
    ファイルパスとタイムスタンプのマッピングをタイムスタンプファイルに保存します。
    
    .PARAMETER Timestamps
    タイムスタンプのハッシュテーブル（キー: ファイルパス, 値: DateTime）
    
    .PARAMETER TimestampFilePath
    タイムスタンプファイルのパス
    
    .EXAMPLE
    Save-Timestamps -Timestamps $timestamps -TimestampFilePath "C:\data\timestamp.txt"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Timestamps,
        
        [Parameter(Mandatory = $true)]
        [string]$TimestampFilePath
    )
    
    try {
        # ディレクトリが存在しない場合は作成
        $directory = Split-Path -Parent $TimestampFilePath
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        # タイムスタンプをファイルに書き込み
        $lines = @()
        foreach ($filePath in $Timestamps.Keys | Sort-Object) {
            $timestamp = $Timestamps[$filePath]
            $timestampStr = $timestamp.ToString("yyyy-MM-dd HH:mm:ss")
            $lines += "$filePath|$timestampStr"
        }
        
        $lines | Out-File -FilePath $TimestampFilePath -Encoding UTF8 -Force
        
        Write-Verbose "タイムスタンプを保存しました: $($Timestamps.Count) 件"
    }
    catch {
        throw "タイムスタンプ保存エラー: $_"
    }
}

function Test-FileUpdated {
    <#
    .SYNOPSIS
    ファイルが更新されているかチェックします。
    
    .DESCRIPTION
    ファイルの現在のタイムスタンプと、保存されているタイムスタンプを比較して、
    ファイルが更新されているかどうかを判定します。
    
    .PARAMETER FilePath
    チェックするファイルのパス
    
    .PARAMETER SavedTimestamp
    保存されているタイムスタンプ（DateTime、null可）
    
    .PARAMETER Timestamps
    タイムスタンプのハッシュテーブル（パス検索用）
    
    .EXAMPLE
    $isUpdated = Test-FileUpdated -FilePath "C:\data\file.xls" -SavedTimestamp $savedTime
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        $SavedTimestamp,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Timestamps
    )
    
    if (-not (Test-Path $FilePath)) {
        return $false
    }
    
    # パスを正規化
    $normalizedPath = [System.IO.Path]::GetFullPath($FilePath)
    
    # SavedTimestampがnullの場合、Timestampsハッシュテーブルから検索を試みる
    if ($null -eq $SavedTimestamp -or -not ($SavedTimestamp -is [DateTime])) {
        if ($Timestamps -and $Timestamps.ContainsKey($normalizedPath)) {
            $SavedTimestamp = $Timestamps[$normalizedPath]
            Write-Verbose "タイムスタンプをハッシュテーブルから取得: $normalizedPath -> $SavedTimestamp"
        }
        else {
            Write-Verbose "タイムスタンプが保存されていません: $normalizedPath"
            return $true
        }
    }
    
    $currentTimestamp = Get-FileTimestamp -FilePath $FilePath
    
    # タイムスタンプを秒単位で切り捨てて比較（ミリ秒の差を無視）
    $currentTimestampRounded = $currentTimestamp.Date.AddSeconds([Math]::Floor($currentTimestamp.TimeOfDay.TotalSeconds))
    $savedTimestampRounded = $SavedTimestamp.Date.AddSeconds([Math]::Floor($SavedTimestamp.TimeOfDay.TotalSeconds))
    
    # 現在のタイムスタンプが保存されているタイムスタンプより新しい場合は更新あり
    $isUpdated = $currentTimestampRounded -gt $savedTimestampRounded
    
    if ($isUpdated) {
        Write-Verbose "ファイルが更新されています: $normalizedPath (保存: $savedTimestampRounded, 現在: $currentTimestampRounded)"
    }
    else {
        Write-Verbose "ファイルは更新されていません: $normalizedPath (保存: $savedTimestampRounded, 現在: $currentTimestampRounded)"
    }
    
    return $isUpdated
}

function Update-FileTimestamp {
    <#
    .SYNOPSIS
    タイムスタンプ情報を更新します。
    
    .DESCRIPTION
    指定されたファイルのタイムスタンプを取得し、タイムスタンプハッシュテーブルに追加または更新します。
    
    .PARAMETER FilePath
    ファイルのパス
    
    .PARAMETER Timestamps
    タイムスタンプのハッシュテーブル（参照渡し）
    
    .EXAMPLE
    Update-FileTimestamp -FilePath "C:\data\file.xls" -Timestamps ([ref]$timestamps)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [ref]$Timestamps
    )
    
    try {
        $currentTimestamp = Get-FileTimestamp -FilePath $FilePath
        # パスを正規化して保存
        $normalizedPath = [System.IO.Path]::GetFullPath($FilePath)
        $Timestamps.Value[$normalizedPath] = $currentTimestamp
        Write-Verbose "タイムスタンプを更新しました: $normalizedPath -> $currentTimestamp"
    }
    catch {
        throw "タイムスタンプ更新エラー: $_"
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Get-FileTimestamp, Load-Timestamps, Save-Timestamps, Test-FileUpdated, Update-FileTimestamp

