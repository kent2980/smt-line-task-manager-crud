# ExcelConverter.psm1
# .xlsファイルを.xlsx形式に変換するモジュール

function Convert-XlsToXlsx {
    <#
    .SYNOPSIS
    .xlsファイルを.xlsx形式に変換します。
    
    .DESCRIPTION
    COMオブジェクトを使用して.xlsファイルを読み込み、.xlsx形式で同じディレクトリに保存します。
    
    .PARAMETER XlsPath
    変換する.xlsファイルのパス
    
    .PARAMETER XlsxPath
    変換後の.xlsxファイルのパス（省略時は同じディレクトリに保存）

    .PARAMETER TimestampPath
    タイムスタンプ永続化ファイルのパス（省略時はタイムスタンプを更新しない）
    
    .EXAMPLE
    Convert-XlsToXlsx -XlsPath "C:\data\file.xls"
    
    .EXAMPLE
    Convert-XlsToXlsx -XlsPath "C:\data\file.xls" -XlsxPath "C:\output\file.xlsx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlsPath,
        
        [Parameter(Mandatory = $false)]
        [string]$XlsxPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Password = ""
    )
    
    # パス検証
    if (-not (Test-Path $XlsPath)) {
        throw "ファイルが存在しません: $XlsPath"
    }

    
    # XlsxPathが指定されていない場合は、同じディレクトリに.xlsx拡張子で保存
    if ([string]::IsNullOrEmpty($XlsxPath)) {
        $directory = Split-Path -Parent $XlsPath
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($XlsPath)
        $XlsxPath = Join-Path $directory "$fileName.xlsx"
    }
    
    # 既存の.xlsxファイルがある場合は削除
    if (Test-Path $XlsxPath) {
        Remove-Item $XlsxPath -Force
    }
    
    Write-Verbose ".xlsファイルを.xlsx形式に変換中: $XlsPath -> $XlsxPath"
    
    # ファイルが使用可能になるまで待機（最大5秒）
    $maxWaitTime = 5
    $waitInterval = 0.1
    $waitedTime = 0
    while ($waitedTime -lt $maxWaitTime) {
        try {
            $fileStream = [System.IO.File]::Open($XlsPath, 'Open', 'Read', 'None')
            $fileStream.Close()
            $fileStream.Dispose()
            break
        }
        catch {
            Start-Sleep -Seconds $waitInterval
            $waitedTime += $waitInterval
            if ($waitedTime -ge $maxWaitTime) {
                throw "ファイルが使用中のため、タイムアウトしました: $XlsPath"
            }
        }
    }
    
    $excel = $null
    $workbook = $null
    $workbooks = $null
    
    try {
        # Excel COMオブジェクトを作成
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        
        # .xlsファイルを開く（パスワード保護されている場合は指定されたパスワードで開く）
        $workbooks = $excel.Workbooks
        
        # Workbooks.Openメソッドを呼び出す（パラメータを明示的に指定）
        # パスワードが指定されている場合は使用、そうでなければ空文字列
        if ([string]::IsNullOrEmpty($Password)) {
            # パスワードなしで開く
            $workbook = $workbooks.Open($XlsPath, $false, $true)
        }
        else {
            # パスワードを指定して開く
            $workbook = $workbooks.Open($XlsPath, $false, $true, $null, $Password)
        }
        
        # .xlsx形式で保存（xlOpenXMLWorkbook = 51、パスワードなしで保存）
        # SaveAsメソッドでパスワードを空文字列に設定して、パスワード保護なしで保存
        $workbook.SaveAs($XlsxPath, 51)  # FileFormat=51（パスワードパラメータを省略してパスワードなしで保存）
        
        Write-Verbose "変換完了: $XlsxPath"
        return $XlsxPath
    }
    catch {
        throw "Excel変換エラー: $_"
    }
    finally {
        # クリーンアップ（確実にExcelを終了させる）
        try {
            if ($workbook) {
                $workbook.Close($false)  # SaveChanges=false
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                $workbook = $null
            }
        }
        catch {
            Write-Verbose "ワークブックのクローズ中にエラー: $_"
        }
        
        try {
            if ($workbooks) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
                $workbooks = $null
            }
        }
        catch {
            Write-Verbose "ワークブックスコレクションの解放中にエラー: $_"
        }
        
        try {
            if ($excel) {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                $excel = $null
            }
        }
        catch {
            Write-Verbose "Excelアプリケーションの終了中にエラー: $_"
        }
        
        # ガベージコレクションを複数回実行して確実に解放
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Excelプロセスが完全に終了するまで待機
        Start-Sleep -Milliseconds 500
        
        # Excelプロセスが残っている場合は強制終了（最後の手段）
        $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
        if ($excelProcesses) {
            Write-Verbose "残存しているExcelプロセスを終了します"
            $excelProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            # プロセス終了を待機
            Start-Sleep -Milliseconds 300
        }
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Convert-XlsToXlsx

