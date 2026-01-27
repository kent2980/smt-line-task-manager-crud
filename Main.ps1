# Main.ps1
# メイン処理スクリプト

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

# モジュールを読み込み
$modules = @(
    "Utils",
    "TimestampManager",
    "ExcelConverter",
    "ExcelReader",
    "JsonConverter",
    "ApiClient",
    "Logger",
    "EmailSender",
    "DataSyncManager"
)

foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath "$module.psm1"
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force
        Write-Verbose "モジュールを読み込みました: $module"
    }
    else {
        Write-Error "モジュールが見つかりません: $modulePath"
        exit 1
    }
}

# ロガーを初期化
$logPath = Initialize-Logger -LogDirectory $Config.LogDirectory

# タイムスタンプを読み込む
$timestamps = Load-Timestamps -TimestampFilePath $Config.TimestampFile -Verbose

# エラー情報を保持する変数
$errorCount = 0
$errorDetails = @()
$processedCount = 0
$skippedCount = 0
    
# データ同期処理結果を保持する変数
$syncAddedCount = 0
$syncUpdatedCount = 0
$syncDeletedCount = 0

# 処理されたファイル名を保持する配列
$processedFiles = @()

# 全Excelデータを蓄積する配列
$allExcelData = @()

# 一時ディレクトリ
$tempDir = $Config.DataDirectoryTemp

if (-not (Test-Path $tempDir)) {
    # ディレクトリが存在しない場合は作成
    New-Item -ItemType Directory -Path $tempDir
}

# xlsx変換後のディレクトリ
$xlsxDir = $Config.DataDirectoryXlsx


if (-not (Test-Path $xlsxDir)) {
    # ディレクトリが存在しない場合は作成
    New-Item -ItemType Directory -Path $xlsxDir
}

try {
    # ファイル番号のループ処理
    for ($i = $Config.FileNumberStart; $i -le $Config.FileNumberEnd; $i++) {
        
        try {
            # ファイル名を生成
            $fileName = $Config.FileNamePattern -f $i
            $fileNameXlsx = $Config.FileNamePatternXlsx -f $i
            $xlsPath = Join-Path $Config.DataDirectory $fileName
            $xlsPathTemp = Join-Path $tempDir $fileName
            $xlsxPath = Join-Path $xlsxDir $fileNameXlsx
            
            # ファイル存在確認
            if (-not (Test-Path $xlsPath)) {
                $errorMessage = "ファイルが存在しません: $xlsPath"
                Write-Log -Message $errorMessage -LogPath $logPath -LogLevel "WARNING"
                $errorDetails += $errorMessage
                $errorCount++
                continue
            }
            
            # タイムスタンプ検証: ファイルが更新されているかチェック
            # パスを正規化して検索
            $normalizedXlsPath = [System.IO.Path]::GetFullPath($xlsPath)
            $savedTimestamp = $timestamps[$normalizedXlsPath]
            
            Write-Verbose "タイムスタンプ検証: ファイル=$normalizedXlsPath, 保存済みタイムスタンプ=$savedTimestamp"
            
            $isUpdated = Test-FileUpdated -FilePath $xlsPath -SavedTimestamp $savedTimestamp -Timestamps $timestamps -Verbose

            if (-not $isUpdated) {
                $skippedCount++
                continue
            }
            
            
            # ステップ1: DataDirectory → DataDirectoryTemp にコピー
            Write-Host "  [1/4] DataDirectory → DataDirectoryTemp にコピー中..."
            Copy-Item -Path $xlsPath -Destination $xlsPathTemp -Force
            
            # ステップ2: DataDirectoryTemp → DataDirectoryXlsx に変換（.xls → .xlsx）
            Write-Host "  [2/4] .xls → .xlsx 変換中..."
            $xlsxPath = Convert-XlsToXlsx -XlsPath $xlsPathTemp -XlsxPath $xlsxPath -Verbose
            
            # ステップ3: Excelファイルを読み取り
            Write-Host "  [3/4] Excelファイル読み取り中..."
            $excelData = Read-ExcelData -XlsxPath $xlsxPath -Verbose
            
            # 読み込んだデータを全データ配列に追加
            if ($excelData -and $excelData.Count -gt 0) {
                $allExcelData += $excelData
                Write-Verbose "データを追加しました（現在の総件数: $($allExcelData.Count)）"
            }
            
            # 処理成功後、タイムスタンプを更新
            Update-FileTimestamp -FilePath $xlsPath -Timestamps ([ref]$timestamps) -Verbose
            $processedCount++
            $processedFiles += $fileName
            
            # 変換した.xlsxファイルを削除（必要に応じてコメントアウト）
            # Remove-Item $xlsxPath -Force
        }
        catch {
            $errorMessage = "ファイル処理エラー ($fileName): $_"
            Write-Log -Message $errorMessage -LogPath $logPath -LogLevel "ERROR"
            Write-Error $errorMessage
            $errorDetails += $errorMessage
            $errorCount++
        }
    }
    
    # タイムスタンプを保存
    Save-Timestamps -Timestamps $timestamps -TimestampFilePath $Config.TimestampFile -Verbose

    # ステップ4: データ同期処理（追加・更新・削除）
    if ($allExcelData.Count -gt 0) {
        Write-Host "`n[4/4] データ同期処理開始（総件数: $($allExcelData.Count)）..."
        
        try {
            # DataSyncManagerモジュールを使用してデータ同期を実行
            $syncResult = Sync-DataWithApi `
                -SourceData $allExcelData `
                -ApiUri $Config.Api.Uri `
                -ApiHeaders $Config.Api.Headers `
                -AppId $Config.AppId `
                -LogPath $logPath `
                -TimeoutSec $Config.Api.TimeoutSec
            
            if ($syncResult.Success) {
                # データ同期処理結果を保持
                $syncAddedCount = $syncResult.AddedCount
                $syncUpdatedCount = $syncResult.UpdatedCount
                $syncDeletedCount = $syncResult.DeletedCount
            }
            else {
                $errorCount += $syncResult.ErrorCount
                $errorDetails += $syncResult.ErrorMessages
                Write-Log -Message "データ同期処理でエラーが発生しました（エラー数: $($syncResult.ErrorCount)）" -LogPath $logPath -LogLevel "ERROR"
                
                # エラー詳細を表示
                if ($syncResult.ErrorMessages.Count -gt 0) {
                    Write-Host "`n  エラー詳細:"
                    foreach ($errorMsg in $syncResult.ErrorMessages) {
                        Write-Host "    - $errorMsg"
                        Write-Log -Message "エラー詳細: $errorMsg" -LogPath $logPath -LogLevel "ERROR"
                    }
                }
                
                # ロールバック対象がある場合はログに記録
                if ($syncResult.RollbackTargets.Count -gt 0) {
                    Write-Log -Message "ロールバック対象: $($syncResult.RollbackTargets.Count) 件" -LogPath $logPath -LogLevel "ERROR"
                }
            }
        }
        catch {
            $syncError = "データ同期処理で例外が発生しました: $_"
            Write-Log -Message $syncError -LogPath $logPath -LogLevel "ERROR"
            Write-Error $syncError
            $errorDetails += $syncError
            $errorCount++
        }
    }
    
    # 処理完了ログ（データ同期処理結果も含める）
    $syncInfo = ""
    if ($allExcelData.Count -gt 0) {
        $syncInfo = ", 追加: $syncAddedCount 件, 更新: $syncUpdatedCount 件, 削除: $syncDeletedCount 件"
    }
    $filesInfo = ""
    if ($processedFiles.Count -gt 0) {
        $filesInfo = ", 更新ファイル: $($processedFiles -join ', ')"
    }
    Write-Log -Message "全処理完了（処理: $processedCount 件, スキップ: $skippedCount 件, エラー: $errorCount 件, データ総件数: $($allExcelData.Count)$syncInfo$filesInfo）" -LogPath $logPath -LogLevel "INFO"
    
    # エラーが発生した場合はメール送信
    if ($errorCount -gt 0) {
        Write-Host "エラーが発生しました。メールを送信します..."
        
        $emailSubject = "【エラー通知】Excelファイル処理でエラーが発生しました"
        $emailBody = @"
処理中に $errorCount 件のエラーが発生しました。

エラー詳細:
$($errorDetails -join "`n`n")

ログファイル: $logPath
実行日時: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
        
        try {
            # Microsoft Graph PowerShellを使用してメール送信
            $emailParams = @{
                Subject = $emailSubject
                Body    = $emailBody
                To      = $Config.Email.To
                From    = $Config.Email.From
            }
            
            # アプリケーション認証の設定がある場合は追加
            if ($Config.Email.TenantId -and $Config.Email.ClientId -and $Config.Email.ClientSecret) {
                $emailParams['TenantId'] = $Config.Email.TenantId
                $emailParams['ClientId'] = $Config.Email.ClientId
                $emailParams['ClientSecret'] = $Config.Email.ClientSecret
            }
            
            Send-ErrorEmail @emailParams -Verbose
            
            Write-Log -Message "エラーメール送信完了" -LogPath $logPath -LogLevel "INFO"
        }
        catch {
            $emailError = "メール送信エラー: $_"
            Write-Log -Message $emailError -LogPath $logPath -LogLevel "ERROR"
            Write-Error $emailError
        }
    }
    else {
        Write-Host "全処理が正常に完了しました。"
    }
}
catch {
    $fatalError = "致命的なエラーが発生しました: $_"
    Write-Log -Message $fatalError -LogPath $logPath -LogLevel "ERROR"
    Write-Error $fatalError
    
    # 致命的なエラーの場合もメール送信
    try {
        $emailParams = @{
            Subject = "【致命的エラー】Excelファイル処理スクリプト"
            Body    = $fatalError
            To      = $Config.Email.To
            From    = $Config.Email.From
        }
        
        # アプリケーション認証の設定がある場合は追加
        if ($Config.Email.TenantId -and $Config.Email.ClientId -and $Config.Email.ClientSecret) {
            $emailParams['TenantId'] = $Config.Email.TenantId
            $emailParams['ClientId'] = $Config.Email.ClientId
            $emailParams['ClientSecret'] = $Config.Email.ClientSecret
        }
        
        Send-ErrorEmail @emailParams
    }
    catch {
        Write-Error "メール送信も失敗しました: $_"
    }
    
    exit 1
}

