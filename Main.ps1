# Main.ps1
# メイン処理スクリプト

[CmdletBinding()]
param(
    # タイムスタンプ・日次同期済み判定を無視して全Excelを再読込し、全対象レコードを更新する。
    [Parameter(Mandatory = $false)]
    [switch]$ForceFullSync,

    # Excel同期を行わず、App86に現在登録されている予定だけを全件再計算する。
    [Parameter(Mandatory = $false)]
    [switch]$RecalculateScheduleOnly
)

if ($ForceFullSync -and $RecalculateScheduleOnly) {
    throw '-ForceFullSync と -RecalculateScheduleOnly は同時に指定できません。'
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulesPath = Join-Path $scriptDirectory 'modules'
$configPath = Join-Path $scriptDirectory 'Config.ps1'

if (-not (Test-Path $configPath)) {
    Write-Error "設定ファイルが見つかりません: $configPath"
    exit 1
}
$Config = . $configPath

$modules = @(
    'Utils',
    'TimestampManager',
    'ExcelConverter',
    'ExcelReader',
    'JsonConverter',
    'ApiClient',
    'Logger',
    'EmailSender',
    'DataSyncManager',
    'ScheduleCalculator',
    'ScheduleSyncManager'
)

foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath "$module.psm1"
    if (-not (Test-Path $modulePath)) {
        Write-Error "モジュールが見つかりません: $modulePath"
        exit 1
    }
    Import-Module $modulePath -Force
    Write-Verbose "モジュールを読み込みました: $module"
}

$logPath = Initialize-Logger -LogDirectory $Config.LogDirectory
$timestamps = Load-Timestamps -TimestampFilePath $Config.TimestampFile -Verbose

$errorCount = 0
$errorDetails = @()
$processedCount = 0
$skippedCount = 0
$syncAddedCount = 0
$syncUpdatedCount = 0
$scheduleUpdatedRecordCount = 0
$scheduleUpdatedRowCount = 0
$processedFiles = @()
$pendingTimestamps = @{}
$allExcelData = @()

$tempDir = $Config.DataDirectoryTemp
$xlsxDir = $Config.DataDirectoryXlsx
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}
if (-not (Test-Path $xlsxDir)) {
    New-Item -ItemType Directory -Path $xlsxDir -Force | Out-Null
}

$dailyFullSyncStateFile = Join-Path $tempDir 'last-full-sync-date.txt'
$dailyFullSyncDate = Get-Date -Format 'yyyy-MM-dd'
$isDailyFullSyncRequired = $true
$expectedFileCount = $Config.FileNumberEnd - $Config.FileNumberStart + 1

function Write-DailyFullSyncState {
    param(
        [string]$StateFile,
        [string]$DateValue
    )

    $stateTempFile = "$StateFile.$([Guid]::NewGuid().ToString('N')).tmp"
    try {
        $stateFileEncoding = [Text.UTF8Encoding]::new($false)
        [IO.File]::WriteAllText($stateTempFile, $DateValue, $stateFileEncoding)
        Move-Item -LiteralPath $stateTempFile -Destination $StateFile -Force
    }
    finally {
        if (Test-Path -LiteralPath $stateTempFile) {
            Remove-Item -LiteralPath $stateTempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Send-RunErrorNotification {
    param(
        [int]$ErrorCount,
        [array]$ErrorDetails,
        [string]$LogPath,
        [hashtable]$Config
    )

    if ($ErrorCount -le 0) {
        return
    }

    $emailParams = @{
        Subject = '【エラー通知】Excelファイル処理でエラーが発生しました'
        Body    = @"
処理中に $ErrorCount 件のエラーが発生しました。

エラー詳細:
$($ErrorDetails -join "`n`n")

ログファイル: $LogPath
実行日時: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
        To      = $Config.Email.To
        From    = $Config.Email.From
    }

    if ($Config.Email.TenantId -and $Config.Email.ClientId -and $Config.Email.ClientSecret) {
        $emailParams['TenantId'] = $Config.Email.TenantId
        $emailParams['ClientId'] = $Config.Email.ClientId
        $emailParams['ClientSecret'] = $Config.Email.ClientSecret
    }

    try {
        Send-ErrorEmail @emailParams -Verbose
        Write-Log -Message 'エラーメール送信完了' -LogPath $LogPath -LogLevel 'INFO'
    }
    catch {
        Write-Log -Message "メール送信エラー: $_" -LogPath $LogPath -LogLevel 'ERROR'
        Write-Error "メール送信エラー: $_"
    }
}

$scriptMutex = $null
$hasMutex = $false
try {
    $scriptMutex = New-Object System.Threading.Mutex($false, 'Global\smt-line-task-manager-crud-main')
    $hasMutex = $scriptMutex.WaitOne(0, $false)
}
catch {
    $hasMutex = $true
}

if (-not $hasMutex) {
    $duplicateRunMessage = '同一スクリプトが既に実行中のため、今回の実行をスキップしました。'
    Write-Log -Message $duplicateRunMessage -LogPath $logPath -LogLevel 'WARNING'
    Write-Host $duplicateRunMessage
    exit 0
}

try {
    # Excelを触らず、現在のApp86だけを使って全予定を再計算する保守モード。
    if ($RecalculateScheduleOnly) {
        Write-Host 'App86の予定を全件再計算します...'
        $scheduleResult = Invoke-App86ScheduleRecalculation `
            -ApiUri $Config.Api.Uri `
            -ApiHeaders $Config.Api.Headers `
            -AppId $Config.AppId `
            -LogPath $logPath `
            -TimeoutSec $Config.Api.TimeoutSec

        if (-not $scheduleResult.Success) {
            throw ($scheduleResult.ErrorMessages -join "`n")
        }

        Write-Host "予定再計算完了（レコード: $($scheduleResult.UpdatedRecordCount) 件, 行: $($scheduleResult.UpdatedRowCount) 件）"
        return
    }

    Test-ExcelComAvailability -Verbose

    if ($ForceFullSync) {
        $isDailyFullSyncRequired = $true
        $forceMessage = "強制全件同期として、対象の全 $expectedFileCount ファイルを処理します。"
        Write-Log -Message $forceMessage -LogPath $logPath -LogLevel 'INFO'
        Write-Host $forceMessage
    }
    else {
        try {
            if (Test-Path -LiteralPath $dailyFullSyncStateFile) {
                $stateFileEncoding = [Text.UTF8Encoding]::new($false, $true)
                $lastFullSyncDate = [IO.File]::ReadAllText($dailyFullSyncStateFile, $stateFileEncoding).Trim()
                $isDailyFullSyncRequired = $lastFullSyncDate -ne $dailyFullSyncDate
            }
        }
        catch {
            $stateReadWarning = "日次全件同期の完了状態を読み込めないため、全ファイルを処理します: $_"
            Write-Log -Message $stateReadWarning -LogPath $logPath -LogLevel 'WARNING'
            Write-Warning $stateReadWarning
            $isDailyFullSyncRequired = $true
        }

        if ($isDailyFullSyncRequired) {
            $dailyMessage = "本日初回の全件同期として、対象の全 $expectedFileCount ファイルを処理します。"
            Write-Log -Message $dailyMessage -LogPath $logPath -LogLevel 'INFO'
            Write-Host $dailyMessage
        }
    }

    for ($i = $Config.FileNumberStart; $i -le $Config.FileNumberEnd; $i++) {
        $fileName = $Config.FileNamePattern -f $i
        $fileNameXlsx = $Config.FileNamePatternXlsx -f $i
        $xlsPath = Join-Path $Config.DataDirectory $fileName
        $xlsxPath = Join-Path $xlsxDir $fileNameXlsx
        $xlsPathTemp = $null

        try {
            if (-not (Test-Path $xlsPath)) {
                $message = "ファイルが存在しません: $xlsPath"
                Write-Log -Message $message -LogPath $logPath -LogLevel 'WARNING'
                $errorDetails += $message
                $errorCount++
                continue
            }

            $normalizedXlsPath = [System.IO.Path]::GetFullPath($xlsPath)
            $savedTimestamp = $timestamps[$normalizedXlsPath]

            if ($ForceFullSync -or $isDailyFullSyncRequired) {
                $isUpdated = $true
            }
            else {
                $isUpdated = Test-FileUpdated `
                    -FilePath $xlsPath `
                    -SavedTimestamp $savedTimestamp `
                    -Timestamps $timestamps `
                    -Verbose
            }

            if (-not $isUpdated) {
                $skippedCount++
                continue
            }

            $timestampForSync = Get-FileTimestamp -FilePath $xlsPath
            $tempFileName = "{0}_{1}.xls" -f `
                [System.IO.Path]::GetFileNameWithoutExtension($fileName), `
                ([System.Guid]::NewGuid().ToString('N'))
            $xlsPathTemp = Join-Path $tempDir $tempFileName

            Write-Host "  [1/4] $fileName を一時ディレクトリへコピー中..."
            Copy-Item -Path $xlsPath -Destination $xlsPathTemp -Force

            Write-Host '  [2/4] .xls → .xlsx 変換中...'
            $xlsxPath = Convert-XlsToXlsx -XlsPath $xlsPathTemp -XlsxPath $xlsxPath -Verbose
            if (Test-Path -LiteralPath $xlsPathTemp) {
                Remove-Item -LiteralPath $xlsPathTemp -Force -ErrorAction SilentlyContinue
            }

            Write-Host '  [3/4] Excelファイル読み取り中...'
            $excelData = Read-ExcelData -XlsxPath $xlsxPath -Verbose
            if ($excelData -and $excelData.Count -gt 0) {
                $allExcelData += $excelData
            }
            else {
                $warningMessage = "空シートを検出したため同期対象から除外します: $fileName"
                Write-Log -Message $warningMessage -LogPath $logPath -LogLevel 'WARNING'
                Write-Warning $warningMessage
            }

            $processedCount++
            $processedFiles += $fileName
            $pendingTimestamps[$normalizedXlsPath] = $timestampForSync
        }
        catch {
            $message = "ファイル処理エラー ($fileName): $_"
            Write-Log -Message $message -LogPath $logPath -LogLevel 'ERROR'
            Write-Error $message
            $errorDetails += $message
            $errorCount++
            if ($xlsPathTemp -and (Test-Path -LiteralPath $xlsPathTemp)) {
                Remove-Item -LiteralPath $xlsPathTemp -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Milliseconds 300
        }
    }

    if ($allExcelData.Count -gt 0) {
        Write-Host "`n[4/4] データ同期処理開始（総件数: $($allExcelData.Count)）..."

        try {
            # 同期前の旧日程も再計算対象へ含めるため、同期直前のApp86を保持する。
            $beforeSyncRecords = @(Get-App86ScheduleRecords `
                -ApiUri $Config.Api.Uri `
                -ApiHeaders $Config.Api.Headers `
                -AppId $Config.AppId `
                -TimeoutSec $Config.Api.TimeoutSec)

            $affectedGroups = @(Get-AffectedScheduleGroups `
                -SourceData $allExcelData `
                -TargetData $beforeSyncRecords)

            $syncResult = Sync-DataWithApi `
                -SourceData $allExcelData `
                -ApiUri $Config.Api.Uri `
                -ApiHeaders $Config.Api.Headers `
                -AppId $Config.AppId `
                -LogPath $logPath `
                -TimeoutSec $Config.Api.TimeoutSec

            if (-not $syncResult.Success) {
                $errorCount += $syncResult.ErrorCount
                $errorDetails += $syncResult.ErrorMessages
                throw "データ同期処理に失敗しました。"
            }

            $syncAddedCount = $syncResult.AddedCount
            $syncUpdatedCount = $syncResult.UpdatedCount

            # Excel PUT後に再GETすることで、kintone Calc「生産時間」の最新値を使用する。
            if ($ForceFullSync) {
                $scheduleResult = Invoke-App86ScheduleRecalculation `
                    -ApiUri $Config.Api.Uri `
                    -ApiHeaders $Config.Api.Headers `
                    -AppId $Config.AppId `
                    -LogPath $logPath `
                    -TimeoutSec $Config.Api.TimeoutSec
            }
            elseif ($affectedGroups.Count -gt 0) {
                $scheduleResult = Invoke-App86ScheduleRecalculation `
                    -ApiUri $Config.Api.Uri `
                    -ApiHeaders $Config.Api.Headers `
                    -AppId $Config.AppId `
                    -LogPath $logPath `
                    -TimeoutSec $Config.Api.TimeoutSec `
                    -Groups $affectedGroups
            }
            else {
                $scheduleResult = [PSCustomObject]@{
                    Success            = $true
                    UpdatedRecordCount = 0
                    UpdatedRowCount    = 0
                    ErrorMessages      = @()
                }
            }

            if (-not $scheduleResult.Success) {
                $errorCount += [Math]::Max(1, $scheduleResult.ErrorMessages.Count)
                $errorDetails += $scheduleResult.ErrorMessages
                throw '予定再計算に失敗したため、ファイルタイムスタンプは確定しません。'
            }

            $scheduleUpdatedRecordCount = $scheduleResult.UpdatedRecordCount
            $scheduleUpdatedRowCount = $scheduleResult.UpdatedRowCount

            # Excel同期と予定再計算の両方が成功した場合だけ、ファイルタイムスタンプを確定する。
            foreach ($processedFilePath in $pendingTimestamps.Keys) {
                $timestamps[$processedFilePath] = $pendingTimestamps[$processedFilePath]
            }
            Save-Timestamps -Timestamps $timestamps -TimestampFilePath $Config.TimestampFile -Verbose

            if ($isDailyFullSyncRequired) {
                if ($processedCount -eq $expectedFileCount) {
                    Write-DailyFullSyncState -StateFile $dailyFullSyncStateFile -DateValue $dailyFullSyncDate
                    Write-Log `
                        -Message "本日の日次全件同期が完了しました: $dailyFullSyncDate" `
                        -LogPath $logPath `
                        -LogLevel 'INFO'
                }
                else {
                    $message = "日次全件同期は未完了です（処理: $processedCount / $expectedFileCount ファイル）。次回も全ファイルを処理します。"
                    Write-Log -Message $message -LogPath $logPath -LogLevel 'WARNING'
                    Write-Warning $message
                }
            }
        }
        catch {
            $syncError = "同期・予定再計算処理でエラーが発生しました: $_"
            Write-Log -Message $syncError -LogPath $logPath -LogLevel 'ERROR'
            Write-Error $syncError
            $errorDetails += $syncError
            $errorCount++
        }
    }

    $syncInfo = ''
    if ($allExcelData.Count -gt 0) {
        $syncInfo = ", 追加: $syncAddedCount 件, 更新: $syncUpdatedCount 件, 予定更新: $scheduleUpdatedRecordCount レコード / $scheduleUpdatedRowCount 行"
    }
    $filesInfo = if ($processedFiles.Count -gt 0) { ", 更新ファイル: $($processedFiles -join ', ')" } else { '' }

    $completionMessage = "全処理完了（処理: $processedCount 件, スキップ: $skippedCount 件, エラー: $errorCount 件, データ総件数: $($allExcelData.Count)$syncInfo$filesInfo）"
    Write-Log -Message $completionMessage -LogPath $logPath -LogLevel 'INFO'
    Write-Host $completionMessage

    if ($errorCount -gt 0) {
        Send-RunErrorNotification `
            -ErrorCount $errorCount `
            -ErrorDetails $errorDetails `
            -LogPath $logPath `
            -Config $Config
    }
    else {
        Write-Host '全処理が正常に完了しました。'
    }
}
catch {
    $fatalError = "致命的なエラーが発生しました: $_"
    Write-Log -Message $fatalError -LogPath $logPath -LogLevel 'ERROR'
    Write-Error $fatalError
    exit 1
}
finally {
    if ($scriptMutex -and $hasMutex) {
        try {
            $scriptMutex.ReleaseMutex()
            $scriptMutex.Dispose()
        }
        catch {
            Write-Verbose "二重起動防止ミューテックスの解放に失敗: $_"
        }
    }
}
