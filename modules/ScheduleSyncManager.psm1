# ScheduleSyncManager.psm1
# App86の最新レコードを取得し、予定開始・終了・休憩時間を書き戻すモジュール

$script:ScheduleUpdateBatchSize = 100
$script:ScheduleFetchBatchSize = 500

function Ensure-ScheduleCalculatorLoaded {
    if (Get-Command Get-App86ScheduleCalculations -ErrorAction SilentlyContinue) {
        return
    }

    $calculatorPath = Join-Path $PSScriptRoot 'ScheduleCalculator.psm1'
    Import-Module $calculatorPath -Force -ErrorAction Stop
}

function Write-ScheduleLog {
    param(
        [string]$Message,
        [string]$LogPath,
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$LogLevel = 'INFO'
    )

    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
        Write-Log -Message $Message -LogPath $LogPath -LogLevel $LogLevel
    }
}

function Get-KintoneFieldValue {
    param(
        [object]$Container,
        [string]$FieldName
    )

    if ($null -eq $Container) {
        return $null
    }

    $property = $Container.PSObject.Properties[$FieldName]
    if ($null -eq $property) {
        return $null
    }

    $value = $property.Value
    if ($null -ne $value -and $null -ne $value.PSObject -and $value.PSObject.Properties['value']) {
        return $value.value
    }

    return $value
}

function Get-App86ScheduleRecords {
    <#
    .SYNOPSIS
    App86の全レコードをindex、$id順で取得します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApiUri,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApiHeaders,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 500
    )

    $records = @()
    $offset = 0

    do {
        $query = "order by index asc, `$id asc limit $BatchSize offset $offset"
        $encodedApp = [System.Uri]::EscapeDataString([string]$AppId)
        $encodedQuery = [System.Uri]::EscapeDataString($query)
        $requestUri = "{0}?app={1}&query={2}&totalCount=true" -f $ApiUri, $encodedApp, $encodedQuery

        $response = Get-ApiData `
            -Uri $requestUri `
            -Method 'GET' `
            -Headers $ApiHeaders `
            -TimeoutSec $TimeoutSec

        $pageRecords = @()
        if ($null -eq $response) {
            $pageRecords = @()
        }
        elseif ($response -is [System.Array]) {
            $pageRecords = @($response)
        }
        elseif ($response.PSObject.Properties['records']) {
            $pageRecords = @($response.records)
        }

        if ($pageRecords.Count -gt 0) {
            $records += $pageRecords
        }

        if ($pageRecords.Count -lt $BatchSize) {
            break
        }

        $offset += $BatchSize
    } while ($true)

    return $records
}

function Get-SourceScheduleGroups {
    param(
        [object]$Item
    )

    $groups = @()
    $lineName = [string](Get-KintoneFieldValue -Container $Item -FieldName 'line_name')
    $rows = Get-KintoneFieldValue -Container $Item -FieldName 'sub_schedule'
    if ([string]::IsNullOrWhiteSpace($lineName) -or $null -eq $rows) {
        return $groups
    }

    foreach ($row in @($rows)) {
        $rowValue = if ($row.PSObject.Properties['value']) { $row.value } else { $row }
        $date = [string](Get-KintoneFieldValue -Container $rowValue -FieldName 'sub_schedule_date')
        if (-not [string]::IsNullOrWhiteSpace($date)) {
            $groups += [PSCustomObject]@{
                Date     = $date.Trim()
                LineName = $lineName.Trim()
            }
        }
    }

    return $groups
}

function Get-AffectedScheduleGroups {
    <#
    .SYNOPSIS
    今回同期される更新元と同期前App86から、再計算が必要な日付×ラインを抽出します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$SourceData,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$TargetData
    )

    $targetByKey = @{}
    foreach ($target in $TargetData) {
        $targetKey = [string](Get-KintoneFieldValue -Container $target -FieldName 'line_lot_number')
        if (-not [string]::IsNullOrWhiteSpace($targetKey)) {
            $targetByKey[$targetKey.Trim()] = $target
        }
    }

    $groupMap = [ordered]@{}
    foreach ($source in $SourceData) {
        $sourceKey = [string](Get-KintoneFieldValue -Container $source -FieldName 'line_lot_number')

        foreach ($group in @(Get-SourceScheduleGroups -Item $source)) {
            $mapKey = "$($group.Date)`u001f$($group.LineName)"
            $groupMap[$mapKey] = $group
        }

        if (-not [string]::IsNullOrWhiteSpace($sourceKey) -and $targetByKey.ContainsKey($sourceKey.Trim())) {
            foreach ($group in @(Get-SourceScheduleGroups -Item $targetByKey[$sourceKey.Trim()])) {
                $mapKey = "$($group.Date)`u001f$($group.LineName)"
                $groupMap[$mapKey] = $group
            }
        }
    }

    return @($groupMap.Values)
}

function ConvertTo-ScheduleUpdateRecords {
    param(
        [array]$Records,
        [array]$Calculations
    )

    Ensure-ScheduleCalculatorLoaded

    $recordsById = @{}
    foreach ($record in $Records) {
        $recordId = [string](Get-KintoneFieldValue -Container $record -FieldName '$id')
        if (-not [string]::IsNullOrWhiteSpace($recordId)) {
            $recordsById[$recordId] = $record
        }
    }

    $calculationsByRecord = @{}
    foreach ($calculation in $Calculations) {
        if (-not $calculationsByRecord.ContainsKey($calculation.RecordId)) {
            $calculationsByRecord[$calculation.RecordId] = @{}
        }
        $calculationsByRecord[$calculation.RecordId][$calculation.RowId] = $calculation
    }

    $updates = @()
    foreach ($recordId in $calculationsByRecord.Keys) {
        if (-not $recordsById.ContainsKey($recordId)) {
            throw "予定計算結果に対応するレコードを取得できません: $recordId"
        }

        $record = $recordsById[$recordId]
        $rows = @(Get-KintoneFieldValue -Container $record -FieldName 'sub_schedule')
        $rowCalculations = $calculationsByRecord[$recordId]
        $payloadRows = @()

        foreach ($row in $rows) {
            $rowId = [string](Get-KintoneFieldValue -Container $row -FieldName 'id')
            if ([string]::IsNullOrWhiteSpace($rowId)) {
                throw "レコードID $recordId のsub_schedule行IDを取得できません。"
            }

            $rowPayloadValue = [ordered]@{}
            if ($rowCalculations.ContainsKey($rowId)) {
                $calculation = $rowCalculations[$rowId]
                $rowPayloadValue['予定開始日時'] = [PSCustomObject]@{
                    value = ConvertTo-KintoneDateTimeValue -Value $calculation.StartAt
                }
                $rowPayloadValue['予定終了日時'] = [PSCustomObject]@{
                    value = ConvertTo-KintoneDateTimeValue -Value $calculation.EndAt
                }
                $rowPayloadValue['休憩時間'] = [PSCustomObject]@{
                    value = [string]$calculation.BreakMinutes
                }
            }

            # テーブルをPUTする場合、既存行を落とさないよう全行IDを必ず含める。
            $payloadRows += [PSCustomObject]@{
                id    = $rowId
                value = [PSCustomObject]$rowPayloadValue
            }
        }

        $revision = [string](Get-KintoneFieldValue -Container $record -FieldName '$revision')
        $update = [ordered]@{
            id     = $recordId
            record = [PSCustomObject]@{
                sub_schedule = [PSCustomObject]@{
                    value = $payloadRows
                }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($revision)) {
            $update['revision'] = $revision
        }

        $updates += [PSCustomObject]$update
    }

    return $updates
}

function Invoke-App86ScheduleRecalculation {
    <#
    .SYNOPSIS
    App86の最新Calc値を再取得して予定開始日時・予定終了日時・休憩時間を更新します。

    .DESCRIPTION
    Groupsを省略した場合は全日付×全ラインを再計算します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApiUri,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApiHeaders,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30,

        [Parameter(Mandatory = $false)]
        [array]$Groups
    )

    Ensure-ScheduleCalculatorLoaded

    $result = [PSCustomObject]@{
        Success            = $false
        UpdatedRecordCount = 0
        UpdatedRowCount    = 0
        ErrorMessages      = @()
    }

    try {
        $records = Get-App86ScheduleRecords `
            -ApiUri $ApiUri `
            -ApiHeaders $ApiHeaders `
            -AppId $AppId `
            -TimeoutSec $TimeoutSec `
            -BatchSize $script:ScheduleFetchBatchSize

        $calculatorParams = @{
            Records = $records
        }
        if ($null -ne $Groups -and $Groups.Count -gt 0) {
            $calculatorParams['Groups'] = $Groups
        }

        $calculations = @(Get-App86ScheduleCalculations @calculatorParams)
        $result.UpdatedRowCount = $calculations.Count

        if ($calculations.Count -eq 0) {
            Write-ScheduleLog -Message '予定再計算対象は0件でした。' -LogPath $LogPath -LogLevel 'INFO'
            $result.Success = $true
            return $result
        }

        $updates = @(ConvertTo-ScheduleUpdateRecords -Records $records -Calculations $calculations)
        for ($index = 0; $index -lt $updates.Count; $index += $script:ScheduleUpdateBatchSize) {
            $endIndex = [Math]::Min($index + $script:ScheduleUpdateBatchSize - 1, $updates.Count - 1)
            $batch = @($updates[$index..$endIndex])
            $body = @{
                app     = $AppId
                records = $batch
            } | ConvertTo-Json -Depth 20

            Send-ApiRequest `
                -Uri $ApiUri `
                -Method 'PUT' `
                -Body $body `
                -Headers $ApiHeaders `
                -TimeoutSec $TimeoutSec | Out-Null

            $result.UpdatedRecordCount += $batch.Count
        }

        $message = "予定再計算完了: レコード $($result.UpdatedRecordCount) 件 / 行 $($result.UpdatedRowCount) 件"
        Write-ScheduleLog -Message $message -LogPath $LogPath -LogLevel 'INFO'
        $result.Success = $true
        return $result
    }
    catch {
        $errorMessage = "予定再計算エラー: $_"
        Write-ScheduleLog -Message $errorMessage -LogPath $LogPath -LogLevel 'ERROR'
        $result.ErrorMessages += $errorMessage
        return $result
    }
}

Export-ModuleMember -Function `
    Get-App86ScheduleRecords, `
    Get-AffectedScheduleGroups, `
    Invoke-App86ScheduleRecalculation
