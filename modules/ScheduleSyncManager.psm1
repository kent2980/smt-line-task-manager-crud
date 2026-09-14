# ScheduleSyncManager.psm1
# App86の最新レコードを取得し、予定開始・終了・休憩時間を書き戻すモジュール

$script:ScheduleUpdateBatchSize = 100
$script:ScheduleFetchBatchSize = 500
$script:WritableScheduleFields = @(
    'sub_schedule_date',
    'sub_lot_volume',
    'sub_index',
    'sub_change_time',
    '有効判定',
    '予定開始日時',
    '予定終了日時',
    '休憩時間'
)

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

function Get-ScheduleGroupMapKey {
    param(
        [string]$Date,
        [string]$LineName
    )

    return "$Date|$LineName"
}

function ConvertTo-ScheduleRangeDate {
    param([object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $parsed = [DateTime]::MinValue
    $isValid = [DateTime]::TryParseExact(
        ([string]$Value).Trim(),
        'yyyy-MM-dd',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed
    )
    if (-not $isValid) {
        return $null
    }

    return $parsed.Date
}

function Get-SourceScheduleDateRanges {
    param([array]$SourceData)

    $rangesByKey = [ordered]@{}
    foreach ($source in @($SourceData)) {
        $startValue = Get-KintoneFieldValue -Container $source -FieldName 'sync_date_range_start'
        $endValue = Get-KintoneFieldValue -Container $source -FieldName 'sync_date_range_end'
        $start = ConvertTo-ScheduleRangeDate -Value $startValue
        $end = ConvertTo-ScheduleRangeDate -Value $endValue

        if ($null -eq $start -or $null -eq $end) {
            $dates = @()
            $rows = Get-KintoneFieldValue -Container $source -FieldName 'sub_schedule'
            foreach ($row in @($rows)) {
                $rowValue = if ($row.PSObject.Properties['value']) { $row.value } else { $row }
                $date = ConvertTo-ScheduleRangeDate -Value (Get-KintoneFieldValue -Container $rowValue -FieldName 'sub_schedule_date')
                if ($null -ne $date) {
                    $dates += $date
                }
            }

            if ($dates.Count -eq 0) {
                continue
            }

            $sorted = @($dates | Sort-Object)
            $start = $sorted[0]
            $end = $sorted[$sorted.Count - 1]
        }

        if ($start -gt $end) {
            continue
        }

        $key = '{0}|{1}' -f $start.ToString('yyyy-MM-dd'), $end.ToString('yyyy-MM-dd')
        $rangesByKey[$key] = [PSCustomObject]@{ Start = $start; End = $end }
    }

    return @($rangesByKey.Values)
}

function Test-ScheduleDateInRanges {
    param(
        [object]$DateValue,
        [array]$DateRanges
    )

    $date = ConvertTo-ScheduleRangeDate -Value $DateValue
    if ($null -eq $date) {
        return $false
    }

    foreach ($range in @($DateRanges)) {
        if ($date -ge $range.Start -and $date -le $range.End) {
            return $true
        }
    }

    return $false
}

function Get-App86ScheduleRecords {
    <#
    .SYNOPSIS
    App86の生産レコードをindex、$id順で取得します。

    .DESCRIPTION
    lot_numberが空のレコードは異常データとして予定再計算対象から除外します。
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
        $query = 'lot_number != "" order by index asc, $id asc limit {0} offset {1}' -f $BatchSize, $offset
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
    param([object]$Item)

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
    今回追加する新予定と、Excel日付範囲内で無効化される既存予定から再計算対象を抽出します。
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

    $groupMap = [ordered]@{}

    foreach ($source in $SourceData) {
        foreach ($group in @(Get-SourceScheduleGroups -Item $source)) {
            $mapKey = Get-ScheduleGroupMapKey -Date $group.Date -LineName $group.LineName
            $groupMap[$mapKey] = $group
        }
    }

    # 更新対象条件と同じ日付範囲で、kintone側の既存予定を影響グループへ追加する。
    $dateRanges = @(Get-SourceScheduleDateRanges -SourceData $SourceData)
    if ($dateRanges.Count -gt 0) {
        foreach ($target in $TargetData) {
            $lineName = [string](Get-KintoneFieldValue -Container $target -FieldName 'line_name')
            if ([string]::IsNullOrWhiteSpace($lineName)) {
                continue
            }
            $lineName = $lineName.Trim()

            $rows = Get-KintoneFieldValue -Container $target -FieldName 'sub_schedule'
            foreach ($row in @($rows)) {
                $rowValue = if ($row.PSObject.Properties['value']) { $row.value } else { $row }
                $date = [string](Get-KintoneFieldValue -Container $rowValue -FieldName 'sub_schedule_date')
                if ([string]::IsNullOrWhiteSpace($date)) {
                    continue
                }

                $date = $date.Trim()
                if (-not (Test-ScheduleDateInRanges -DateValue $date -DateRanges $dateRanges)) {
                    continue
                }

                $group = [PSCustomObject]@{
                    Date     = $date
                    LineName = $lineName
                }
                $mapKey = Get-ScheduleGroupMapKey -Date $date -LineName $lineName
                $groupMap[$mapKey] = $group
            }
        }
    }

    return @($groupMap.Values)
}

function Copy-WritableScheduleRowValues {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RowValue
    )

    $payload = [ordered]@{}
    foreach ($fieldCode in $script:WritableScheduleFields) {
        if ($null -eq $RowValue.PSObject.Properties[$fieldCode]) {
            continue
        }

        $payload[$fieldCode] = [PSCustomObject]@{
            value = Get-KintoneFieldValue -Container $RowValue -FieldName $fieldCode
        }
    }

    # Calc「生産時間」は読み取り専用なので意図的にpayloadへ含めない。
    return $payload
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

            $rowValue = $row.value
            $rowPayloadValue = Copy-WritableScheduleRowValues -RowValue $rowValue

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

            # 計算対象外（False）の行もrow IDと現在値を含めて保持し、履歴を削除・上書きしない。
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

        $calculatorParams = @{ Records = $records }
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