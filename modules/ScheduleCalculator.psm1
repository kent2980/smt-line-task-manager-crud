# ScheduleCalculator.psm1
# App86のsub_scheduleから予定開始・終了・休憩時間を算出する純粋計算モジュール

$script:JstOffset = [TimeSpan]::FromHours(9)
$script:ScheduleStartHour = 8
$script:ScheduleStartMinute = 30
$script:BreakDefinitions = @(
    [PSCustomObject]@{ StartHour = 10; StartMinute = 0; EndHour = 10; EndMinute = 10 },
    [PSCustomObject]@{ StartHour = 12; StartMinute = 0; EndHour = 12; EndMinute = 40 },
    [PSCustomObject]@{ StartHour = 15; StartMinute = 0; EndHour = 15; EndMinute = 10 }
)

function Get-FieldValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Container,

        [Parameter(Mandatory = $true)]
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

function ConvertTo-HourValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context,

        [Parameter(Mandatory = $false)]
        [switch]$AllowEmpty
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        if ($AllowEmpty) {
            return 0.0
        }

        throw "$Context が空です。"
    }

    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
        $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $numericValue = [double]$Value
        if ($numericValue -lt 0) {
            throw "$Context に負数は指定できません: $Value"
        }
        return $numericValue
    }

    $text = ([string]$Value).Trim()
    $match = [regex]::Match($text, '^([-+]?\d+(?:[\.,]\d+)?)\s*[Hh]?$')
    if (-not $match.Success) {
        throw "$Context の形式が不正です: $Value"
    }

    $normalized = $match.Groups[1].Value.Replace(',', '.')
    $parsed = 0.0
    $isParsed = [double]::TryParse(
        $normalized,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$parsed
    )
    if (-not $isParsed -or $parsed -lt 0) {
        throw "$Context の数値変換に失敗しました: $Value"
    }

    return $parsed
}

function ConvertTo-SortNumber {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        throw "$Context が空です。"
    }

    $parsed = 0.0
    $isParsed = [double]::TryParse(
        ([string]$Value).Trim(),
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$parsed
    )
    if (-not $isParsed) {
        throw "$Context の数値変換に失敗しました: $Value"
    }

    return $parsed
}

function New-JstDateTimeOffset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Date,

        [Parameter(Mandatory = $true)]
        [int]$Hour,

        [Parameter(Mandatory = $true)]
        [int]$Minute
    )

    $parsedDate = [DateTime]::MinValue
    $validDate = [DateTime]::TryParseExact(
        $Date,
        'yyyy-MM-dd',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsedDate
    )
    if (-not $validDate) {
        throw "sub_schedule_date の形式が不正です: $Date"
    }

    return [DateTimeOffset]::new(
        $parsedDate.Year,
        $parsedDate.Month,
        $parsedDate.Day,
        $Hour,
        $Minute,
        0,
        $script:JstOffset
    )
}

function Add-WorkDurationWithBreaks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$StartAt,

        [Parameter(Mandatory = $true)]
        [double]$WorkMinutes,

        [Parameter(Mandatory = $true)]
        [string]$ScheduleDate
    )

    if ($WorkMinutes -lt 0) {
        throw "作業時間に負数は指定できません: $WorkMinutes"
    }

    $cursor = $StartAt
    $remainingMinutes = $WorkMinutes
    $breakMinutes = 0
    $epsilon = 0.000001

    foreach ($definition in $script:BreakDefinitions) {
        $breakStart = New-JstDateTimeOffset -Date $ScheduleDate -Hour $definition.StartHour -Minute $definition.StartMinute
        $breakEnd = New-JstDateTimeOffset -Date $ScheduleDate -Hour $definition.EndHour -Minute $definition.EndMinute

        if ($cursor -ge $breakEnd) {
            continue
        }

        # 前工程の終了などで休憩内から始まる場合は、残りの休憩を先に消化する。
        if ($cursor -ge $breakStart -and $cursor -lt $breakEnd) {
            $delay = ($breakEnd - $cursor).TotalMinutes
            $breakMinutes += [int][Math]::Round($delay)
            $cursor = $breakEnd
        }

        if ($remainingMinutes -le $epsilon) {
            break
        }

        if ($cursor -lt $breakStart) {
            $availableMinutes = ($breakStart - $cursor).TotalMinutes

            if ($remainingMinutes -lt ($availableMinutes - $epsilon)) {
                $cursor = $cursor.AddMinutes($remainingMinutes)
                $remainingMinutes = 0
                break
            }

            if ([Math]::Abs($remainingMinutes - $availableMinutes) -le $epsilon) {
                # 名目終了が休憩開始と完全一致した場合も、その休憩を当該予定に含める。
                $cursor = $breakEnd
                $remainingMinutes = 0
                $breakMinutes += [int][Math]::Round(($breakEnd - $breakStart).TotalMinutes)
                break
            }

            $remainingMinutes -= $availableMinutes
            $cursor = $breakEnd
            $breakMinutes += [int][Math]::Round(($breakEnd - $breakStart).TotalMinutes)
        }
    }

    if ($remainingMinutes -gt $epsilon) {
        $cursor = $cursor.AddMinutes($remainingMinutes)
    }

    return [PSCustomObject]@{
        EndAt        = $cursor
        BreakMinutes = $breakMinutes
    }
}

function Get-GroupKey {
    param(
        [string]$Date,
        [string]$LineName
    )

    # 日付とライン名はいずれもApp86の管理値であり、区切りには通常値に含まれないパイプを使う。
    return "$Date|$LineName"
}

function Get-SelectedGroupKeySet {
    param(
        [Parameter(Mandatory = $false)]
        [array]$Groups
    )

    if ($null -eq $Groups -or $Groups.Count -eq 0) {
        return $null
    }

    $set = @{}
    foreach ($group in $Groups) {
        $date = [string](Get-FieldValue -Container $group -FieldName 'Date')
        $lineName = [string](Get-FieldValue -Container $group -FieldName 'LineName')
        if ([string]::IsNullOrWhiteSpace($date) -or [string]::IsNullOrWhiteSpace($lineName)) {
            continue
        }

        $set[(Get-GroupKey -Date $date.Trim() -LineName $lineName.Trim())] = $true
    }

    return $set
}

function Get-App86ScheduleCalculations {
    <#
    .SYNOPSIS
    App86レコードを日付×ラインでグループ化し、予定開始・終了・休憩時間を計算します。

    .DESCRIPTION
    Groupsを省略した場合は全予定を厳密に検証して計算します。
    Groupsを指定した場合は対象グループに含まれる行だけを厳密に検証し、無関係な過去データの不備で
    通常同期の部分再計算が停止しないようにします。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Records,

        [Parameter(Mandatory = $false)]
        [array]$Groups
    )

    $selectedGroupKeys = Get-SelectedGroupKeySet -Groups $Groups
    $isPartialCalculation = $null -ne $selectedGroupKeys
    $itemsByGroup = @{}

    foreach ($record in $Records) {
        $recordType = [string](Get-FieldValue -Container $record -FieldName 'record_type')
        if ($recordType.Trim().ToUpperInvariant() -eq 'SETTING') {
            continue
        }

        $rows = Get-FieldValue -Container $record -FieldName 'sub_schedule'
        if ($null -eq $rows -or @($rows).Count -eq 0) {
            continue
        }

        $lineName = [string](Get-FieldValue -Container $record -FieldName 'line_name')
        if ([string]::IsNullOrWhiteSpace($lineName)) {
            if ($isPartialCalculation) {
                continue
            }
            $recordIdForError = [string](Get-FieldValue -Container $record -FieldName '$id')
            throw "レコードID $recordIdForError の line_name が空です。"
        }
        $lineName = $lineName.Trim()

        $recordId = [string](Get-FieldValue -Container $record -FieldName '$id')
        $revision = [string](Get-FieldValue -Container $record -FieldName '$revision')
        $recordIndex = $null
        $recordIdNumber = $null
        $rowPosition = 0

        foreach ($row in @($rows)) {
            $rowValue = $row.value
            $date = [string](Get-FieldValue -Container $rowValue -FieldName 'sub_schedule_date')

            if ([string]::IsNullOrWhiteSpace($date)) {
                if ($isPartialCalculation) {
                    $rowPosition++
                    continue
                }
                throw "レコードID $recordId / テーブル位置 $rowPosition の sub_schedule_date が空です。"
            }
            $date = $date.Trim()

            $groupKey = Get-GroupKey -Date $date -LineName $lineName
            if ($isPartialCalculation -and -not $selectedGroupKeys.ContainsKey($groupKey)) {
                $rowPosition++
                continue
            }

            # ここからは実際の計算対象行なので、必要な識別子・値を厳密に検証する。
            if ([string]::IsNullOrWhiteSpace($recordId)) {
                throw '予定計算対象レコードの$idを取得できません。'
            }

            $rowId = [string](Get-FieldValue -Container $row -FieldName 'id')
            if ([string]::IsNullOrWhiteSpace($rowId)) {
                throw "レコードID $recordId の sub_schedule 行IDを取得できません。"
            }

            New-JstDateTimeOffset -Date $date -Hour 0 -Minute 0 | Out-Null

            if ($null -eq $recordIndex) {
                $recordIndex = ConvertTo-SortNumber `
                    -Value (Get-FieldValue -Container $record -FieldName 'index') `
                    -Context "レコードID $recordId の index"
            }

            if ($null -eq $recordIdNumber) {
                $parsedRecordId = 0.0
                $recordIdNumber = [double]::MaxValue
                if ([double]::TryParse(
                    $recordId,
                    [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$parsedRecordId
                )) {
                    $recordIdNumber = $parsedRecordId
                }
            }

            $productionHours = ConvertTo-HourValue `
                -Value (Get-FieldValue -Container $rowValue -FieldName '生産時間') `
                -Context "レコードID $recordId / $date / 生産時間"
            $changeHours = ConvertTo-HourValue `
                -Value (Get-FieldValue -Container $rowValue -FieldName 'sub_change_time') `
                -Context "レコードID $recordId / $date / sub_change_time" `
                -AllowEmpty

            $item = [PSCustomObject]@{
                RecordId        = $recordId
                RecordIdNumber  = $recordIdNumber
                Revision        = $revision
                RowId           = $rowId
                RowPosition     = $rowPosition
                Date            = $date
                LineName        = $lineName
                RecordIndex     = $recordIndex
                ProductionHours = $productionHours
                ChangeHours     = $changeHours
            }

            if (-not $itemsByGroup.ContainsKey($groupKey)) {
                $itemsByGroup[$groupKey] = @()
            }
            $itemsByGroup[$groupKey] += $item
            $rowPosition++
        }
    }

    $calculations = @()

    foreach ($groupKey in $itemsByGroup.Keys) {
        $groupItems = @($itemsByGroup[$groupKey]) | Sort-Object `
            @{ Expression = { $_.RecordIndex }; Ascending = $true }, `
            @{ Expression = { $_.RecordIdNumber }; Ascending = $true }, `
            @{ Expression = { $_.RecordId }; Ascending = $true }, `
            @{ Expression = { $_.RowPosition }; Ascending = $true }

        if ($groupItems.Count -eq 0) {
            continue
        }

        $cursor = New-JstDateTimeOffset `
            -Date $groupItems[0].Date `
            -Hour $script:ScheduleStartHour `
            -Minute $script:ScheduleStartMinute

        foreach ($item in $groupItems) {
            $workMinutes = ($item.ProductionHours + $item.ChangeHours) * 60.0
            $startAt = $cursor
            $advance = Add-WorkDurationWithBreaks `
                -StartAt $startAt `
                -WorkMinutes $workMinutes `
                -ScheduleDate $item.Date
            $cursor = $advance.EndAt

            $calculations += [PSCustomObject]@{
                RecordId        = $item.RecordId
                Revision        = $item.Revision
                RowId           = $item.RowId
                Date            = $item.Date
                LineName        = $item.LineName
                RecordIndex     = $item.RecordIndex
                RowPosition     = $item.RowPosition
                ProductionHours = $item.ProductionHours
                ChangeHours     = $item.ChangeHours
                StartAt         = $startAt
                EndAt           = $advance.EndAt
                BreakMinutes    = $advance.BreakMinutes
            }
        }
    }

    return $calculations
}

function ConvertTo-KintoneDateTimeValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$Value
    )

    return $Value.UtcDateTime.ToString(
        "yyyy-MM-dd'T'HH:mm:ss'Z'",
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

Export-ModuleMember -Function `
    ConvertTo-HourValue, `
    Add-WorkDurationWithBreaks, `
    Get-App86ScheduleCalculations, `
    ConvertTo-KintoneDateTimeValue
