# DataSyncManager.psm1
# データ同期処理を行うモジュール（追加・更新）

$script:ScheduleWritableFields = @(
    'sub_schedule_date',
    'sub_lot_volume',
    'sub_index',
    'sub_change_time',
    '有効判定',
    '予定開始日時',
    '予定終了日時',
    '休憩時間'
)

$script:SourceMetadataFields = @(
    'sync_date_range_start',
    'sync_date_range_end'
)

function Get-DataFieldValue {
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

function Get-SubScheduleRows {
    param([object]$Item)

    $rows = Get-DataFieldValue -Container $Item -FieldName 'sub_schedule'
    if ($null -eq $rows) {
        return @()
    }

    return @($rows)
}

function Get-SubScheduleRowValue {
    param([object]$Row)

    if ($null -eq $Row) {
        return $null
    }

    if ($Row.PSObject.Properties['value']) {
        return $Row.value
    }

    return $Row
}

function ConvertTo-SyncDate {
    param(
        [object]$Value,
        [string]$Context
    )

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
        throw "$Context の日付形式が不正です: $Value"
    }

    return $parsed.Date
}

function Get-ItemSyncDateRange {
    param([object]$Item)

    $startValue = Get-DataFieldValue -Container $Item -FieldName 'sync_date_range_start'
    $endValue = Get-DataFieldValue -Container $Item -FieldName 'sync_date_range_end'

    if (-not [string]::IsNullOrWhiteSpace([string]$startValue) -and
        -not [string]::IsNullOrWhiteSpace([string]$endValue)) {
        $start = ConvertTo-SyncDate -Value $startValue -Context 'sync_date_range_start'
        $end = ConvertTo-SyncDate -Value $endValue -Context 'sync_date_range_end'
        if ($start -gt $end) {
            throw "同期日付範囲が逆転しています: $startValue ～ $endValue"
        }

        return [PSCustomObject]@{ Start = $start; End = $end }
    }

    # 既存テストや旧呼び出しとの互換用フォールバック。
    # 本番のExcelReaderは必ずヘッダー由来の範囲メタデータを付与する。
    $dates = @()
    foreach ($row in @(Get-SubScheduleRows -Item $Item)) {
        $rowValue = Get-SubScheduleRowValue -Row $row
        $dateValue = Get-DataFieldValue -Container $rowValue -FieldName 'sub_schedule_date'
        $date = ConvertTo-SyncDate -Value $dateValue -Context 'sub_schedule_date'
        if ($null -ne $date) {
            $dates += $date
        }
    }

    if ($dates.Count -eq 0) {
        return $null
    }

    $sorted = @($dates | Sort-Object)
    return [PSCustomObject]@{
        Start = $sorted[0]
        End   = $sorted[$sorted.Count - 1]
    }
}

function Get-SyncDateRanges {
    param([array]$SourceData)

    $rangesByKey = [ordered]@{}
    foreach ($item in @($SourceData)) {
        $range = Get-ItemSyncDateRange -Item $item
        if ($null -eq $range) {
            continue
        }

        $key = '{0}|{1}' -f $range.Start.ToString('yyyy-MM-dd'), $range.End.ToString('yyyy-MM-dd')
        $rangesByKey[$key] = $range
    }

    return @($rangesByKey.Values)
}

function Test-DateInRanges {
    param(
        [object]$DateValue,
        [array]$DateRanges
    )

    if ($null -eq $DateRanges -or $DateRanges.Count -eq 0) {
        return $false
    }

    $date = ConvertTo-SyncDate -Value $DateValue -Context 'sub_schedule_date'
    if ($null -eq $date) {
        return $false
    }

    foreach ($range in $DateRanges) {
        if ($date -ge $range.Start -and $date -le $range.End) {
            return $true
        }
    }

    return $false
}

function Test-TargetHasScheduleInRanges {
    param(
        [object]$Target,
        [array]$DateRanges
    )

    foreach ($row in @(Get-SubScheduleRows -Item $Target)) {
        $rowValue = Get-SubScheduleRowValue -Row $row
        $dateValue = Get-DataFieldValue -Container $rowValue -FieldName 'sub_schedule_date'
        if (Test-DateInRanges -DateValue $dateValue -DateRanges $DateRanges) {
            return $true
        }
    }

    return $false
}

function Get-LineLotNumberValue {
    <#
    .SYNOPSIS
    line_lot_numberを同期キーとして利用できる文字列へ正規化します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $rawValue = Get-DataFieldValue -Container $Item -FieldName 'line_lot_number'
    if ($null -eq $rawValue) {
        return $null
    }

    $key = ([string]$rawValue).Trim()
    if ([string]::IsNullOrWhiteSpace($key)) {
        return $null
    }

    return $key
}

function Compare-DataByLineLotNumber {
    <#
    .SYNOPSIS
    line_lot_number一致、またはExcel日付範囲内の既存sub_schedule行を条件に更新対象を抽出します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$SourceData,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$TargetData,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    try {
        $sourceKeys = @{}
        foreach ($item in $SourceData) {
            $key = Get-LineLotNumberValue -Item $item
            if ($null -eq $key) {
                throw '更新元データのline_lot_numberが空です'
            }
            if ($sourceKeys.ContainsKey($key)) {
                throw "更新元データ内でline_lot_numberが重複しています: $key"
            }
            $sourceKeys[$key] = $item
        }

        $targetKeys = @{}
        foreach ($item in $TargetData) {
            $key = Get-LineLotNumberValue -Item $item
            if ($null -eq $key) {
                throw '更新先データのline_lot_numberが空です'
            }
            if ($targetKeys.ContainsKey($key)) {
                throw "更新先データ内でline_lot_numberが重複しています: $key"
            }
            $targetKeys[$key] = $item
        }

        $toAdd = @()
        foreach ($key in $sourceKeys.Keys) {
            if (-not $targetKeys.ContainsKey($key)) {
                $toAdd += $sourceKeys[$key]
            }
        }

        $toUpdate = @()
        $updateKeys = @{}
        foreach ($key in $sourceKeys.Keys) {
            if (-not $targetKeys.ContainsKey($key)) {
                continue
            }

            $sourceItem = $sourceKeys[$key]
            $targetItem = $targetKeys[$key]
            $sourceRange = Get-ItemSyncDateRange -Item $sourceItem
            $dateRanges = if ($null -eq $sourceRange) { @() } else { @($sourceRange) }

            $toUpdate += [PSCustomObject]@{
                Source     = $sourceItem
                Target     = $targetItem
                Key        = $key
                UpdateMode = 'Full'
                DateRanges = $dateRanges
            }
            $updateKeys[$key] = $true
        }

        # 既存条件に加え、kintone側のsub_scheduleにExcel日付範囲内の既存行があれば更新対象にする。
        $allDateRanges = @(Get-SyncDateRanges -SourceData $SourceData)
        if ($allDateRanges.Count -gt 0) {
            foreach ($key in $targetKeys.Keys) {
                if ($updateKeys.ContainsKey($key)) {
                    continue
                }

                $targetItem = $targetKeys[$key]
                if (-not (Test-TargetHasScheduleInRanges -Target $targetItem -DateRanges $allDateRanges)) {
                    continue
                }

                $toUpdate += [PSCustomObject]@{
                    Source     = $null
                    Target     = $targetItem
                    Key        = $key
                    UpdateMode = 'ScheduleOnly'
                    DateRanges = $allDateRanges
                }
                $updateKeys[$key] = $true
            }
        }

        return [PSCustomObject]@{
            ToAdd     = $toAdd
            ToUpdate  = $toUpdate
            DateRanges = $allDateRanges
        }
    }
    catch {
        $errorMsg = "データ突合エラー: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
        throw $errorMsg
    }
}

function Copy-ExistingScheduleRowValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RowValue,

        [Parameter(Mandatory = $true)]
        [bool]$Deactivate
    )

    $payload = [ordered]@{}
    foreach ($fieldCode in $script:ScheduleWritableFields) {
        if ($null -eq $RowValue.PSObject.Properties[$fieldCode]) {
            continue
        }

        $payload[$fieldCode] = [PSCustomObject]@{
            value = Get-DataFieldValue -Container $RowValue -FieldName $fieldCode
        }
    }

    if ($Deactivate) {
        $payload['有効判定'] = [PSCustomObject]@{ value = 'False' }
    }

    return [PSCustomObject]$payload
}

function ConvertTo-NewScheduleRows {
    param([object]$SourceItem)

    if ($null -eq $SourceItem) {
        return @()
    }

    $rows = @()
    foreach ($sourceRow in @(Get-SubScheduleRows -Item $SourceItem)) {
        $rowValue = Get-SubScheduleRowValue -Row $sourceRow
        $payloadValue = [ordered]@{}
        foreach ($fieldCode in @('sub_schedule_date', 'sub_lot_volume', 'sub_index', 'sub_change_time')) {
            if ($null -eq $rowValue.PSObject.Properties[$fieldCode]) {
                continue
            }

            $payloadValue[$fieldCode] = [PSCustomObject]@{
                value = Get-DataFieldValue -Container $rowValue -FieldName $fieldCode
            }
        }
        $payloadValue['有効判定'] = [PSCustomObject]@{ value = 'True' }

        $rows += [PSCustomObject]@{
            value = [PSCustomObject]$payloadValue
        }
    }

    return $rows
}

function Merge-SubScheduleRows {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TargetItem,

        [Parameter(Mandatory = $false)]
        [object]$SourceItem,

        [Parameter(Mandatory = $false)]
        [array]$DateRanges
    )

    $payloadRows = @()
    foreach ($row in @(Get-SubScheduleRows -Item $TargetItem)) {
        $rowId = [string](Get-DataFieldValue -Container $row -FieldName 'id')
        if ([string]::IsNullOrWhiteSpace($rowId)) {
            throw '既存sub_schedule行の行IDを取得できません。物理削除を防ぐため更新を中止します。'
        }

        $rowValue = Get-SubScheduleRowValue -Row $row
        $dateValue = Get-DataFieldValue -Container $rowValue -FieldName 'sub_schedule_date'
        $deactivate = Test-DateInRanges -DateValue $dateValue -DateRanges $DateRanges

        $payloadRows += [PSCustomObject]@{
            id    = $rowId
            value = Copy-ExistingScheduleRowValue -RowValue $rowValue -Deactivate $deactivate
        }
    }

    $payloadRows += @(ConvertTo-NewScheduleRows -SourceItem $SourceItem)
    return $payloadRows
}

function Remove-SourceMetadataFields {
    param([object]$Record)

    foreach ($fieldName in $script:SourceMetadataFields) {
        if ($null -ne $Record.PSObject.Properties[$fieldName]) {
            $Record.PSObject.Properties.Remove($fieldName)
        }
    }
}

function ConvertTo-AddRecordPayload {
    param([object]$SourceItem)

    $record = ConvertTo-WrappedJsonObject -InputObject $SourceItem
    Remove-SourceMetadataFields -Record $record
    $record | Add-Member -MemberType NoteProperty -Name 'schedule_date' -Value ([PSCustomObject]@{
        value = Get-ScheduleDate -InputObject $SourceItem
    }) -Force
    $record | Add-Member -MemberType NoteProperty -Name 'sub_schedule' -Value ([PSCustomObject]@{
        value = @(ConvertTo-NewScheduleRows -SourceItem $SourceItem)
    }) -Force

    return $record
}

function ConvertTo-UpdateRecordPayload {
    param([object]$UpdateItem)

    $targetItem = $UpdateItem.Target
    $recordId = Get-LineLotNumberValue -Item $targetItem
    if ($null -eq $recordId) {
        throw "更新先のline_lot_numberを取得できません: $($UpdateItem.Key)"
    }

    $dateRanges = @($UpdateItem.DateRanges)
    $mergedRows = @(Merge-SubScheduleRows `
        -TargetItem $targetItem `
        -SourceItem $UpdateItem.Source `
        -DateRanges $dateRanges)

    if ($UpdateItem.UpdateMode -eq 'ScheduleOnly') {
        $record = [PSCustomObject]@{
            sub_schedule = [PSCustomObject]@{ value = $mergedRows }
        }
    }
    else {
        $record = ConvertTo-WrappedJsonObject -InputObject $UpdateItem.Source
        Remove-SourceMetadataFields -Record $record
        if ($null -ne $record.PSObject.Properties['line_lot_number']) {
            $record.PSObject.Properties.Remove('line_lot_number')
        }
        $record | Add-Member -MemberType NoteProperty -Name 'schedule_date' -Value ([PSCustomObject]@{
            value = Get-ScheduleDate -InputObject $UpdateItem.Source
        }) -Force
        $record | Add-Member -MemberType NoteProperty -Name 'sub_schedule' -Value ([PSCustomObject]@{
            value = $mergedRows
        }) -Force
    }

    return [PSCustomObject]@{
        updateKey = [PSCustomObject]@{
            field = 'line_lot_number'
            value = $recordId
        }
        record = $record
    }
}

function Sync-DataWithApi {
    <#
    .SYNOPSIS
    更新元データと更新先データを同期します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$SourceData,

        [Parameter(Mandatory = $true)]
        [string]$ApiUri,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApiHeaders,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [Parameter(Mandatory = $false)]
        [int]$GetBatchSize = 500,

        [Parameter(Mandatory = $false)]
        [int]$PostBatchSize = 100,

        [Parameter(Mandatory = $false)]
        [int]$PutBatchSize = 100
    )

    $result = [PSCustomObject]@{
        Success         = $false
        AddedCount      = 0
        UpdatedCount    = 0
        ErrorCount      = 0
        ErrorMessages   = @()
        RollbackTargets = @()
    }

    try {
        if ($null -eq $SourceData -or $SourceData.Count -eq 0) {
            $errorMsg = '更新元データが空です'
            Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
            $result.ErrorMessages += $errorMsg
            $result.ErrorCount++
            return $result
        }

        try {
            $targetData = Get-TargetDataFromApi `
                -ApiUri $ApiUri `
                -ApiHeaders $ApiHeaders `
                -AppId $AppId `
                -TimeoutSec $TimeoutSec `
                -BatchSize $GetBatchSize `
                -LogPath $LogPath
            if ($null -eq $targetData) {
                $targetData = @()
            }
        }
        catch {
            $errorMsg = "更新先データの取得に失敗しました: $_"
            Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
            $result.ErrorMessages += $errorMsg
            $result.ErrorCount++
            return $result
        }

        $targetRecords = $targetData
        if ($targetData.PSObject.Properties['records']) {
            $targetRecords = $targetData.records
        }

        $comparisonResult = Compare-DataByLineLotNumber `
            -SourceData $SourceData `
            -TargetData $targetRecords `
            -LogPath $LogPath

        if ($comparisonResult.ToAdd.Count -gt 0) {
            $addResult = Invoke-AddData `
                -DataToAdd $comparisonResult.ToAdd `
                -ApiUri $ApiUri `
                -ApiHeaders $ApiHeaders `
                -AppId $AppId `
                -TimeoutSec $TimeoutSec `
                -BatchSize $PostBatchSize `
                -LogPath $LogPath

            if (-not $addResult.Success) {
                $errorMsg = '追加処理に失敗しました'
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
                $result.ErrorMessages += $errorMsg
                $result.ErrorMessages += $addResult.ErrorMessages
                $result.ErrorCount += $addResult.ErrorCount
                $result.RollbackTargets += @{
                    Operation = 'POST'
                    Data      = $addResult.ProcessedData
                }
                return $result
            }

            $result.AddedCount = $addResult.ProcessedCount
        }

        if ($comparisonResult.ToUpdate.Count -gt 0) {
            $updateResult = Invoke-UpdateData `
                -DataToUpdate $comparisonResult.ToUpdate `
                -ApiUri $ApiUri `
                -ApiHeaders $ApiHeaders `
                -AppId $AppId `
                -TimeoutSec $TimeoutSec `
                -BatchSize $PutBatchSize `
                -LogPath $LogPath

            if (-not $updateResult.Success) {
                $errorMsg = '更新処理に失敗しました'
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
                $result.ErrorMessages += $errorMsg
                $result.ErrorMessages += $updateResult.ErrorMessages
                $result.ErrorCount += $updateResult.ErrorCount
                $result.RollbackTargets += @{
                    Operation = 'PUT'
                    Data      = $updateResult.ProcessedData
                }
                if ($result.AddedCount -gt 0) {
                    $result.RollbackTargets += @{
                        Operation = 'POST'
                        Data      = $comparisonResult.ToAdd
                    }
                }
                return $result
            }

            $result.UpdatedCount = $updateResult.ProcessedCount
        }

        $result.Success = $true
        return $result
    }
    catch {
        $errorMsg = "データ同期処理でエラーが発生しました: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
        $result.ErrorMessages += $errorMsg
        $result.ErrorCount++
        return $result
    }
}

function Get-TargetDataFromApi {
    <#
    .SYNOPSIS
    APIから更新先データを取得します（500件単位でページング）。
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
        [int]$BatchSize = 500,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $allRecords = @()
    $offset = 0

    try {
        do {
            $query = if ($offset -gt 0) {
                "limit $BatchSize offset $offset"
            }
            else {
                "limit $BatchSize"
            }

            $encodedApp = [System.Uri]::EscapeDataString([string]$AppId)
            $encodedQuery = [System.Uri]::EscapeDataString($query)
            $requestUri = "{0}?app={1}&query={2}&totalCount=true" -f $ApiUri, $encodedApp, $encodedQuery

            $response = Get-ApiData `
                -Uri $requestUri `
                -Method 'GET' `
                -Headers $ApiHeaders `
                -TimeoutSec $TimeoutSec `
                -Verbose

            if ($null -eq $response) {
                break
            }

            if ($response -is [System.Array]) {
                if ($response.Count -gt 0) {
                    $allRecords += $response
                }
                break
            }

            if (-not $response.PSObject.Properties['records']) {
                Write-Log -Message 'レスポンスにrecordsが含まれていません。' -LogPath $LogPath -LogLevel 'WARNING'
                break
            }

            $records = @($response.records)
            if ($records.Count -eq 0) {
                break
            }

            $allRecords += $records
            if ($records.Count -lt $BatchSize) {
                break
            }

            $offset += $BatchSize
        } while ($true)

        return @($allRecords)
    }
    catch {
        $errorMsg = "更新先データ取得エラー: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
        throw $errorMsg
    }
}

function Invoke-AddData {
    <#
    .SYNOPSIS
    追加処理を実行します（API POST、100件単位）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$DataToAdd,

        [Parameter(Mandatory = $true)]
        [string]$ApiUri,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApiHeaders,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 100,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $result = [PSCustomObject]@{
        Success        = $false
        ProcessedCount = 0
        ErrorCount     = 0
        ErrorMessages  = @()
        ProcessedData  = @()
    }

    try {
        $payloadRecords = @($DataToAdd | ForEach-Object { ConvertTo-AddRecordPayload -SourceItem $_ })
        $totalBatches = [Math]::Ceiling($payloadRecords.Count / $BatchSize)

        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $BatchSize
            $endIndex = [Math]::Min($startIndex + $BatchSize - 1, $payloadRecords.Count - 1)
            $batchData = @($payloadRecords[$startIndex..$endIndex])
            $batchNumber = $batchIndex + 1

            try {
                $jsonData = @{
                    app     = $AppId
                    records = $batchData
                } | ConvertTo-Json -Depth 20

                Send-ApiRequest `
                    -Uri $ApiUri `
                    -Method 'POST' `
                    -Body $jsonData `
                    -Headers $ApiHeaders `
                    -TimeoutSec $TimeoutSec `
                    -Verbose | Out-Null

                $result.ProcessedCount += $batchData.Count
                $result.ProcessedData += @($DataToAdd[$startIndex..$endIndex])
            }
            catch {
                $errorMsg = "追加処理 バッチ $batchNumber エラー: $_"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
                $result.ErrorMessages += $errorMsg
                $result.ErrorCount++
                throw $errorMsg
            }
        }

        $result.Success = $true
        return $result
    }
    catch {
        $result.Success = $false
        return $result
    }
}

function Invoke-UpdateData {
    <#
    .SYNOPSIS
    更新処理を実行します。通常更新とsub_schedule専用更新を同一バッチで扱います。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$DataToUpdate,

        [Parameter(Mandatory = $true)]
        [string]$ApiUri,

        [Parameter(Mandatory = $true)]
        [hashtable]$ApiHeaders,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 100,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $result = [PSCustomObject]@{
        Success        = $false
        ProcessedCount = 0
        ErrorCount     = 0
        ErrorMessages  = @()
        ProcessedData  = @()
    }

    try {
        $updateRecords = @()
        foreach ($item in $DataToUpdate) {
            $updateRecords += ConvertTo-UpdateRecordPayload -UpdateItem $item
        }

        if ($updateRecords.Count -eq 0) {
            Write-Log -Message '更新対象データがありません' -LogPath $LogPath -LogLevel 'WARNING'
            $result.Success = $true
            return $result
        }

        $totalBatches = [Math]::Ceiling($updateRecords.Count / $BatchSize)
        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $BatchSize
            $endIndex = [Math]::Min($startIndex + $BatchSize - 1, $updateRecords.Count - 1)
            $batchData = @($updateRecords[$startIndex..$endIndex])
            $batchNumber = $batchIndex + 1

            try {
                $jsonData = @{
                    app     = $AppId
                    records = $batchData
                } | ConvertTo-Json -Depth 20

                Send-ApiRequest `
                    -Uri $ApiUri `
                    -Method 'PUT' `
                    -Body $jsonData `
                    -Headers $ApiHeaders `
                    -TimeoutSec $TimeoutSec `
                    -Verbose | Out-Null

                $result.ProcessedCount += $batchData.Count
                $result.ProcessedData += @($DataToUpdate[$startIndex..$endIndex])
            }
            catch {
                $errorMsg = "更新処理 バッチ $batchNumber エラー: $_"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel 'ERROR'
                $result.ErrorMessages += $errorMsg
                $result.ErrorCount++
                throw $errorMsg
            }
        }

        $result.Success = $true
        return $result
    }
    catch {
        $result.Success = $false
        return $result
    }
}

Export-ModuleMember -Function `
    Sync-DataWithApi, `
    Get-TargetDataFromApi, `
    Compare-DataByLineLotNumber, `
    Invoke-AddData, `
    Invoke-UpdateData