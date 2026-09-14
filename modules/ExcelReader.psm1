# ExcelReader.psm1
# Excelファイルを読み取るモジュール（テンプレート）

function ConvertTo-SyncLotNumber {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $normalized = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    return $normalized
}

function Get-SubScheduleChangeTime {
    <#
    .SYNOPSIS
    日別予定へ登録する切替時間を決定します。

    .DESCRIPTION
    同一Excel行で生産台数セルが左隣の列から連続している場合は、同じ生産指図の
    日跨ぎ継続とみなし、右側の予定には切替時間0を登録します。
    連続区間の先頭、または前回予定セルとの間に空列がある場合はAP列の切替時間を使用します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$ChangeTime,

        [Parameter(Mandatory = $true)]
        [int]$CurrentColumn,

        [Parameter(Mandatory = $false)]
        [Nullable[int]]$PreviousScheduledColumn
    )

    if ($null -ne $PreviousScheduledColumn -and $CurrentColumn -eq ([int]$PreviousScheduledColumn + 1)) {
        return 0
    }

    return $ChangeTime
}

function Get-ExcelScheduleDateRange {
    <#
    .SYNOPSIS
    Excelの予定日ヘッダー全体から同期対象の日付範囲を取得します。

    .DESCRIPTION
    各ページのI列～AE列にある日付ヘッダーを読み取り、最小日付～最大日付を返します。
    予定台数が入っていない日も範囲に含めることで、Excelから予定が消えた日の既存予定を
    kintone側で無効化できるようにします。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$PageCount
    )

    $dates = @()
    for ($page = 1; $page -le $PageCount; $page++) {
        $scheduleDateRow = 8 + (($page - 1) * 62)
        for ($column = 9; $column -le 31; $column++) {
            $rawDate = $Worksheet.Cells[$scheduleDateRow, $column].value
            if ($null -eq $rawDate -or [string]::IsNullOrWhiteSpace([string]$rawDate)) {
                continue
            }

            $convertedDate = ConvertTo-DateString -DateValue $rawDate
            $parsedDate = [DateTime]::MinValue
            $isValidDate = [DateTime]::TryParseExact(
                [string]$convertedDate,
                'yyyy-MM-dd',
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsedDate
            )
            if (-not $isValidDate) {
                throw "予定日ヘッダーの日付形式が不正です: page=$page, column=$column, value=$rawDate"
            }

            $dates += $parsedDate.Date
        }
    }

    if ($dates.Count -eq 0) {
        throw 'Excel予定日ヘッダーから有効な日付を取得できません。'
    }

    $sortedDates = @($dates | Sort-Object)
    return [PSCustomObject]@{
        Start = $sortedDates[0].ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        End   = $sortedDates[$sortedDates.Count - 1].ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
}

function Read-ExcelData {
    <#
    .SYNOPSIS
    Excelファイルを読み取ります。

    .DESCRIPTION
    ImportExcelモジュールを使用してExcelファイルを読み込みます。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlsxPath,

        [Parameter(Mandatory = $false)]
        [string]$WorksheetName
    )

    if (-not (Test-Path $XlsxPath)) {
        throw "ファイルが存在しません: $XlsxPath"
    }

    if (-not (Get-Module ImportExcel)) {
        Import-Module ImportExcel -ErrorAction Stop
    }

    if (-not (Get-Module Utils)) {
        $scriptDirectory = Split-Path -Parent $PSScriptRoot
        $utilsPath = Join-Path $scriptDirectory "modules\Utils.psm1"
        if (Test-Path $utilsPath) {
            Import-Module $utilsPath -Force
        }
        else {
            Write-Warning "Utilsモジュールが見つかりません。日付変換機能が使用できない可能性があります。"
        }
    }

    Write-Verbose "Excelファイルを読み込み中: $XlsxPath"

    try {
        $fileName = Split-Path -Leaf $XlsxPath
        $fileName = $fileName -replace '\.xlsx', ''

        if ($WorksheetName) {
            $excelPackage = Open-ExcelPackage -Path $XlsxPath -WorksheetName $WorksheetName
        }
        else {
            $excelPackage = Open-ExcelPackage -Path $XlsxPath
        }

        $worksheet = $excelPackage.Workbook.Worksheets[1]
        $pageCount = $worksheet.Cells["A53"].value
        $scheduleDateRange = Get-ExcelScheduleDateRange -Worksheet $worksheet -PageCount $pageCount

        $index = 1
        $startRow = 10
        $endRow = 53
        $data = @()

        for ($page = 1; $page -le $pageCount; $page++) {
            $scheduleDateRow = 8 + (($page - 1) * 62)
            for ($row = $startRow; $row -le $endRow; $row += 2) {
                $rowDown = $row + 1
                $modelName = $worksheet.Cells["B$row"].value
                $lotNumber = $worksheet.Cells["D$row"].value

                if ([string]::IsNullOrEmpty($modelName)) {
                    continue
                }

                $lotNumber = ConvertTo-SyncLotNumber -Value $lotNumber
                if ($null -eq $lotNumber) {
                    Write-Warning "lot_numberが空のため同期対象から除外します: file=$fileName, page=$page, row=$row"
                    continue
                }

                $lineName = $fileName
                $boardName = $worksheet.Cells["B$rowDown"].value
                $modelCode = $worksheet.Cells["D$rowDown"].value
                $standardDateValue = $worksheet.Cells["F$row"].value
                $standardDate = ConvertTo-DateString -DateValue $standardDateValue
                $lotVolume = $worksheet.Cells["G$rowDown"].value
                $tact = $worksheet.Cells["AK$row"].value
                $utilizationRate = $worksheet.Cells["AL$row"].value
                $hourProductionVolume = $worksheet.Cells["AL$rowDown"].value
                $changeTime = $worksheet.Cells["AP$row"].value
                $boardDivision = $worksheet.Cells["AQ$row"].value
                $inputQuantity = $worksheet.Cells["AQ$rowDown"].value
                $tanaban = $worksheet.Cells["H$row"].value

                $subSchedule = @()
                $previousScheduledColumn = $null
                for ($column = 9; $column -le 31; $column++) {
                    $dateValue = $worksheet.Cells[$scheduleDateRow, $column].value
                    $convertedDate = ConvertTo-DateString -DateValue $dateValue
                    $columnValue = $worksheet.Cells[$row, $column].value

                    if (-not [string]::IsNullOrEmpty($columnValue)) {
                        $dailyChangeTime = Get-SubScheduleChangeTime `
                            -ChangeTime $changeTime `
                            -CurrentColumn $column `
                            -PreviousScheduledColumn $previousScheduledColumn

                        $subSchedule += [PSCustomObject]@{
                            sub_schedule_date = $convertedDate
                            sub_lot_volume    = $columnValue
                            sub_index         = $index
                            sub_change_time   = $dailyChangeTime
                            有効判定          = 'True'
                        }

                        $previousScheduledColumn = $column
                    }
                }

                $existingRecords = @($data | Where-Object { $_.lot_number -eq $lotNumber })
                if ($existingRecords.Count -gt 0) {
                    foreach ($existingRecord in $existingRecords) {
                        $existingRecord.sub_schedule += $subSchedule
                    }
                }
                else {
                    $rowData = [PSCustomObject]@{
                        line_lot_number        = $lineName + $lotNumber
                        lot_number             = $lotNumber
                        line_name              = $lineName
                        index                  = $index
                        model_name             = $modelName
                        board_name             = $boardName
                        model_code             = $modelCode
                        standard_date          = $standardDate
                        lot_volume             = $lotVolume
                        tact                   = $tact
                        utilization_rate       = $utilizationRate
                        hour_production_volume = $hourProductionVolume
                        change_time            = $changeTime
                        board_division         = $boardDivision
                        input_quantity         = $inputQuantity
                        tanaban                = $tanaban
                        sub_schedule           = $subSchedule
                        sync_date_range_start  = $scheduleDateRange.Start
                        sync_date_range_end    = $scheduleDateRange.End
                    }

                    $data += $rowData
                }

                $index++
            }
            $startRow = $startRow + 62
            $endRow = 53 + $page * 62
        }

        Write-Verbose "Excelデータの読み込み完了（件数: $($data.Count), 日付範囲: $($scheduleDateRange.Start) ～ $($scheduleDateRange.End)）"
        return $data
    }
    catch {
        throw "Excelデータ読み取りエラー: $_"
    }
}

Export-ModuleMember -Function Read-ExcelData