# ExcelReader.psm1
# Excelファイルを読み取るモジュール（テンプレート）

function Read-ExcelData {
    <#
    .SYNOPSIS
    Excelファイルを読み取ります。
    
    .DESCRIPTION
    ImportExcelモジュールを使用してExcelファイルを読み込みます。
    この関数はテンプレートです。実際の読み取り処理は実装してください。
    
    .PARAMETER XlsxPath
    読み取る.xlsxファイルのパス
    
    .PARAMETER WorksheetName
    読み取るワークシート名（省略時は最初のシート）
    
    .EXAMPLE
    $data = Read-ExcelData -XlsxPath "C:\data\file.xlsx"
    
    .EXAMPLE
    $data = Read-ExcelData -XlsxPath "C:\data\file.xlsx" -WorksheetName "Sheet1"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlsxPath,
        
        [Parameter(Mandatory = $false)]
        [string]$WorksheetName
    )
    
    # パス検証
    if (-not (Test-Path $XlsxPath)) {
        throw "ファイルが存在しません: $XlsxPath"
    }
    
    # ImportExcelモジュールの確認
    if (-not (Get-Module ImportExcel)) {
        Import-Module ImportExcel -ErrorAction Stop
    }
    
    # Utilsモジュールの確認（日付変換関数を使用するため）
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

        # XlsxPathから拡張子を抜いたファイル名を取得
        $fileName = Split-Path -Leaf $XlsxPath
        $fileName = $fileName -replace '\.xlsx', ''

        # Excelパッケージを開く
        if ($WorksheetName) {
            $excelPackage = Open-ExcelPackage -Path $XlsxPath -WorksheetName $WorksheetName
        }
        else {
            $excelPackage = Open-ExcelPackage -Path $XlsxPath
        }

        $pageCount = $excelPackage.Workbook.Worksheets[1].Cells["A53"].value

        $index = 1
        $startRow = 10
        $endRow = 53
        $data = @()  # データ配列を初期化

        for ($page = 1; $page -le $pageCount; $page++) {
            for ($row = $startRow; $row -le $endRow; $row += 2) {
                # 2行ごとにデータを取得
                $rowDown = $row + 1
                $modelName = $excelPackage.Workbook.Worksheets[1].Cells["B$row"].value
                $lotNumber = $excelPackage.Workbook.Worksheets[1].Cells["D$row"].value

                # モデル名が空白の場合はスキップ
                if ([string]::IsNullOrEmpty($modelName)) {
                    continue
                }

                $line_name = $fileName
                $boradName = $excelPackage.Workbook.Worksheets[1].Cells["B$rowDown"].value
                $modelCode = $excelPackage.Workbook.Worksheets[1].Cells["D$rowDown"].value
                # 日付を文字列として取得
                $standardDateValue = $excelPackage.Workbook.Worksheets[1].Cells["F$row"].value
                $standardDate = ConvertTo-DateString -DateValue $standardDateValue
                $lotVolume = $excelPackage.Workbook.Worksheets[1].Cells["G$rowDown"].value
                $tact = $excelPackage.Workbook.Worksheets[1].Cells["AK$row"].value
                $utilizationRate = $excelPackage.Workbook.Worksheets[1].Cells["AL$row"].value
                $hourProductionVolume = $excelPackage.Workbook.Worksheets[1].Cells["AL$rowDown"].value
                $changeTime = $excelPackage.Workbook.Worksheets[1].Cells["AP$row"].value
                $boardDivision = $excelPackage.Workbook.Worksheets[1].Cells["AQ$row"].value
                $inputQuantity = $excelPackage.Workbook.Worksheets[1].Cells["AQ$rowDown"].value
                $tanaban = $excelPackage.Workbook.Worksheets[1].Cells["H$row"].value
                
                $subSchedule = @()
                # I列からAE列をループ処理
                for ($column = 9; $column -le 31; $column++) {
                    
                    # 列番号を文字列に変換
                    $dateStr = $excelPackage.Workbook.Worksheets[1].Cells[8, $column].value
                    $convertedDate = ConvertTo-DateString -DateValue $dateStr
                    $columnValue = $excelPackage.Workbook.Worksheets[1].Cells[$row, $column].value
                    
                    # 値が空白でない場合はサブスケジュールに追加
                    if (-not [string]::IsNullOrEmpty($columnValue)) {
                        $subSchedule += [PSCustomObject]@{
                            sub_schedule_date = $convertedDate
                            sub_lot_volume    = $columnValue
                            sub_index         = $index
                        }
                    }
                }

                # lot_numberが既に存在する場合は、sub_scheduleを既存のものに追加
                if ($data | Where-Object { $_.lot_number -eq $lotNumber }) {
                    $data | Where-Object { $_.lot_number -eq $lotNumber } | ForEach-Object {
                        $_.sub_schedule += $subSchedule
                    }
                }
                else {
                    # データをオブジェクトとして作成（単純な構造で返す）
                    $rowData = [PSCustomObject]@{
                        line_lot_number        = $line_name + $lotNumber
                        lot_number             = $lotNumber
                        line_name              = $line_name
                        index                  = $index
                        model_name             = $modelName
                        board_name             = $boradName
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
                    }
                
                    # データ配列に追加
                    $data += $rowData
            
                }
                # インデックスを更新
                $index++
            }
            $startRow = $startRow + 62
            $endRow = 53 + $page * 62
        }
        
        Write-Verbose "Excelデータの読み込み完了（件数: $($data.Count)）"

        return $data
    }
    catch {
        throw "A12セルの値取得エラー: $_"
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Read-ExcelData

