# JsonConverter.psm1
# データをJSON形式に変換するモジュール

function Get-ScheduleDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )

    if ($InputObject.sub_schedule -and $InputObject.sub_schedule.Count -gt 0) {
        return $InputObject.sub_schedule[0].sub_schedule_date
    }

    return $InputObject.standard_date
}

function ConvertTo-JsonData {
    <#
    .SYNOPSIS
    データをJSON形式に変換します。
    
    .DESCRIPTION
    PowerShellのConvertTo-Jsonコマンドレットを使用してデータをJSON形式に変換します。
    各プロパティを{value: ...}形式でラップし、records配列でラップしたJSONオブジェクトを生成します。

    .PARAMETER AppId
    キントーンアプリID
    
    .PARAMETER InputObject
    変換するオブジェクト（配列）
    
    .PARAMETER Depth
    JSONの深さ（デフォルト: 10）
    
    .PARAMETER Compress
    圧縮形式で出力するか（デフォルト: $false）
    
    .PARAMETER WrapWithValue
    各プロパティを{value: ...}形式でラップするか（デフォルト: $true）
    
    .EXAMPLE
    $json = ConvertTo-JsonData -InputObject $data
    
    .EXAMPLE
    $json = ConvertTo-JsonData -InputObject $data -Depth 20 -Compress
    
    .EXAMPLE
    $json = ConvertTo-JsonData -InputObject $data -WrapWithValue $false
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$AppId,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$InputObject,
        
        [Parameter(Mandatory = $false)]
        [int]$Depth = 10,
        
        [Parameter(Mandatory = $false)]
        [switch]$Compress,
        
        [Parameter(Mandatory = $false)]
        [bool]$WrapWithValue = $true
    )
    
    try {
        # 配列でない場合は配列に変換
        $dataArray = if ($InputObject -is [System.Array]) { $InputObject } else { @($InputObject) }
        
        # WrapWithValueがfalseの場合は、そのままJSON変換
        if (-not $WrapWithValue) {
            $result = @{
                records = $dataArray
            }
            
            $params = @{
                InputObject = $result
                Depth       = $Depth
                Compress    = $Compress
            }
            
            $json = ConvertTo-Json @params -ErrorAction Stop
            Write-Verbose "JSON変換完了（サイズ: $($json.Length) 文字）"
            return $json
        }
        
        # WrapWithValueがtrueの場合は、各プロパティを{value: ...}形式でラップ
        $wrappedRecords = @()
        foreach ($item in $dataArray) {
            $wrappedItem = ConvertTo-WrappedJsonObject -InputObject $item
            # schedule_dateを追加（sub_scheduleの最初の日付、またはstandard_dateを使用）
            $scheduleDate = Get-ScheduleDate -InputObject $item
            
            # schedule_dateを最初に追加
            $wrappedHash = @{
                schedule_date = [PSCustomObject]@{value = $scheduleDate }
            }
            
            # 既存のプロパティを順序付きで追加
            $propertyOrder = @(
                "line_name", "index", "model_name", "lot_number", "board_name",
                "tact", "utilization_rate", "hour_production_volume", "change_time",
                "board_division", "input_quantity", "standard_date", "model_code",
                "lot_volume", "sub_schedule", "line_lot_number"
            )
            
            foreach ($propName in $propertyOrder) {
                if ($wrappedItem.PSObject.Properties[$propName]) {
                    $wrappedHash[$propName] = $wrappedItem.PSObject.Properties[$propName].Value
                }
            }
            
            # 残りのプロパティを追加（tanabanなど）
            foreach ($property in $wrappedItem.PSObject.Properties) {
                if (-not $wrappedHash.ContainsKey($property.Name) -and $propertyOrder -notcontains $property.Name) {
                    $wrappedHash[$property.Name] = $property.Value
                }
            }
            
            $wrappedRecords += [PSCustomObject]$wrappedHash
        }
        
        $result = @{
            app       = $AppId
            records   = $wrappedRecords
        }
        
        $params = @{
            InputObject = $result
            Depth       = $Depth
            Compress    = $Compress
        }
        
        $json = ConvertTo-Json @params -ErrorAction Stop
        
        Write-Verbose "JSON変換完了（サイズ: $($json.Length) 文字、レコード数: $($wrappedRecords.Count)）"
        return $json
    }
    catch {
        throw "JSON変換エラー: $_"
    }
}

function ConvertTo-WrappedJsonObject {
    <#
    .SYNOPSIS
    オブジェクトの各プロパティを{value: ...}形式でラップします。
    
    .DESCRIPTION
    オブジェクトの各プロパティを{value: ...}形式でラップした新しいオブジェクトを生成します。
    配列の場合は各要素に対して再帰的に処理します。
    
    .PARAMETER InputObject
    変換するオブジェクト
    
    .EXAMPLE
    $wrapped = ConvertTo-WrappedJsonObject -InputObject $data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )
    
    # nullの場合はそのまま返す
    if ($null -eq $InputObject) {
        return $null
    }
    
    # 配列の場合は各要素に対して再帰的に処理
    if ($InputObject -is [System.Array] -or $InputObject -is [System.Collections.IEnumerable]) {
        $result = @()
        foreach ($item in $InputObject) {
            $result += ConvertTo-WrappedJsonObject -InputObject $item
        }
        return $result
    }
    
    # PSCustomObjectの場合は各プロパティをラップ
    if ($InputObject -is [PSCustomObject]) {
        $wrappedHash = @{}
        
        foreach ($property in $InputObject.PSObject.Properties) {
            $propertyName = $property.Name
            $propertyValue = $property.Value
            
            # sub_scheduleの場合は特別処理（配列内の各要素をラップ）
            if ($propertyName -eq "sub_schedule" -and $propertyValue -is [System.Array]) {
                $wrappedSubSchedule = @()
                foreach ($subItem in $propertyValue) {
                    if ($subItem -is [PSCustomObject]) {
                        $wrappedSubItem = [PSCustomObject]@{
                            value = [PSCustomObject]@{
                                sub_schedule_date = [PSCustomObject]@{value = $subItem.sub_schedule_date }
                                sub_lot_volume    = [PSCustomObject]@{value = $subItem.sub_lot_volume }
                                sub_index         = [PSCustomObject]@{value = $subItem.sub_index }
                            }
                        }
                        $wrappedSubSchedule += $wrappedSubItem
                    }
                    else {
                        $wrappedSubSchedule += $subItem
                    }
                }
                $wrappedHash[$propertyName] = [PSCustomObject]@{value = $wrappedSubSchedule }
            }
            # その他のプロパティは{value: ...}形式でラップ
            else {
                # プロパティの値が配列の場合は再帰的に処理
                if ($propertyValue -is [System.Array]) {
                    $wrappedArray = @()
                    foreach ($item in $propertyValue) {
                        $wrappedArray += ConvertTo-WrappedJsonObject -InputObject $item
                    }
                    $wrappedHash[$propertyName] = [PSCustomObject]@{value = $wrappedArray }
                }
                else {
                    $wrappedHash[$propertyName] = [PSCustomObject]@{value = $propertyValue }
                }
            }
        }
        
        return [PSCustomObject]$wrappedHash
    }
    
    # その他の型の場合はそのまま返す
    return $InputObject
}

# モジュールをエクスポート
Export-ModuleMember -Function Get-ScheduleDate, ConvertTo-JsonData, ConvertTo-WrappedJsonObject

