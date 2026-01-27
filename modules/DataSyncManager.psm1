# DataSyncManager.psm1
# データ同期処理を行うモジュール（追加・更新・削除）

function Sync-DataWithApi {
    <#
    .SYNOPSIS
    更新元データと更新先データを同期します。
    
    .DESCRIPTION
    フローチャートに基づいて、更新元データと更新先データを突合し、
    追加・更新・削除処理を順次実行します。
    
    .PARAMETER SourceData
    更新元データ（Excelから読み込んだデータ配列）
    
    .PARAMETER ApiUri
    APIエンドポイントURL
    
    .PARAMETER ApiHeaders
    APIリクエストヘッダー（ハッシュテーブル）
    
    .PARAMETER AppId
    キントーンアプリID
    
    .PARAMETER TimeoutSec
    タイムアウト秒数（デフォルト: 30）
    
    .PARAMETER LogPath
    ログファイルのパス
    
    .PARAMETER GetBatchSize
    GETリクエストのバッチサイズ（デフォルト: 500）
    
    .PARAMETER PostBatchSize
    POSTリクエストのバッチサイズ（デフォルト: 100）
    
    .PARAMETER PutBatchSize
    PUTリクエストのバッチサイズ（デフォルト: 100）
    
    .PARAMETER DeleteBatchSize
    DELETEリクエストのバッチサイズ（デフォルト: 100）
    
    .EXAMPLE
    $result = Sync-DataWithApi -SourceData $excelData -ApiUri $Config.Api.Uri -ApiHeaders $Config.Api.Headers -AppId $Config.AppId -LogPath $logPath
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
        [int]$PutBatchSize = 100,
        
        [Parameter(Mandatory = $false)]
        [int]$DeleteBatchSize = 100
    )
    
    # 処理結果を保持するオブジェクト
    $result = [PSCustomObject]@{
        Success         = $false
        AddedCount      = 0
        UpdatedCount    = 0
        DeletedCount    = 0
        ErrorCount      = 0
        ErrorMessages   = @()
        RollbackTargets = @()
    }
    
    try {
        # ステップ1: 更新元データ取得
        if ($null -eq $SourceData -or $SourceData.Count -eq 0) {
            $errorMsg = "更新元データが空です"
            Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
            $result.ErrorMessages += $errorMsg
            $result.ErrorCount++
            return $result
        }
        
        # ステップ2: 更新先データ取得（API GET、500件単位）
        try {
            $targetData = Get-TargetDataFromApi -ApiUri $ApiUri -ApiHeaders $ApiHeaders -AppId $AppId -TimeoutSec $TimeoutSec -BatchSize $GetBatchSize -LogPath $LogPath
            
            # nullの場合は空の配列として扱う（0件として処理）
            if ($null -eq $targetData) {
                Write-Log -Message "更新先データがnullでした。空の配列として扱います（0件）" -LogPath $LogPath -LogLevel "INFO"
                $targetData = @()
            }
        }
        catch {
            $errorMsg = "更新先データの取得に失敗しました: $_"
            Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
            Write-Host "エラー: $errorMsg"
            $result.ErrorMessages += $errorMsg
            $result.ErrorCount++
            return $result
        }
        
        # targetDataがレスポンスオブジェクトの場合、records配列を取得
        $targetRecords = $targetData
        if ($targetData.PSObject.Properties['records']) {
            $targetRecords = $targetData.records
        }
        
        # ステップ3: ライン_ロットナンバーで突合
        $comparisonResult = Compare-DataByLineLotNumber -SourceData $SourceData -TargetData $targetRecords -LogPath $LogPath
        
        # ステップ4: 追加処理（API POST、100件単位）
        if ($comparisonResult.ToAdd.Count -gt 0) {
            $addResult = Invoke-AddData -DataToAdd $comparisonResult.ToAdd -ApiUri $ApiUri -ApiHeaders $ApiHeaders -AppId $AppId -TimeoutSec $TimeoutSec -BatchSize $PostBatchSize -LogPath $LogPath
            
            if (-not $addResult.Success) {
                $errorMsg = "追加処理に失敗しました"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
                $result.ErrorMessages += $errorMsg
                $result.ErrorMessages += $addResult.ErrorMessages
                $result.ErrorCount += $addResult.ErrorCount
                $result.RollbackTargets += @{
                    Operation = "POST"
                    Data      = $addResult.ProcessedData
                }
                return $result
            }
            
            $result.AddedCount = $addResult.ProcessedCount
            Write-Log -Message "追加処理完了: $($result.AddedCount) 件" -LogPath $LogPath -LogLevel "INFO"
        }
        
        # ステップ5: 更新処理（API PUT、100件単位・全項目）
        if ($comparisonResult.ToUpdate.Count -gt 0) {
            $updateResult = Invoke-UpdateData -DataToUpdate $comparisonResult.ToUpdate -ApiUri $ApiUri -ApiHeaders $ApiHeaders -AppId $AppId -TimeoutSec $TimeoutSec -BatchSize $PutBatchSize -LogPath $LogPath
            if (-not $updateResult.Success) {
                $errorMsg = "更新処理に失敗しました"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
                $result.ErrorMessages += $errorMsg
                $result.ErrorMessages += $updateResult.ErrorMessages
                $result.ErrorCount += $updateResult.ErrorCount
                $result.RollbackTargets += @{
                    Operation = "PUT"
                    Data      = $updateResult.ProcessedData
                }
                # 追加処理もロールバック対象に追加
                if ($result.AddedCount -gt 0) {
                    $result.RollbackTargets += @{
                        Operation = "POST"
                        Data      = $comparisonResult.ToAdd
                    }
                }
                return $result
            }
            
            $result.UpdatedCount = $updateResult.ProcessedCount
        }
        
        # ステップ6: 削除処理（API DELETE、100件単位）
        if ($comparisonResult.ToDelete.Count -gt 0) {
            Write-Log -Message "削除処理開始: $($comparisonResult.ToDelete.Count) 件" -LogPath $LogPath -LogLevel "INFO"
            $deleteResult = Invoke-DeleteData -DataToDelete $comparisonResult.ToDelete -ApiUri $ApiUri -ApiHeaders $ApiHeaders -AppId $AppId -TimeoutSec $TimeoutSec -BatchSize $DeleteBatchSize -LogPath $LogPath
            
            if (-not $deleteResult.Success) {
                $errorMsg = "削除処理に失敗しました"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
                $result.ErrorMessages += $errorMsg
                $result.ErrorMessages += $deleteResult.ErrorMessages
                $result.ErrorCount += $deleteResult.ErrorCount
                $result.RollbackTargets += @{
                    Operation = "DELETE"
                    Data      = $deleteResult.ProcessedData
                }
                # 追加・更新処理もロールバック対象に追加
                if ($result.AddedCount -gt 0) {
                    $result.RollbackTargets += @{
                        Operation = "POST"
                        Data      = $comparisonResult.ToAdd
                    }
                }
                if ($result.UpdatedCount -gt 0) {
                    $result.RollbackTargets += @{
                        Operation = "PUT"
                        Data      = $comparisonResult.ToUpdate
                    }
                }
                return $result
            }
            
            $result.DeletedCount = $deleteResult.ProcessedCount
            Write-Log -Message "削除処理完了: $($result.DeletedCount) 件" -LogPath $LogPath -LogLevel "INFO"
        }
        
        # 全処理成功
        $result.Success = $true
        
        return $result
    }
    catch {
        $errorMsg = "データ同期処理でエラーが発生しました: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
        $result.ErrorMessages += $errorMsg
        $result.ErrorCount++
        
        # ロールバック対象を記録
        if ($result.AddedCount -gt 0 -or $result.UpdatedCount -gt 0 -or $result.DeletedCount -gt 0) {
            Write-Log -Message "ロールバック対象を記録しました（成功分も含めて全体失敗扱い）" -LogPath $LogPath -LogLevel "ERROR"
        }
        
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
            # キントーンAPIのGETリクエスト形式
            $bodyObject = @{
                app        = $AppId
                totalCount = $true
            }
            
            # オフセットとリミットを設定
            if ($offset -gt 0) {
                $query = "limit $BatchSize offset $offset"
            }
            else {
                $query = "limit $BatchSize"
            }

            $bodyObject.query = $query
            
            $bodyJson = $bodyObject | ConvertTo-Json -Depth 10
            
            Write-Verbose "API GET リクエスト送信（offset: $offset, limit: $BatchSize）"
            
            $response = Get-ApiData -Uri $ApiUri -Method "GET" -Body $bodyJson -Headers $ApiHeaders -TimeoutSec $TimeoutSec -Verbose

            if ($null -ne $response) {
                
                # プロパティ名を確認（ログ出力なし）
                # レスポンスが配列の場合（直接records配列が返される場合）
                if ($response -is [System.Array]) {
                    if ($response.Count -gt 0) {
                        $allRecords += $response
                        $currentCount = $response.Count
                        break
                    }
                    else {
                        Write-Log -Message "レスポンスが空の配列です" -LogPath $LogPath -LogLevel "INFO"
                        break
                    }
                }
                # レスポンスがオブジェクトでrecordsプロパティがある場合
                elseif ($response.PSObject.Properties['records']) {
                    $records = $response.records
                    # レスポンスが空でないかつrecordsが存在する場合
                    if ($records -and $records.Count -gt 0) {
                        $allRecords += $records
                        $currentCount = $records.Count
                        
                        # 次のページがあるかチェック
                        if ($currentCount -lt $BatchSize) {
                            break
                        }
                        
                        $offset += $BatchSize
                    }
                    else {
                        Write-Log -Message "recordsが空です" -LogPath $LogPath -LogLevel "INFO"
                        break
                    }
                }
                else {
                    try {
                        $responseJson = $response | ConvertTo-Json -Depth 5
                        Write-Log -Message "レスポンスにrecordsが含まれていません。レスポンス内容: $responseJson" -LogPath $LogPath -LogLevel "WARNING"
                    }
                    catch {
                        Write-Log -Message "レスポンスにrecordsが含まれていません。JSON変換に失敗: $_" -LogPath $LogPath -LogLevel "WARNING"
                    }
                    break
                }
            }
            else {
                Write-Log -Message "レスポンスがnullです" -LogPath $LogPath -LogLevel "WARNING"
                break
            }
        } while ($true)
        
        # 空の配列でも正常に返す（データが存在しないだけ）
        if ($allRecords.Count -eq 0) {
            Write-Log -Message "取得したデータが0件です（これは正常な場合があります）" -LogPath $LogPath -LogLevel "INFO"
        }
        
        # 確実に配列を返す（nullの場合は空の配列を返す）
        # $allRecordsは初期化時に@()で設定されているので、nullになることはないはずだが、念のためチェック
        if ($null -eq $allRecords) {
            Write-Log -Message "警告: allRecordsがnullです。空の配列を返します。" -LogPath $LogPath -LogLevel "WARNING"
            Write-Host "DEBUG: allRecordsがnullです。空の配列を返します。"
            $allRecords = @()
        }
        
        # 返り値が確実に配列であることを確認
        if ($allRecords -isnot [System.Array]) {
            Write-Log -Message "警告: allRecordsが配列ではありません。型: $($allRecords.GetType().FullName)" -LogPath $LogPath -LogLevel "WARNING"
            Write-Host "DEBUG: allRecordsが配列ではありません。型: $($allRecords.GetType().FullName)"
            # 配列に変換
            $allRecords = @($allRecords)
        }
        
        Write-Host "DEBUG: Get-TargetDataFromApi終了 - 返り値: 型=$($allRecords.GetType().FullName), 件数=$($allRecords.Count)"
        return $allRecords
    }
    catch {
        $errorMsg = "更新先データ取得エラー: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
        throw $errorMsg
    }
}

function Compare-DataByLineLotNumber {
    <#
    .SYNOPSIS
    ライン_ロットナンバーでデータを突合し、追加・更新・削除対象を抽出します。
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
        # 更新元データのキー（line_name_lot_number）を作成
        $sourceKeys = @{}
        foreach ($item in $SourceData) {
            $key = "$($item.line_name)$($item.lot_number)"
            $sourceKeys[$key] = $item
        }
        
        # 更新先データのキーを作成（キントーンAPIのレスポンス形式を考慮）
        $targetKeys = @{}
        foreach ($item in $TargetData) {
            # キントーンAPIのレスポンスは {value: ...} 形式でラップされている
            $lineName = if ($item.line_name.value) { $item.line_name.value } else { $item.line_name }
            $lotNumber = if ($item.lot_number.value) { $item.lot_number.value } else { $item.lot_number }
            $key = "${lineName}${lotNumber}"
            $targetKeys[$key] = $item
        }
        
        # 追加対象（更新元のみ存在）
        $toAdd = @()
        foreach ($key in $sourceKeys.Keys) {
            if (-not $targetKeys.ContainsKey($key)) {
                $toAdd += $sourceKeys[$key]
            }
        }
        
        # 更新対象（両方に存在）
        $toUpdate = @()
        foreach ($key in $sourceKeys.Keys) {
            if ($targetKeys.ContainsKey($key)) {
                # 更新元データと更新先データの両方を含むオブジェクトを作成
                $sourceItem = $sourceKeys[$key]
                $targetItem = $targetKeys[$key]
                
                # デバッグ: データの存在確認
                if ($null -eq $sourceItem) {
                    Write-Log -Message "警告: sourceItemがnullです（キー: $key）" -LogPath $LogPath -LogLevel "WARNING"
                }
                if ($null -eq $targetItem) {
                    Write-Log -Message "警告: targetItemがnullです（キー: $key）" -LogPath $LogPath -LogLevel "WARNING"
                }
                
                $toUpdate += [PSCustomObject]@{
                    Source = $sourceItem
                    Target = $targetItem
                    Key    = $key
                }
            }
        }
        
        # 削除対象（更新先のみ存在）
        $toDelete = @()
        foreach ($key in $targetKeys.Keys) {
            if (-not $sourceKeys.ContainsKey($key)) {
                $toDelete += $targetKeys[$key]
            }
        }
        
        $returnItem = [PSCustomObject]@{
            ToAdd    = $toAdd
            ToUpdate = $toUpdate
            ToDelete = $toDelete
        }
        return $returnItem
    }
    catch {
        $errorMsg = "データ突合エラー: $_"
        Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
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
        $totalBatches = [Math]::Ceiling($DataToAdd.Count / $BatchSize)
        
        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $BatchSize
            $endIndex = [Math]::Min($startIndex + $BatchSize - 1, $DataToAdd.Count - 1)
            $batchData = $DataToAdd[$startIndex..$endIndex]
            $batchNumber = $batchIndex + 1
            
            Write-Log -Message "追加処理 バッチ $batchNumber / $totalBatches（$($batchData.Count) 件）" -LogPath $LogPath -LogLevel "INFO"
            
            try {
                # JSON変換
                $jsonData = ConvertTo-JsonData -AppId $AppId -InputObject $batchData -Depth 20 -Verbose
                
                # API POST送信
                Send-ApiRequest `
                    -Uri $ApiUri `
                    -Method "POST" `
                    -Body $jsonData `
                    -Headers $ApiHeaders `
                    -TimeoutSec $TimeoutSec `
                    -Verbose | Out-Null
                
                $result.ProcessedCount += $batchData.Count
                $result.ProcessedData += $batchData
                Write-Log -Message "追加処理 バッチ $batchNumber 成功（$($batchData.Count) 件）" -LogPath $LogPath -LogLevel "INFO"
            }
            catch {
                $errorMsg = "追加処理 バッチ $batchNumber エラー: $_"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
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
    更新処理を実行します（API PUT、100件単位・全項目）。
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
        Write-Host "host01"
        # 更新対象データを準備（キントーンAPIのPUT形式に変換）
        $updateRecords = @()
        foreach ($item in $DataToUpdate) {
            $sourceItem = $item.Source
            $targetItem = $item.Target
            
            # 更新先のレコードIDを取得（キントーンAPIのレスポンス形式を考慮）
            $recordId = $null
            if ($targetItem.PSObject.Properties['line_lot_number']) {
                $idProp = $targetItem.PSObject.Properties['line_lot_number']
                Write-Host $idProp
                $recordId = if ($idProp.Value.value) { $idProp.Value.value } else { $idProp.Value }
            }
            
            if ($null -eq $recordId) {
                $errorMsg = "レコードIDが取得できません: $($item.Key)"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "WARNING"
                continue
            }
            
            # 更新元データをキントーンAPI形式に変換（全項目）
            # JsonConverterモジュールの関数を使用
            $updateRecord = ConvertTo-WrappedJsonObject -InputObject $sourceItem
            # $updateRecordからline_lot_numberを削除
            $updateRecord.PSObject.Properties.Remove('line_lot_number')
            # $updateRecordをrecordキーでラップ
            $updateRecord = @{record = $updateRecord }
            # updateKeyを直接追加（ハッシュテーブルなので直接キーを設定）
            $updateRecord['updateKey'] = [PSCustomObject]@{field = 'line_lot_number'; value = $recordId }
            $updateRecords += $updateRecord

        }

        # $updateRecordsをjsonファイルに出力
        $updateRecords | ConvertTo-Json -Depth 20 | Out-File -FilePath "updateRecords.json"
        
        if ($updateRecords.Count -eq 0) {
            Write-Log -Message "更新対象データがありません" -LogPath $LogPath -LogLevel "WARNING"
            $result.Success = $true
            return $result
        }
        
        $totalBatches = [Math]::Ceiling($updateRecords.Count / $BatchSize)
        
        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $BatchSize
            $endIndex = [Math]::Min($startIndex + $BatchSize - 1, $updateRecords.Count - 1)
            $batchData = $updateRecords[$startIndex..$endIndex]
            $batchNumber = $batchIndex + 1
            
            try {
                # JSON変換
                $resultObject = @{
                    app     = $AppId
                    records = $batchData
                }
                $jsonData = $resultObject | ConvertTo-Json -Depth 20
                # $jsonDataをjsonファイルに出力
                $jsonData | Out-File -FilePath "jsonData.json"
                # API PUT送信
                Send-ApiRequest `
                    -Uri $ApiUri `
                    -Method "PUT" `
                    -Body $jsonData `
                    -Headers $ApiHeaders `
                    -TimeoutSec $TimeoutSec `
                    -Verbose | Out-Null
                
                $result.ProcessedCount += $batchData.Count
                $result.ProcessedData += $DataToUpdate[$startIndex..$endIndex]
            }
            catch {
                $errorMsg = "更新処理 バッチ $batchNumber エラー: $_"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
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

function Invoke-DeleteData {
    <#
    .SYNOPSIS
    削除処理を実行します（API DELETE、100件単位）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$DataToDelete,
        
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
        # 削除対象のレコードIDを抽出
        $recordIds = @()
        foreach ($item in $DataToDelete) {
            # キントーンAPIのレスポンス形式を考慮
            $recordId = $null
            if ($item.PSObject.Properties['$id']) {
                $idProp = $item.PSObject.Properties['$id']
                $recordId = if ($idProp.Value.value) { $idProp.Value.value } else { $idProp.Value }
            }
            
            if ($null -ne $recordId) {
                $recordIds += $recordId
            }
        }
        
        if ($recordIds.Count -eq 0) {
            Write-Log -Message "削除対象のレコードIDが取得できませんでした" -LogPath $LogPath -LogLevel "WARNING"
            $result.Success = $true
            return $result
        }
        
        $totalBatches = [Math]::Ceiling($recordIds.Count / $BatchSize)
        
        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $BatchSize
            $endIndex = [Math]::Min($startIndex + $BatchSize - 1, $recordIds.Count - 1)
            $batchIds = $recordIds[$startIndex..$endIndex]
            $batchNumber = $batchIndex + 1
            
            Write-Log -Message "削除処理 バッチ $batchNumber / $totalBatches（$($batchIds.Count) 件）" -LogPath $LogPath -LogLevel "INFO"
            
            try {
                # キントーンAPIのDELETE形式
                $bodyObject = @{
                    app = $AppId
                    ids = $batchIds
                }
                $jsonData = $bodyObject | ConvertTo-Json -Depth 10
                
                # API DELETE送信
                Send-ApiRequest `
                    -Uri $ApiUri `
                    -Method "DELETE" `
                    -Body $jsonData `
                    -Headers $ApiHeaders `
                    -TimeoutSec $TimeoutSec `
                    -Verbose | Out-Null
                
                $result.ProcessedCount += $batchIds.Count
                $result.ProcessedData += $DataToDelete[$startIndex..$endIndex]
                Write-Log -Message "削除処理 バッチ $batchNumber 成功（$($batchIds.Count) 件）" -LogPath $LogPath -LogLevel "INFO"
            }
            catch {
                $errorMsg = "削除処理 バッチ $batchNumber エラー: $_"
                Write-Log -Message $errorMsg -LogPath $LogPath -LogLevel "ERROR"
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

# モジュールをエクスポート
Export-ModuleMember -Function Sync-DataWithApi, Get-TargetDataFromApi, Compare-DataByLineLotNumber, Invoke-AddData, Invoke-UpdateData, Invoke-DeleteData

