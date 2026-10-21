# ApiClient.psm1
# API送信を行うモジュール

function Send-ApiRequest {
    <#
    .SYNOPSIS
    APIにリクエストを送信します。
    
    .DESCRIPTION
    Invoke-RestMethodを使用してAPIにJSONデータを送信します。
    
    .PARAMETER Uri
    APIのエンドポイントURL
    
    .PARAMETER Method
    HTTPメソッド（デフォルト: POST）
    
    .PARAMETER Body
    送信するJSONデータ
    
    .PARAMETER Headers
    追加のHTTPヘッダー（ハッシュテーブル）
    
    .PARAMETER ContentType
    コンテンツタイプ（デフォルト: application/json）
    
    .PARAMETER TimeoutSec
    タイムアウト秒数（デフォルト: 30）
    
    .EXAMPLE
    Send-ApiRequest -Uri "https://api.example.com/data" -Body $jsonData
    
    .EXAMPLE
    $headers = @{ "Authorization" = "Bearer token123" }
    Send-ApiRequest -Uri "https://api.example.com/data" -Body $jsonData -Headers $headers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("GET", "POST", "PUT", "PATCH")]
        [string]$Method = "POST",
        
        [Parameter(Mandatory = $false)]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json; charset=utf-8",
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30
    )
    
    try {
        Write-Verbose "APIリクエスト送信中: $Method $Uri"
        
        $params = @{
            Uri = $Uri
            Method = $Method
            TimeoutSec = $TimeoutSec
            ErrorAction = "Stop"
        }
        
        # GETはContent-Type不要（必要なメソッドのみ付与）
        if ($Method -ne "GET") {
            $params.ContentType = $ContentType
        }
        
        # Bodyが指定されている場合は追加
        if ($Body) {
            if ($Method -eq "GET") {
                $params.Body = $Body
            }
            else {
                # 日本語文字化けを避けるため、UTF-8バイト列で送信
                $params.Body = [System.Text.Encoding]::UTF8.GetBytes($Body)
            }
        }
        
        # Headersが指定されている場合は追加
        if ($Headers) {
            # GETではContent-Typeヘッダーを送らない
            if ($Method -eq "GET" -and $Headers.ContainsKey("Content-Type")) {
                $requestHeaders = @{}
                foreach ($key in $Headers.Keys) {
                    if ($key -ne "Content-Type") {
                        $requestHeaders[$key] = $Headers[$key]
                    }
                }
                $params.Headers = $requestHeaders
            }
            else {
                $params.Headers = $Headers
            }
        }
        
        $response = Invoke-RestMethod @params
        
        Write-Verbose "APIリクエスト成功"
        return $response
    }
    catch {
        $errorMessage = "API送信エラー: $_"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $errorMessage += " (ステータスコード: $statusCode)"
        }
        throw $errorMessage
    }
}

function Get-ApiData {
    <#
    .SYNOPSIS
    APIからデータを取得します。
    
    .DESCRIPTION
    Invoke-RestMethodを使用してAPIからデータを取得します（GETリクエスト）。
    
    .PARAMETER Uri
    APIのエンドポイントURL
    
    .PARAMETER Method
    HTTPメソッド（デフォルト: GET）
    
    .PARAMETER Body
    送信するJSONデータ（GETリクエストの場合、通常はクエリパラメータとして使用）
    
    .PARAMETER Headers
    追加のHTTPヘッダー（ハッシュテーブル）
    
    .PARAMETER ContentType
    コンテンツタイプ（デフォルト: application/json）
    
    .PARAMETER TimeoutSec
    タイムアウト秒数（デフォルト: 30）
    
    .EXAMPLE
    $data = Get-ApiData -Uri "https://api.example.com/data"
    
    .EXAMPLE
    $headers = @{ "Authorization" = "Bearer token123" }
    $body = @{ app = "app123"; query = "line_name = 'GC01'" } | ConvertTo-Json
    $data = Get-ApiData -Uri "https://api.example.com/data" -Body $body -Headers $headers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("GET", "POST", "PUT", "PATCH")]
        [string]$Method = "GET",
        
        [Parameter(Mandatory = $false)]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 30
    )
    
    try {
        Write-Verbose "APIデータ取得中: $Method $Uri"
        
        $params = @{
            Uri = $Uri
            Method = $Method
            TimeoutSec = $TimeoutSec
            ErrorAction = "Stop"
        }
        
        # GETはContent-Type不要（必要なメソッドのみ付与）
        if ($Method -ne "GET") {
            $params.ContentType = $ContentType
        }
        
        # Bodyが指定されている場合は追加
        if ($Body) {
            $params.Body = $Body
        }
        
        # Headersが指定されている場合は追加
        if ($Headers) {
            # GETではContent-Typeヘッダーを送らない
            if ($Method -eq "GET" -and $Headers.ContainsKey("Content-Type")) {
                $requestHeaders = @{}
                foreach ($key in $Headers.Keys) {
                    if ($key -ne "Content-Type") {
                        $requestHeaders[$key] = $Headers[$key]
                    }
                }
                $params.Headers = $requestHeaders
            }
            else {
                $params.Headers = $Headers
            }
        }
        
        $response = Invoke-RestMethod @params
        
        Write-Verbose "APIデータ取得成功"
        return $response
    }
    catch {
        $errorMessage = "APIデータ取得エラー: $_"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $errorMessage += " (ステータスコード: $statusCode)"
        }
        throw $errorMessage
    }
}


function Split-KintoneRecordBatches {
    <#
    .SYNOPSIS
    kintone APIへ送信するレコードをリクエストサイズ上限内のバッチへ分割します。

    .DESCRIPTION
    レコード件数と、app・recordsを含む実際のJSONをUTF-8へ変換した
    バイトサイズの両方を確認し、各APIリクエストが指定上限を
    超えないようレコード配列を分割します。

    1レコードだけでサイズ上限を超える場合は、それ以上分割できないため
    API送信前に例外を送出します。この関数自体はAPI通信を行いません。

    .PARAMETER Records
    kintoneへ送信するレコードの一覧。

    .PARAMETER AppId
    送信先kintoneアプリID。

    .PARAMETER MaxRecords
    1リクエストあたりの最大レコード件数。既定値は100件です。

    .PARAMETER MaxBodyBytes
    1リクエストあたりの最大JSON UTF-8バイト数。既定値は45MBです。

    .OUTPUTS
    各APIリクエストに使用するレコード配列の一覧。

    .EXAMPLE
    $batches = @(Split-KintoneRecordBatches -Records $records -AppId 86)
    foreach ($batch in $batches) {
        # $batch を records としてAPIへ送信する
    }

    .NOTES
    JSONサイズはConvertTo-Json -Depth 20で生成した実際のリクエスト本文を
    UTF-8へ変換したバイト数で判定します。

    .THROWS
    MaxRecordsまたはMaxBodyBytesが1未満の場合。
    1レコードだけでMaxBodyBytesを超過する場合。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Records,

        [Parameter(Mandatory = $true)]
        [int]$AppId,

        [Parameter(Mandatory = $false)]
        [int]$MaxRecords = 100,

        [Parameter(Mandatory = $false)]
        [long]$MaxBodyBytes = 45MB
    )

    if ($MaxRecords -lt 1) {
        throw 'MaxRecordsは1以上を指定してください。'
    }
    if ($MaxBodyBytes -lt 1) {
        throw 'MaxBodyBytesは1以上を指定してください。'
    }
    if ($null -eq $Records -or $Records.Count -eq 0) {
        return
    }

    $batches = [System.Collections.ArrayList]::new()
    $currentBatch = [System.Collections.ArrayList]::new()

    foreach ($record in $Records) {
        $candidate = @($currentBatch.ToArray()) + @($record)
        $candidateJson = @{
            app     = $AppId
            records = $candidate
        } | ConvertTo-Json -Depth 20
        $candidateBytes = [System.Text.Encoding]::UTF8.GetByteCount($candidateJson)

        if ($candidate.Count -le $MaxRecords -and $candidateBytes -le $MaxBodyBytes) {
            [void]$currentBatch.Add($record)
            continue
        }

        if ($currentBatch.Count -gt 0) {
            [void]$batches.Add(@($currentBatch.ToArray()))
            $currentBatch.Clear()
        }

        $singleJson = @{
            app     = $AppId
            records = @($record)
        } | ConvertTo-Json -Depth 20
        $singleBytes = [System.Text.Encoding]::UTF8.GetByteCount($singleJson)
        if ($singleBytes -gt $MaxBodyBytes) {
            throw ("kintone API送信対象の1レコードがサイズ上限を超えています。requestSize={0} bytes, maxBodyBytes={1} bytes" -f $singleBytes, $MaxBodyBytes)
        }

        [void]$currentBatch.Add($record)
    }

    if ($currentBatch.Count -gt 0) {
        [void]$batches.Add(@($currentBatch.ToArray()))
    }

    foreach ($batch in $batches) {
        Write-Output -NoEnumerate $batch
    }
}

# モジュールをエクスポート
Export-ModuleMember -Function Send-ApiRequest, Get-ApiData, Split-KintoneRecordBatches

