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
        [ValidateSet("GET", "POST", "PUT", "DELETE", "PATCH")]
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
        [ValidateSet("GET", "POST", "PUT", "DELETE", "PATCH")]
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

# モジュールをエクスポート
Export-ModuleMember -Function Send-ApiRequest, Get-ApiData

