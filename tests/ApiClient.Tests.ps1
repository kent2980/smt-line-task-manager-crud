$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\ApiClient.psm1'
Import-Module $modulePath -Force

Describe 'Split-KintoneRecordBatches' {
    It '0件の場合は空配列を返す' {
        $batches = @(Split-KintoneRecordBatches -Records @() -AppId 86)
        $batches.Count | Should Be 0
    }

    It '上限内の複数レコードは1バッチにまとめる' {
        $records = @(
            [PSCustomObject]@{ record = [PSCustomObject]@{ name = [PSCustomObject]@{ value = 'A' } } }
            [PSCustomObject]@{ record = [PSCustomObject]@{ name = [PSCustomObject]@{ value = 'B' } } }
        )

        $batches = @(Split-KintoneRecordBatches -Records $records -AppId 86 -MaxRecords 100 -MaxBodyBytes 1MB)

        $batches.Count | Should Be 1
        @($batches[0]).Count | Should Be 2
    }

    It 'レコード件数上限で分割する' {
        $records = 1..5 | ForEach-Object { [PSCustomObject]@{ id = $_ } }

        $batches = @(Split-KintoneRecordBatches -Records $records -AppId 86 -MaxRecords 2 -MaxBodyBytes 1MB)

        $batches.Count | Should Be 3
        @($batches[0]).Count | Should Be 2
        @($batches[1]).Count | Should Be 2
        @($batches[2]).Count | Should Be 1
    }

    It '100件未満でもUTF-8 JSONサイズ上限を超える場合は複数バッチへ分割する' {
        $records = @(
            [PSCustomObject]@{ record = [PSCustomObject]@{ text = [PSCustomObject]@{ value = ('あ' * 200) } } }
            [PSCustomObject]@{ record = [PSCustomObject]@{ text = [PSCustomObject]@{ value = ('い' * 200) } } }
            [PSCustomObject]@{ record = [PSCustomObject]@{ text = [PSCustomObject]@{ value = ('う' * 200) } } }
        )
        $singleBody = @{ app = 86; records = @($records[0]) } | ConvertTo-Json -Depth 20
        $singleBytes = [System.Text.Encoding]::UTF8.GetByteCount($singleBody)
        $maxBytes = $singleBytes + 50

        $batches = @(Split-KintoneRecordBatches -Records $records -AppId 86 -MaxRecords 100 -MaxBodyBytes $maxBytes)

        $batches.Count | Should Be 3
        foreach ($batch in $batches) {
            $body = @{ app = 86; records = @($batch) } | ConvertTo-Json -Depth 20
            [System.Text.Encoding]::UTF8.GetByteCount($body) | Should BeLessThan ($maxBytes + 1)
        }
    }

    It '1レコードだけでサイズ上限を超える場合はAPI送信前にエラーにする' {
        $record = [PSCustomObject]@{
            record = [PSCustomObject]@{
                text = [PSCustomObject]@{ value = ('あ' * 1000) }
            }
        }

        { Split-KintoneRecordBatches -Records @($record) -AppId 86 -MaxBodyBytes 100 } |
            Should Throw '*1レコードがサイズ上限を超えています*'
    }

    It '不正なサイズ上限はエラーにする' {
        { Split-KintoneRecordBatches -Records @([PSCustomObject]@{ id = 1 }) -AppId 86 -MaxBodyBytes 0 } |
            Should Throw '*MaxBodyBytesは1以上*'
    }
}
