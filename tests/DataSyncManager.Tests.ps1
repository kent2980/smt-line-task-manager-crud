$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\DataSyncManager.psm1'
Import-Module $modulePath -Force

# エラー系テストではCompare-DataByLineLotNumber内のログ出力が呼ばれるため、
# このテストファイル単体でも実行できるよう最小限のスタブを用意する。
function Write-Log {
    param(
        [string]$Message,
        [string]$LogPath,
        [string]$LogLevel
    )
}

Describe 'Compare-DataByLineLotNumber' {
    It 'line_lot_numberを基準に追加・更新を振り分け、更新先だけのレコードは削除対象にしない' {
        $sourceData = @(
            [PSCustomObject]@{
                line_lot_number = 'GC01001'
                line_name       = 'GC01'
                lot_number      = '001'
            }
            [PSCustomObject]@{
                line_lot_number = 'GC01002'
                line_name       = 'GC01'
                lot_number      = '002'
            }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01002' }
                line_name       = [PSCustomObject]@{ value = 'GC01' }
                lot_number      = [PSCustomObject]@{ value = '002' }
            }
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01003' }
                line_name       = [PSCustomObject]@{ value = 'GC01' }
                lot_number      = [PSCustomObject]@{ value = '003' }
            }
        )

        $result = Compare-DataByLineLotNumber `
            -SourceData $sourceData `
            -TargetData $targetData `
            -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 1
        $result.ToAdd[0].line_lot_number | Should Be 'GC01001'
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01002'
        ($null -eq $result.PSObject.Properties['ToDelete']) | Should Be $true
    }

    It 'line_nameやlot_numberが異なっていてもline_lot_numberが一致すれば更新対象にする' {
        $sourceData = @(
            [PSCustomObject]@{
                line_lot_number = 'GC01002'
                line_name       = 'GC01'
                lot_number      = '002'
            }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01002' }
                line_name       = [PSCustomObject]@{ value = 'GC-01' }
                lot_number      = [PSCustomObject]@{ value = '999' }
            }
        )

        $result = Compare-DataByLineLotNumber `
            -SourceData $sourceData `
            -TargetData $targetData `
            -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 0
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01002'
    }

    It 'line_nameとlot_numberの結合値が一致してもline_lot_numberが異なれば追加対象にする' {
        $sourceData = @(
            [PSCustomObject]@{
                line_lot_number = 'SOURCE-KEY'
                line_name       = 'GC01'
                lot_number      = '002'
            }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'TARGET-KEY' }
                line_name       = [PSCustomObject]@{ value = 'GC01' }
                lot_number      = [PSCustomObject]@{ value = '002' }
            }
        )

        $result = Compare-DataByLineLotNumber `
            -SourceData $sourceData `
            -TargetData $targetData `
            -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 1
        $result.ToAdd[0].line_lot_number | Should Be 'SOURCE-KEY'
        $result.ToUpdate.Count | Should Be 0
    }

    It 'line_lot_numberの前後空白を除去して比較する' {
        $sourceData = @(
            [PSCustomObject]@{ line_lot_number = '  GC01002  ' }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01002' }
            }
        )

        $result = Compare-DataByLineLotNumber `
            -SourceData $sourceData `
            -TargetData $targetData `
            -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 0
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01002'
    }

    It '更新元のline_lot_numberが空の場合は同期対象にせずエラーにする' {
        $sourceData = @(
            [PSCustomObject]@{ line_lot_number = '   ' }
        )

        {
            Compare-DataByLineLotNumber `
                -SourceData $sourceData `
                -TargetData @() `
                -LogPath 'test.log'
        } | Should Throw
    }

    It '更新元に同一line_lot_numberが複数ある場合は暗黙に上書きせずエラーにする' {
        $sourceData = @(
            [PSCustomObject]@{ line_lot_number = 'GC01002' }
            [PSCustomObject]@{ line_lot_number = 'GC01002' }
        )

        {
            Compare-DataByLineLotNumber `
                -SourceData $sourceData `
                -TargetData @() `
                -LogPath 'test.log'
        } | Should Throw
    }

    It '更新先に同一line_lot_numberが複数ある場合は曖昧な更新を行わずエラーにする' {
        $sourceData = @(
            [PSCustomObject]@{ line_lot_number = 'GC01002' }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01002' }
            }
            [PSCustomObject]@{
                line_lot_number = [PSCustomObject]@{ value = 'GC01002' }
            }
        )

        {
            Compare-DataByLineLotNumber `
                -SourceData $sourceData `
                -TargetData $targetData `
                -LogPath 'test.log'
        } | Should Throw
    }
}
