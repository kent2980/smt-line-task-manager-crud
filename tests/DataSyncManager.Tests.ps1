$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\DataSyncManager.psm1'
Import-Module $modulePath -Force

Describe 'Compare-DataByLineLotNumber' {
    It 'returns only add and update results and retains target-only records' {
        $sourceData = @(
            [PSCustomObject]@{ line_name = 'GC01'; lot_number = '001' }
            [PSCustomObject]@{ line_name = 'GC01'; lot_number = '002' }
        )
        $targetData = @(
            [PSCustomObject]@{
                line_name  = [PSCustomObject]@{ value = 'GC01' }
                lot_number = [PSCustomObject]@{ value = '002' }
            }
            [PSCustomObject]@{
                line_name  = [PSCustomObject]@{ value = 'GC01' }
                lot_number = [PSCustomObject]@{ value = '003' }
            }
        )

        $result = Compare-DataByLineLotNumber `
            -SourceData $sourceData `
            -TargetData $targetData `
            -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 1
        $result.ToAdd[0].lot_number | Should Be '001'
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01002'
        ($null -eq $result.PSObject.Properties['ToDelete']) | Should Be $true
    }
}
