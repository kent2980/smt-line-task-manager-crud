$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\ExcelReader.psm1'
Import-Module $modulePath -Force

InModuleScope ExcelReader {
    Describe 'ConvertTo-SyncLotNumber' {
        It 'returns null for null' {
            $result = ConvertTo-SyncLotNumber -Value $null
            $null -eq $result | Should Be $true
        }

        It 'returns null for whitespace' {
            $result = ConvertTo-SyncLotNumber -Value '   '
            $null -eq $result | Should Be $true
        }

        It 'trims a valid lot number' {
            ConvertTo-SyncLotNumber -Value '  001234  ' | Should Be '001234'
        }
    }
}
