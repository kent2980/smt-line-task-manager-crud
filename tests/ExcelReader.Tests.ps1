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

    Describe 'Get-SubScheduleChangeTime' {
        It 'uses the Excel AP change time for the first scheduled cell' {
            Get-SubScheduleChangeTime `
                -ChangeTime 0.5 `
                -CurrentColumn 9 `
                -PreviousScheduledColumn $null | Should Be 0.5
        }

        It 'uses zero when the scheduled cell is immediately to the right of the previous scheduled cell' {
            Get-SubScheduleChangeTime `
                -ChangeTime 0.5 `
                -CurrentColumn 10 `
                -PreviousScheduledColumn 9 | Should Be 0
        }

        It 'uses the Excel AP change time again when there is an empty column between scheduled cells' {
            Get-SubScheduleChangeTime `
                -ChangeTime 0.5 `
                -CurrentColumn 11 `
                -PreviousScheduledColumn 9 | Should Be 0.5
        }

        It 'keeps zero for every continuation cell in a multi-day consecutive run' {
            $first = Get-SubScheduleChangeTime -ChangeTime 0.75 -CurrentColumn 9 -PreviousScheduledColumn $null
            $second = Get-SubScheduleChangeTime -ChangeTime 0.75 -CurrentColumn 10 -PreviousScheduledColumn 9
            $third = Get-SubScheduleChangeTime -ChangeTime 0.75 -CurrentColumn 11 -PreviousScheduledColumn 10

            $first | Should Be 0.75
            $second | Should Be 0
            $third | Should Be 0
        }
    }
}
