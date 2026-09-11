$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\ScheduleSyncManager.psm1'
Import-Module $modulePath -Force

function New-SourceRecord {
    param(
        [string]$Key,
        [string]$Line,
        [array]$Dates
    )

    [PSCustomObject]@{
        line_lot_number = $Key
        line_name       = $Line
        sub_schedule    = @(
            $Dates | ForEach-Object {
                [PSCustomObject]@{
                    sub_schedule_date = $_
                    sub_lot_volume    = 1
                    sub_index         = 1
                    sub_change_time   = 0
                }
            }
        )
    }
}

function New-TargetRecord {
    param(
        [string]$Key,
        [string]$Line,
        [array]$Dates
    )

    [PSCustomObject]@{
        line_lot_number = [PSCustomObject]@{ value = $Key }
        line_name       = [PSCustomObject]@{ value = $Line }
        sub_schedule    = [PSCustomObject]@{
            value = @(
                $Dates | ForEach-Object {
                    [PSCustomObject]@{
                        id = "row-$_"
                        value = [PSCustomObject]@{
                            sub_schedule_date = [PSCustomObject]@{ value = $_ }
                        }
                    }
                }
            )
        }
    }
}

Describe 'Get-AffectedScheduleGroups' {
    It 'includes both old and new groups when a schedule is reallocated' {
        $source = @(
            New-SourceRecord -Key 'GC01001' -Line 'GC01' -Dates @('2026-09-11', '2026-09-12')
        )
        $target = @(
            New-TargetRecord -Key 'GC01001' -Line 'GC01' -Dates @('2026-09-10', '2026-09-11')
        )

        $groups = @(Get-AffectedScheduleGroups -SourceData $source -TargetData $target)
        $keys = @($groups | ForEach-Object { "$($_.Date)|$($_.LineName)" })

        $groups.Count | Should Be 3
        $keys | Should Contain '2026-09-10|GC01'
        $keys | Should Contain '2026-09-11|GC01'
        $keys | Should Contain '2026-09-12|GC01'
    }

    It 'does not include target-only records because the Excel sync does not update them' {
        $source = @(
            New-SourceRecord -Key 'GC01001' -Line 'GC01' -Dates @('2026-09-11')
        )
        $target = @(
            New-TargetRecord -Key 'GC01099' -Line 'GC01' -Dates @('2026-09-09')
        )

        $groups = @(Get-AffectedScheduleGroups -SourceData $source -TargetData $target)
        $keys = @($groups | ForEach-Object { "$($_.Date)|$($_.LineName)" })

        $groups.Count | Should Be 1
        $keys | Should Contain '2026-09-11|GC01'
        $keys | Should Not Contain '2026-09-09|GC01'
    }
}
