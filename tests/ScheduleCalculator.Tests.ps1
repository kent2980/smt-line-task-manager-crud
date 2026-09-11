$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\ScheduleCalculator.psm1'
Import-Module $modulePath -Force

function New-ScheduleRow {
    param(
        [string]$Id,
        [string]$Date,
        [string]$ProductionHours,
        [object]$ChangeHours = '',
        [int]$SubIndex = 1
    )

    return [PSCustomObject]@{
        id = $Id
        value = [PSCustomObject]@{
            sub_schedule_date = [PSCustomObject]@{ value = $Date }
            sub_lot_volume    = [PSCustomObject]@{ value = '100' }
            sub_index         = [PSCustomObject]@{ value = [string]$SubIndex }
            sub_change_time   = [PSCustomObject]@{ value = $ChangeHours }
            生産時間          = [PSCustomObject]@{ value = $ProductionHours }
        }
    }
}

function New-ScheduleRecord {
    param(
        [string]$Id,
        [string]$LineName = 'GC01',
        [int]$Index,
        [array]$Rows,
        [string]$RecordType = ''
    )

    return [PSCustomObject]@{
        '$id'            = [PSCustomObject]@{ value = $Id }
        '$revision'      = [PSCustomObject]@{ value = '1' }
        record_type      = [PSCustomObject]@{ value = $RecordType }
        line_name        = [PSCustomObject]@{ value = $LineName }
        index            = [PSCustomObject]@{ value = [string]$Index }
        line_lot_number  = [PSCustomObject]@{ value = "$LineName-$Id" }
        sub_schedule     = [PSCustomObject]@{ value = $Rows }
    }
}

Describe 'ConvertTo-HourValue' {
    It 'accepts values with H suffix' {
        (ConvertTo-HourValue -Value '1.5 H' -Context 'test') | Should Be 1.5
    }

    It 'uses zero for an empty optional daily change time' {
        (ConvertTo-HourValue -Value '' -Context 'test' -AllowEmpty) | Should Be 0
    }

    It 'rejects invalid values' {
        { ConvertTo-HourValue -Value 'abc H' -Context 'test' } | Should Throw
    }
}

Describe 'App86 schedule calculation' {
    It 'starts each date-line group at 08:30 and carries the previous end to the next record' {
        $date = '2026-09-11'
        $records = @(
            New-ScheduleRecord -Id '1' -Index 1 -Rows @(
                New-ScheduleRow -Id 'row-1' -Date $date -ProductionHours '1.0 H' -ChangeHours '0.5'
            )
            New-ScheduleRecord -Id '2' -Index 2 -Rows @(
                New-ScheduleRow -Id 'row-2' -Date $date -ProductionHours '1.0' -ChangeHours ''
            )
            New-ScheduleRecord -Id '3' -Index 3 -Rows @(
                New-ScheduleRow -Id 'row-3' -Date $date -ProductionHours '1.5' -ChangeHours '0'
            )
        )

        $result = @(Get-App86ScheduleCalculations -Records $records)

        $result.Count | Should Be 3
        $result[0].StartAt.ToString('HH:mm') | Should Be '08:30'
        $result[0].EndAt.ToString('HH:mm') | Should Be '10:10'
        $result[0].BreakMinutes | Should Be 10

        $result[1].StartAt.ToString('HH:mm') | Should Be '10:10'
        $result[1].EndAt.ToString('HH:mm') | Should Be '11:10'
        $result[1].BreakMinutes | Should Be 0

        $result[2].StartAt.ToString('HH:mm') | Should Be '11:10'
        $result[2].EndAt.ToString('HH:mm') | Should Be '13:20'
        $result[2].BreakMinutes | Should Be 40
    }

    It 'includes a break when the nominal end exactly equals the break start' {
        $start = [DateTimeOffset]::new(2026, 9, 11, 14, 30, 0, [TimeSpan]::FromHours(9))

        $result = Add-WorkDurationWithBreaks -StartAt $start -WorkMinutes 30 -ScheduleDate '2026-09-11'

        $result.EndAt.ToString('HH:mm') | Should Be '15:10'
        $result.BreakMinutes | Should Be 10
    }

    It 'does not cap the result at the 17:30 regular end time' {
        $record = New-ScheduleRecord -Id '1' -Index 1 -Rows @(
            New-ScheduleRow -Id 'row-1' -Date '2026-09-11' -ProductionHours '10' -ChangeHours '0'
        )

        $result = @(Get-App86ScheduleCalculations -Records @($record))

        $result[0].StartAt.ToString('HH:mm') | Should Be '08:30'
        $result[0].EndAt.ToString('HH:mm') | Should Be '19:30'
        $result[0].BreakMinutes | Should Be 60
    }

    It 'calculates each date independently from 08:30' {
        $records = @(
            New-ScheduleRecord -Id '1' -Index 1 -Rows @(
                New-ScheduleRow -Id 'row-1' -Date '2026-09-11' -ProductionHours '10' -ChangeHours '0'
                New-ScheduleRow -Id 'row-2' -Date '2026-09-12' -ProductionHours '1' -ChangeHours '0'
            )
        )

        $result = @(Get-App86ScheduleCalculations -Records $records)
        $nextDay = @($result | Where-Object { $_.Date -eq '2026-09-12' })[0]

        $nextDay.StartAt.ToString('yyyy-MM-dd HH:mm') | Should Be '2026-09-12 08:30'
    }

    It 'excludes SETTING records' {
        $records = @(
            New-ScheduleRecord -Id '1' -Index 1 -RecordType 'SETTING' -Rows @(
                New-ScheduleRow -Id 'row-1' -Date '2026-09-11' -ProductionHours '1' -ChangeHours '0'
            )
        )

        $result = @(Get-App86ScheduleCalculations -Records $records)

        $result.Count | Should Be 0
    }

    It 'can calculate only the requested date-line groups' {
        $records = @(
            New-ScheduleRecord -Id '1' -Index 1 -Rows @(
                New-ScheduleRow -Id 'row-1' -Date '2026-09-11' -ProductionHours '1' -ChangeHours '0'
                New-ScheduleRow -Id 'row-2' -Date '2026-09-12' -ProductionHours '1' -ChangeHours '0'
            )
        )
        $groups = @([PSCustomObject]@{ Date = '2026-09-12'; LineName = 'GC01' })

        $result = @(Get-App86ScheduleCalculations -Records $records -Groups $groups)

        $result.Count | Should Be 1
        $result[0].Date | Should Be '2026-09-12'
        $result[0].StartAt.ToString('HH:mm') | Should Be '08:30'
    }
}
