$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\JsonConverter.psm1'
Import-Module $modulePath -Force

function New-TestRecord {
    param(
        [int]$Index = 99,
        [array]$SubSchedule
    )

    [PSCustomObject]@{
        line_name        = 'TEST'
        index            = $Index
        standard_date    = '2026-01-01'
        sub_schedule     = $SubSchedule
        line_lot_number  = 'TEST001'
    }
}

function New-TestSubSchedule {
    param(
        [string]$Date,
        [int]$Index
    )

    [PSCustomObject]@{
        sub_schedule_date = $Date
        sub_lot_volume    = 1
        sub_index         = $Index
    }
}

Describe 'JsonConverter index selection' {
    It 'selects today before a future date' {
        $today = (Get-Date).Date
        $record = New-TestRecord -SubSchedule @(
            New-TestSubSchedule -Date $today.AddDays(1).ToString('yyyy-MM-dd') -Index 2
            New-TestSubSchedule -Date $today.ToString('yyyy-MM-dd') -Index 1
        )

        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $wrapped.index.value | Should Be 1
    }

    It 'selects the nearest future date regardless of table order' {
        $today = (Get-Date).Date
        $record = New-TestRecord -SubSchedule @(
            New-TestSubSchedule -Date $today.AddDays(5).ToString('yyyy-MM-dd') -Index 5
            New-TestSubSchedule -Date $today.AddDays(2).ToString('yyyy-MM-dd') -Index 2
            New-TestSubSchedule -Date $today.AddDays(3).ToString('yyyy-MM-dd') -Index 3
        )

        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $wrapped.index.value | Should Be 2
    }

    It 'selects the latest past date when no current or future date exists' {
        $today = (Get-Date).Date
        $record = New-TestRecord -SubSchedule @(
            New-TestSubSchedule -Date $today.AddDays(-5).ToString('yyyy-MM-dd') -Index 5
            New-TestSubSchedule -Date $today.AddDays(-1).ToString('yyyy-MM-dd') -Index 1
            New-TestSubSchedule -Date $today.AddDays(-3).ToString('yyyy-MM-dd') -Index 3
        )

        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $wrapped.index.value | Should Be 1
    }

    It 'ignores invalid dates and preserves the original index when none are valid' {
        $record = New-TestRecord -Index 99 -SubSchedule @(
            New-TestSubSchedule -Date 'invalid' -Index 1
            New-TestSubSchedule -Date '2026/01/01' -Index 2
        )

        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $wrapped.index.value | Should Be 99
    }

    It 'uses the first table row when the selected date is duplicated' {
        $today = (Get-Date).Date
        $record = New-TestRecord -SubSchedule @(
            New-TestSubSchedule -Date $today.AddDays(1).ToString('yyyy-MM-dd') -Index 10
            New-TestSubSchedule -Date $today.AddDays(1).ToString('yyyy-MM-dd') -Index 20
        )

        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $wrapped.index.value | Should Be 10
    }

    It 'applies the same index to registration JSON and update wrapping without mutating input' {
        $today = (Get-Date).Date
        $record = New-TestRecord -Index 99 -SubSchedule @(
            New-TestSubSchedule -Date $today.AddDays(1).ToString('yyyy-MM-dd') -Index 7
        )

        $json = ConvertTo-JsonData -AppId 1 -InputObject $record -Depth 20 | ConvertFrom-Json
        $wrapped = ConvertTo-WrappedJsonObject -InputObject $record

        $json.records[0].index.value | Should Be 7
        $wrapped.index.value | Should Be 7
        $record.index | Should Be 99
    }
}
