function Get-ApiData {
    param(
        [string]$Uri,
        [string]$Method,
        [hashtable]$Headers,
        [int]$TimeoutSec
    )

    throw 'test stub: Get-ApiData should be mocked before use'
}

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
        ($keys -contains '2026-09-09|GC01') | Should Be $false
    }
}

InModuleScope ScheduleSyncManager {
    Describe 'Get-App86ScheduleRecords' {
        It 'filters out records whose lot_number is empty' {
            $script:CapturedRequestUri = $null
            Mock Get-ApiData {
                param($Uri, $Method, $Headers, $TimeoutSec)
                $script:CapturedRequestUri = $Uri
                [PSCustomObject]@{ records = @() }
            }

            $records = @(Get-App86ScheduleRecords `
                -ApiUri 'https://example.cybozu.com/k/v1/records.json' `
                -ApiHeaders @{} `
                -AppId 86)

            $records.Count | Should Be 0
            $decoded = [System.Uri]::UnescapeDataString($script:CapturedRequestUri)
            $decoded | Should Match 'lot_number != ""'
        }
    }

    Describe 'Schedule update payload' {
        It 'keeps every table row and writable source value while omitting the Calc field' {
            $record = [PSCustomObject]@{
                '$id'       = [PSCustomObject]@{ value = '10' }
                '$revision' = [PSCustomObject]@{ value = '7' }
                sub_schedule = [PSCustomObject]@{
                    value = @(
                        [PSCustomObject]@{
                            id = 'row-a'
                            value = [PSCustomObject]@{
                                sub_schedule_date = [PSCustomObject]@{ value = '2026-09-11' }
                                sub_lot_volume    = [PSCustomObject]@{ value = '3200' }
                                sub_index         = [PSCustomObject]@{ value = '9' }
                                sub_change_time   = [PSCustomObject]@{ value = '0.5' }
                                生産時間          = [PSCustomObject]@{ value = '3.38' }
                                予定開始日時      = [PSCustomObject]@{ value = '' }
                                予定終了日時      = [PSCustomObject]@{ value = '' }
                                休憩時間          = [PSCustomObject]@{ value = '' }
                            }
                        }
                        [PSCustomObject]@{
                            id = 'row-b'
                            value = [PSCustomObject]@{
                                sub_schedule_date = [PSCustomObject]@{ value = '2026-09-12' }
                                sub_lot_volume    = [PSCustomObject]@{ value = '760' }
                                sub_index         = [PSCustomObject]@{ value = '9' }
                                sub_change_time   = [PSCustomObject]@{ value = '1.0' }
                                生産時間          = [PSCustomObject]@{ value = '0.80' }
                                予定開始日時      = [PSCustomObject]@{ value = '2026-09-11T23:30:00Z' }
                                予定終了日時      = [PSCustomObject]@{ value = '2026-09-12T00:30:00Z' }
                                休憩時間          = [PSCustomObject]@{ value = '0' }
                            }
                        }
                    )
                }
            }

            $calculation = [PSCustomObject]@{
                RecordId     = '10'
                RowId        = 'row-a'
                StartAt      = [DateTimeOffset]::new(2026, 9, 11, 8, 30, 0, [TimeSpan]::FromHours(9))
                EndAt        = [DateTimeOffset]::new(2026, 9, 11, 10, 10, 0, [TimeSpan]::FromHours(9))
                BreakMinutes = 10
            }

            $updates = @(ConvertTo-ScheduleUpdateRecords -Records @($record) -Calculations @($calculation))
            $rows = @($updates[0].record.sub_schedule.value)

            $updates.Count | Should Be 1
            $updates[0].revision | Should Be '7'
            $rows.Count | Should Be 2
            $rows[0].id | Should Be 'row-a'
            $rows[1].id | Should Be 'row-b'

            $rows[0].value.sub_schedule_date.value | Should Be '2026-09-11'
            $rows[0].value.sub_lot_volume.value | Should Be '3200'
            $rows[0].value.sub_change_time.value | Should Be '0.5'
            ($null -eq $rows[0].value.PSObject.Properties['生産時間']) | Should Be $true
            $rows[0].value.予定開始日時.value | Should Be '2026-09-10T23:30:00Z'
            $rows[0].value.予定終了日時.value | Should Be '2026-09-11T01:10:00Z'
            $rows[0].value.休憩時間.value | Should Be '10'

            # 計算対象外の行も既存値を維持する。
            $rows[1].value.sub_schedule_date.value | Should Be '2026-09-12'
            $rows[1].value.予定開始日時.value | Should Be '2026-09-11T23:30:00Z'
            $rows[1].value.予定終了日時.value | Should Be '2026-09-12T00:30:00Z'
        }
    }
}
