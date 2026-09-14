$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules\DataSyncManager.psm1'
Import-Module $modulePath -Force

function Write-Log {
    param(
        [string]$Message,
        [string]$LogPath,
        [string]$LogLevel
    )
}

function New-SourceScheduleRow {
    param(
        [string]$Date,
        [string]$Volume = '100',
        [string]$Index = '1',
        [string]$ChangeTime = '0.5'
    )

    [PSCustomObject]@{
        sub_schedule_date = $Date
        sub_lot_volume    = $Volume
        sub_index         = $Index
        sub_change_time   = $ChangeTime
        有効判定          = 'True'
    }
}

function New-SourceRecord {
    param(
        [string]$Key,
        [array]$Rows = @(),
        [string]$RangeStart = '2026-09-14',
        [string]$RangeEnd = '2026-10-06'
    )

    [PSCustomObject]@{
        line_lot_number        = $Key
        line_name              = 'GC01'
        lot_number             = $Key
        index                  = 1
        model_name             = 'MODEL'
        standard_date          = '2026-09-14'
        sub_schedule           = $Rows
        sync_date_range_start  = $RangeStart
        sync_date_range_end    = $RangeEnd
    }
}

function New-TargetScheduleRow {
    param(
        [string]$Id,
        [string]$Date,
        [string]$Active = 'True',
        [string]$StartAt = '2026-09-14T00:00:00Z'
    )

    [PSCustomObject]@{
        id = $Id
        value = [PSCustomObject]@{
            sub_schedule_date = [PSCustomObject]@{ value = $Date }
            sub_lot_volume    = [PSCustomObject]@{ value = '50' }
            sub_index         = [PSCustomObject]@{ value = '9' }
            sub_change_time   = [PSCustomObject]@{ value = '0' }
            有効判定          = [PSCustomObject]@{ value = $Active }
            予定開始日時      = [PSCustomObject]@{ value = $StartAt }
            予定終了日時      = [PSCustomObject]@{ value = '2026-09-14T01:00:00Z' }
            休憩時間          = [PSCustomObject]@{ value = '0' }
            生産時間          = [PSCustomObject]@{ value = '1' }
        }
    }
}

function New-TargetRecord {
    param(
        [string]$Key,
        [array]$Rows = @()
    )

    [PSCustomObject]@{
        line_lot_number = [PSCustomObject]@{ value = $Key }
        line_name       = [PSCustomObject]@{ value = 'GC01' }
        lot_number      = [PSCustomObject]@{ value = $Key }
        sub_schedule    = [PSCustomObject]@{ value = $Rows }
    }
}

Describe 'Compare-DataByLineLotNumber' {
    It 'line_lot_number一致を従来どおり通常更新対象にする' {
        $sourceData = @(
            New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-15'
            )
        )
        $targetData = @(
            New-TargetRecord -Key 'GC01001' -Rows @(
                New-TargetScheduleRow -Id 'row-1' -Date '2026-09-15'
            )
        )

        $result = Compare-DataByLineLotNumber -SourceData $sourceData -TargetData $targetData -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 0
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01001'
        $result.ToUpdate[0].UpdateMode | Should Be 'Full'
    }

    It 'line_lot_numberが一致しなくてもExcel日付範囲内の既存行があれば更新対象にする' {
        $sourceData = @(
            New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-15'
            )
        )
        $targetData = @(
            New-TargetRecord -Key 'GC01999' -Rows @(
                New-TargetScheduleRow -Id 'row-old' -Date '2026-09-20'
            )
        )

        $result = Compare-DataByLineLotNumber -SourceData $sourceData -TargetData $targetData -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 1
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01999'
        $result.ToUpdate[0].UpdateMode | Should Be 'ScheduleOnly'
        $null -eq $result.ToUpdate[0].Source | Should Be $true
    }

    It 'Excel日付範囲外の既存行しかないレコードは追加条件の更新対象にしない' {
        $sourceData = @(
            New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-15'
            )
        )
        $targetData = @(
            New-TargetRecord -Key 'GC01999' -Rows @(
                New-TargetScheduleRow -Id 'row-old' -Date '2026-09-01'
            )
        )

        $result = Compare-DataByLineLotNumber -SourceData $sourceData -TargetData $targetData -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 1
        $result.ToUpdate.Count | Should Be 0
    }

    It 'line_lot_numberの前後空白を除去して比較する' {
        $sourceData = @(
            New-SourceRecord -Key '  GC01002  ' -Rows @()
        )
        $targetData = @(
            New-TargetRecord -Key 'GC01002' -Rows @()
        )

        $result = Compare-DataByLineLotNumber -SourceData $sourceData -TargetData $targetData -LogPath 'test.log'

        $result.ToAdd.Count | Should Be 0
        $result.ToUpdate.Count | Should Be 1
        $result.ToUpdate[0].Key | Should Be 'GC01002'
    }

    It '更新元のline_lot_numberが空の場合はエラーにする' {
        $sourceData = @([PSCustomObject]@{ line_lot_number = '   ' })

        {
            Compare-DataByLineLotNumber -SourceData $sourceData -TargetData @() -LogPath 'test.log'
        } | Should Throw
    }

    It '更新元に同一line_lot_numberが複数ある場合はエラーにする' {
        $sourceData = @(
            [PSCustomObject]@{ line_lot_number = 'GC01002' }
            [PSCustomObject]@{ line_lot_number = 'GC01002' }
        )

        {
            Compare-DataByLineLotNumber -SourceData $sourceData -TargetData @() -LogPath 'test.log'
        } | Should Throw
    }

    It '更新先に同一line_lot_numberが複数ある場合はエラーにする' {
        $sourceData = @([PSCustomObject]@{ line_lot_number = 'GC01002' })
        $targetData = @(
            New-TargetRecord -Key 'GC01002' -Rows @()
            New-TargetRecord -Key 'GC01002' -Rows @()
        )

        {
            Compare-DataByLineLotNumber -SourceData $sourceData -TargetData $targetData -LogPath 'test.log'
        } | Should Throw
    }
}

InModuleScope DataSyncManager {
    Describe 'sub_schedule履歴保持payload' {
        It '既存行を削除せず範囲内だけFalseにし新予定をTrueで末尾追加する' {
            $source = New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-20' -Volume '120'
            )
            $target = New-TargetRecord -Key 'GC01001' -Rows @(
                New-TargetScheduleRow -Id 'row-before' -Date '2026-09-01' -Active 'True'
                New-TargetScheduleRow -Id 'row-in-range' -Date '2026-09-20' -Active 'True'
                New-TargetScheduleRow -Id 'row-history' -Date '2026-09-21' -Active 'False'
            )

            $comparison = Compare-DataByLineLotNumber -SourceData @($source) -TargetData @($target) -LogPath 'test.log'
            $payload = ConvertTo-UpdateRecordPayload -UpdateItem $comparison.ToUpdate[0]
            $rows = @($payload.record.sub_schedule.value)

            $rows.Count | Should Be 4
            $rows[0].id | Should Be 'row-before'
            $rows[0].value.有効判定.value | Should Be 'True'
            $rows[1].id | Should Be 'row-in-range'
            $rows[1].value.有効判定.value | Should Be 'False'
            $rows[2].id | Should Be 'row-history'
            $rows[2].value.有効判定.value | Should Be 'False'
            $null -eq $rows[3].PSObject.Properties['id'] | Should Be $true
            $rows[3].value.sub_schedule_date.value | Should Be '2026-09-20'
            $rows[3].value.sub_lot_volume.value | Should Be '120'
            $rows[3].value.有効判定.value | Should Be 'True'
        }

        It '日付範囲条件だけで選ばれたレコードはsub_schedule以外をPUTしない' {
            $source = New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-20'
            )
            $target = New-TargetRecord -Key 'GC01999' -Rows @(
                New-TargetScheduleRow -Id 'row-old' -Date '2026-09-20' -Active 'True'
            )

            $comparison = Compare-DataByLineLotNumber -SourceData @($source) -TargetData @($target) -LogPath 'test.log'
            $scheduleOnly = @($comparison.ToUpdate | Where-Object { $_.Key -eq 'GC01999' })[0]
            $payload = ConvertTo-UpdateRecordPayload -UpdateItem $scheduleOnly

            $payload.updateKey.value | Should Be 'GC01999'
            @($payload.record.PSObject.Properties).Count | Should Be 1
            $null -eq $payload.record.PSObject.Properties['sub_schedule'] | Should Be $false
            $payload.record.sub_schedule.value[0].id | Should Be 'row-old'
            $payload.record.sub_schedule.value[0].value.有効判定.value | Should Be 'False'
        }

        It '新規POST用の予定行はTrueにし同期メタデータはpayloadへ含めない' {
            $source = New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-20'
            )

            $payload = ConvertTo-AddRecordPayload -SourceItem $source

            $payload.sub_schedule.value[0].value.有効判定.value | Should Be 'True'
            $null -eq $payload.PSObject.Properties['sync_date_range_start'] | Should Be $true
            $null -eq $payload.PSObject.Properties['sync_date_range_end'] | Should Be $true
        }

        It '既存行に行IDが無い場合は物理削除リスクを避けて失敗する' {
            $source = New-SourceRecord -Key 'GC01001' -Rows @(
                New-SourceScheduleRow -Date '2026-09-20'
            )
            $target = New-TargetRecord -Key 'GC01001' -Rows @(
                [PSCustomObject]@{
                    value = [PSCustomObject]@{
                        sub_schedule_date = [PSCustomObject]@{ value = '2026-09-20' }
                    }
                }
            )
            $comparison = Compare-DataByLineLotNumber -SourceData @($source) -TargetData @($target) -LogPath 'test.log'

            { ConvertTo-UpdateRecordPayload -UpdateItem $comparison.ToUpdate[0] } | Should Throw
        }
    }
}