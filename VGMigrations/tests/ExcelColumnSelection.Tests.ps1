$ErrorActionPreference = 'Stop'

Describe 'Excel column selection' {
    It 'uses the configured header column for the selected workbook field' {
        $path = Join-Path $PSScriptRoot 'fixtures/column-selection.xlsx'
        if (-not (Test-Path $path)) {
            Set-ItResult -Skipped -Because 'fixture workbook is not present'
            return
        }

        $rows = & {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            try {
                $wb = $excel.Workbooks.Open((Resolve-Path $path).Path)
                $ws = $wb.Worksheets.Item(1)
                $used = $ws.UsedRange
                $headers = @{}
                for ($c = 1; $c -le $used.Columns.Count; $c++) {
                    $headers[[string]$used.Cells.Item(1, $c).Text] = $c
                }
                $rows = @()
                for ($r = 2; $r -le $used.Rows.Count; $r++) {
                    $rows += [PSCustomObject]@{
                        Value = [string]$used.Cells.Item($r, $headers['License']).Text
                    }
                }
                $rows
            } finally {
                $excel.Quit()
            }
        }

        $rows.Count | Should -BeGreaterThan 0
        $rows[0].Value | Should -Be 'J'
    }
}
