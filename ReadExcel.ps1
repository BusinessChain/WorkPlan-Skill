param([Parameter(Mandatory)][string]$Path)

Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'

function Write-Shape($shape, $cell) {
    if (-not $shape.Visible) { return }
    $text = ""
    try { if ($shape.TextFrame2.HasText) { $text = "$($shape.TextFrame2.TextRange.Text)" -replace "`r", "`n" } } catch {}
    if ($shape.Type -eq 17) {
        "--- TEXTBOX: $($shape.Name)"
        $text
    }
    elseif ($shape.Type -eq 6) {
        foreach ($item in $shape.GroupItems) { Write-Shape $item $cell }
    }
    elseif ($text) {
        "--- SHAPE at ${cell}: $($shape.Name) | $text"
    }
}

$excel = New-Object -ComObject Excel.Application
$excelProcessId = 0
[void][Win32.User32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
$excel.DisplayAlerts = $false
try {
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    "=== NAMES"
    foreach ($name in $workbook.Names) {
        if (-not $name.Name.StartsWith("_xlfn.")) { "$($name.Name) -> $($name.RefersTo)" }
    }
    foreach ($sheet in $workbook.Worksheets) {
        "=== SHEET: $($sheet.Name)"
        $lastRow = $sheet.Cells.Find("*", [Type]::Missing, -4123, 2, 1, 2)
        if ($null -eq $lastRow) { "(no cells)" }
        else {
            $lastColumn = $sheet.Cells.Find("*", [Type]::Missing, -4123, 2, 2, 2)
            $rowCount = $lastRow.Row
            $columnCount = $lastColumn.Column
            $formulas = $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($rowCount, $columnCount)).Formula
            for ($i = 1; $i -le $rowCount; $i++) {
                $row = for ($j = 1; $j -le $columnCount; $j++) { "$($formulas[$i, $j])" }
                if (($row -join '').Trim()) { "R$i | " + ($row -join ' | ') }
            }
        }
        foreach ($shape in $sheet.Shapes) {
            Write-Shape $shape $shape.TopLeftCell.Address($false, $false)
        }
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
    Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
}
