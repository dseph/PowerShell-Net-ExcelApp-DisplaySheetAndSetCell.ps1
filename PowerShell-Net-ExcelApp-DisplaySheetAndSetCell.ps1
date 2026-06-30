param(
    [string]$WorkbookPath,
    [string]$WorksheetName = 'Sheet1',
    [string]$CellAddress = 'A1',
    [Parameter(Mandatory = $true)]
    [string]$Text
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$workbook = $null
$worksheet = $null

try {
    if ($WorkbookPath) {
        $resolvedPath = (Resolve-Path -Path $WorkbookPath).Path
        $workbook = $excel.Workbooks.Open($resolvedPath)
    }
    else {
        $workbook = $excel.Workbooks.Add()
    }

    try {
        $worksheet = $workbook.Worksheets.Item($WorksheetName)
    }
    catch {
        throw "Worksheet '$WorksheetName' was not found in the workbook."
    }

    $worksheet.Activate() | Out-Null
    try {
        $worksheet.Range($CellAddress).Value2 = $Text
    }
    catch {
        throw "Cell address '$CellAddress' is invalid."
    }
}
catch {
    if ($worksheet) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }

    if ($workbook) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    throw
}
