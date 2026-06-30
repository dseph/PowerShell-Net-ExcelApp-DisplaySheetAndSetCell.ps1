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

try {
    if ($WorkbookPath) {
        $resolvedPath = (Resolve-Path -Path $WorkbookPath).Path
        $workbook = $excel.Workbooks.Open($resolvedPath)
    }
    else {
        $workbook = $excel.Workbooks.Add()
    }

    $worksheet = $workbook.Worksheets.Item($WorksheetName)
    $worksheet.Activate() | Out-Null
    $worksheet.Range($CellAddress).Value2 = $Text
}
catch {
    throw
}
