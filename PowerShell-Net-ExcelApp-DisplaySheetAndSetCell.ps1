param(
    [string]$WorkbookPath,
    [string]$WorksheetName = 'Sheet1',
    [string]$CellAddress = 'A1',
    [Parameter(Mandatory = $true)]
    [string]$Text
)

$excel = New-Object -ComObject Excel.Application

$workbook = $null
$worksheet = $null
$hadError = $false

try {
    if ($WorkbookPath) {
        try {
            $resolvedPath = (Resolve-Path -Path $WorkbookPath -ErrorAction Stop).Path
        }
        catch {
            throw "Workbook file not found at path: $WorkbookPath"
        }

        $workbook = $excel.Workbooks.Open($resolvedPath)
    }
    else {
        $workbook = $excel.Workbooks.Add()
    }

    try {
        $worksheet = $workbook.Worksheets.Item($WorksheetName)
    }
    catch {
        throw "Worksheet '$WorksheetName' was not found in the workbook. $($_.Exception.Message)"
    }

    $excel.Visible = $true

    try {
        $worksheet.Activate() | Out-Null
    }
    catch {
        throw "Failed to activate worksheet '$WorksheetName'. $($_.Exception.Message)"
    }

    try {
        $worksheet.Range($CellAddress).Value2 = $Text
    }
    catch {
        throw "Failed to set value for cell '$CellAddress'. $($_.Exception.Message)"
    }
}
catch {
    $hadError = $true
    throw
}
finally {
    if ($hadError) {
        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }
    }

    if ($worksheet) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }

    if ($workbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
