# PowerShell-Net-ExcelApp-DisplaySheetAndSetCell.ps1

# ----------------------------------------
# This script demonstrates how to use .NET to create an instance of Excel, 
# display a worksheet, and set a cell value. It was generated using CoPilot.  
# ----------------------------------------

try {
    # Get COM type from ProgID
    $excelType = [System.Type]::GetTypeFromProgID("Excel.Application")

    if (-not $excelType) {
        throw "Excel COM type not found. Is Excel installed and registered?"
    }

    # Create instance using .NET Activator
    $excel = [System.Activator]::CreateInstance($excelType)

    # Make Excel visible
    $excel.Visible = $true

    Write-Host "Excel started successfully (via .NET Activator)"

    # Add workbook
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)

    # Write test value
    $sheet.Cells.Item(1,1).Value2 = "Hello from PowerShell (.NET)"

    Start-Sleep 5

}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
