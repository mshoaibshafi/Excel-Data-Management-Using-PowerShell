function Fill-ExcelWorkbook {
    param(
      [Parameter(Mandatory=$true)]
      [string]$WorkbookName,
      [Parameter(Mandatory=$true)]
      [array]$Column1Data,
      [Parameter(Mandatory=$true)]
      [array]$Column2Data
    )
  
    # Load the Excel COM object
    $excel = New-Object -ComObject Excel.Application
  
    # Open the workbook
    $workbook = $excel.Workbooks.Open($WorkbookName)
  
    # Get the first worksheet in the workbook
    $worksheet = $workbook.Worksheets.Item(1)
  
    # Set the column data
    $rowCount = $Column1Data.Length
    for ($i = 0; $i -lt $rowCount; $i++) {
      $worksheet.Cells.Item($i+2,1) = $Column1Data[$i]
      $worksheet.Cells.Item($i+2,2) = $Column2Data[$i]
    }
  
    # Auto-fit the column widths
    $worksheet.Columns.Item(1).AutoFit()
    $worksheet.Columns.Item(2).AutoFit()
  
    # Save and close the workbook
    $workbook.Save()
    $workbook.Close()
  
    # Quit Excel
    $excel.Quit()
  }
  