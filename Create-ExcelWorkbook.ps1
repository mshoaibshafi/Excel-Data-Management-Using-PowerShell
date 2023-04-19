function Create-ExcelWorkbook {
  param(
    [Parameter(Mandatory=$true)]
    [string]$WorkbookName,
    [Parameter(Mandatory=$true)]
    [string]$Column1Name,
    [Parameter(Mandatory=$true)]
    [string]$Column2Name  )
  
    # Load the Excel COM object
    $excel = New-Object -ComObject Excel.Application
  
    # Create a new workbook
    $workbook = $excel.Workbooks.Add()
  
    # Save a workbook with postfix "-new"
    $newWorkbookName = $workbookName -replace "\.xlsx$", "-new.xlsx"
    $workbook.SaveAs($newWorkbookName)

    # Get the first worksheet in the workbook
    $worksheet = $workbook.Worksheets.Item(1)
  
    # Set the column headers
    $worksheet.Cells.Item(1,1) = $Column1Name
    $worksheet.Cells.Item(1,2) = $Column2Name
  
    # Auto-fit the column widths
    $worksheet.Columns.Item(1).AutoFit()
    $worksheet.Columns.Item(2).AutoFit()
    
    # Save and close the workbook
    $workbook.Save()
    $workbook.Close()
  
    # Quit Excel
    $excel.Quit()
  }
  