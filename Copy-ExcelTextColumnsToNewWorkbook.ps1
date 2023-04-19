function Copy-ExcelTextColumnsToNewWorkbook {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory=$true)]
      [string]$SourceWorkbook,
  
      [Parameter(Mandatory=$true)]
      [string]$ColumnNumber
  
    )
  
    # Load the Excel module
    # Import-Module -Name "Microsoft.Office.Interop.Excel"
  
    # Create a new Excel Application object
    $Excel = New-Object -ComObject Excel.Application
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    # Disable alerts and screen updating for a faster operation
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false
  
    # Open the source workbook
    $SourceWorkbookObj = $Excel.Workbooks.Open($SourceWorkbook)
  
    # Get the source worksheet
    $SourceWorksheet = $SourceWorkbookObj.Sheets.Item(1)

    Write-Host "Column 1 name: $Column1Name"

    # Get the values in column Column1Number
    $columnA = $SourceWorksheet.Range("${ColumnNumber}:${ColumnNumber}").Value2
    $LastRow =  $SourceWorksheet.Cells.Find("*", [System.Type]::Missing, [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues, [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart, [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows, [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious, $false, $false, [System.Type]::Missing).Row

    Write-Output "Column ${ColumnNumber} Range is : $LastRow"
    $columnANumber = $columnA | ForEach-Object {
        if ($_ -eq $null -or $_ -eq "" -or $_ -eq "N/A") {
            return $null  # Skip empty, blank, or N/A cells
        }
        $_
    }

    $index = 0
    foreach ($value in $columnANumber) {
        if ($index -ge $LastRow) {
            break
        }
        Write-Output "${index}: $value "
        $index++
    }


    # Close the workbook and quit Excel
    $SourceWorkbookObj.Close()
    $Excel.Quit()

    # Release the COM objects from memory
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($SourceWorksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($SourceWorkbookObj) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

  
    # Create a new Excel object and add a new workbook
    $newExcel = New-Object -ComObject Excel.Application
    $newWorkbook = $newExcel.Workbooks.Add()

    # Get the first worksheet in the new workbook
    $newWorksheet = $newWorkbook.Sheets.Item(1)

    # # Set the column names
    # $newWorksheet.Cells.Item(1,1).Value2 = $Column1Name


    $index = 1
    foreach ($value in $columnANumber) {
        if ($index -gt $LastRow) {
            break
        }
        $newWorksheet.Cells.Item($index,1).Value2 = $value
        $index++
    }
  
    # Get the name of the source workbook without the extension
    $SourceWorkbookName = [System.IO.Path]::GetFileNameWithoutExtension($SourceWorkbook)
  
    # Append "-new" to the source workbook name to create the destination workbook name
    $DestinationWorkbookName = $SourceWorkbookName + "-col-${ColumnNumber}-new.xlsx"
  
    # Save the destination workbook
    $newWorksheet.SaveAs($DestinationWorkbookName)
  
    # Close the workbooks and quit Excel
    $newWorksheet.Close($false)
    $newExcel.Quit()
  }
  