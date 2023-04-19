function Extract-ExcelTextAndNumberColumnsToNewWorkbook-FileA {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory=$true)]
      [string]$SourceWorkbook,
  
      [Parameter(Mandatory=$true)]
      [string]$Column1Number,

      [Parameter(Mandatory=$true)]
      [string]$Column2Number

  
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

    Write-Output "First Column is : ${Column1Number}"
    Write-Output "Second Column is : ${Column2Number}"

    # Get the values in column Column1Number
    $ColumnX = $SourceWorksheet.Range("${Column1Number}:${Column1Number}").Value2
    $ColumnY = $SourceWorksheet.Range("${Column2Number}:${Column2Number}").Value2
    $LastRow =  $SourceWorksheet.Cells.Find("*", [System.Type]::Missing, [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues, [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart, [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows, [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlPrevious, $false, $false, [System.Type]::Missing).Row

    Write-Output "Column ${Column1Number} Range is : $LastRow"
    $ColumnXNumber = $ColumnX | ForEach-Object {
        if ($_ -eq $null -or $_ -eq "" -or $_ -eq "N/A") {
            return $null  # Skip empty, blank, or N/A cells
        }
        $_
    }

    # $index = 0
    # foreach ($value in $ColumnXNumber) {
    #     if ($index -ge $LastRow) {
    #         break
    #     }
    #     Write-Output "${index}: $value "
    #     $index++
    # }
 

    Write-Output "Column ${Column2Number} Range is : $LastRow"
    $ColumnYNumber = $ColumnY  | ForEach-Object {
        
        if ($_ -eq $null -or $_ -eq "" -or $_ -eq 1 -or $_ -eq "-" ) {
            return $null  # Skip empty, blank, or N/A cells
        }
    
        if ($_ -as [double]) {
            "{0:N3}%" -f ($_ * 100) 
        }
        else {
            $_
        }
    }

    $index = 0
    foreach ($value in $ColumnYNumber) {
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

    # Write the values to the new worksheet
    $index = 0
    foreach ($value in $ColumnXNumber) {
    if ($value -ne $null -and $ColumnYNumber[$index] -ne $null -and $ColumnYNumber[$index] -ne "100.000%") {
        $newWorksheet.Cells.Item($index,1).Value2 = $value
        # $newWorksheet.Cells.Item($index,2).Value2 = $ColumnYNumber[$index]
        
        Write-Output "$value : $($ColumnYNumber[$index])"

        $index++
        if ($index -ge $LastRow) {
             break
         }
    }
    }
    

  
    # Get the name of the source workbook without the extension
    $SourceWorkbookName = [System.IO.Path]::GetFileNameWithoutExtension($SourceWorkbook)
  
    # Append "-new" to the source workbook name to create the destination workbook name
    $DestinationWorkbookName = $SourceWorkbookName + "-col-${Column1Number}-and-${Column2Number}-combined.xlsx"
  
    # Save the destination workbook
    $newWorksheet.SaveAs($DestinationWorkbookName)
  
    # Close the workbooks and quit Excel
    # $newWorksheet.Close($false)
    $newExcel.Quit()
  }
  