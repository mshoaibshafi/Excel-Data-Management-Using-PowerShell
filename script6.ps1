# Get the file name from the command-line arguments
$fileName = $args[0]

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Disable alerts and screen updating for a faster operation
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open($fileName)

# Get the first worksheet
$worksheet = $workbook.Sheets.Item(1)

# Get the values in column Y as percentages
$columnY = $worksheet.Range("Y:Y").Value2
$columnYPercent = $columnY | ForEach-Object { 
    if ($_ -eq $null -or $_ -eq "" -or $_ -eq 1 -or $_ -eq "-" ) {
        return $null  # Skip empty, blank, 100% or - cells
    }
    "{0:N3}%" -f ($_ * 100) 
}

# Get the values in column C as numbers
$columnC = $worksheet.Range("C:C").Value2
$columnCNumber = $columnC | ForEach-Object {
    if ($_ -eq $null -or $_ -eq "" -or $_ -eq "N/A") {
        return $null  # Skip empty, blank, or N/A cells
    }
    $_
}

# Print the values in column C and their corresponding values in column Y
$index = 0
foreach ($value in $columnCNumber) {
    if ($value -ne $null -and $columnYPercent[$index] -ne $null -and $columnYPercent[$index] -ne "100.000%") {
        Write-Output "$value : $($columnYPercent[$index])"
    }
    $index++
}

# Close the workbook and quit Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Create a new Excel object and add a new workbook
$newExcel = New-Object -ComObject Excel.Application
$newWorkbook = $newExcel.Workbooks.Add()

# Get the first worksheet in the new workbook
$newWorksheet = $newWorkbook.Sheets.Item(1)

# Set the column names
$newWorksheet.Cells.Item(1,1).Value2 = "Site"
$newWorksheet.Cells.Item(1,2).Value2 = "Availability"

# Write the values to the new worksheet
$index = 2
foreach ($value in $columnCNumber) {
    if ($value -ne $null -and $columnYPercent[$index-2] -ne $null -and $columnYPercent[$index-2] -ne "100.000%") {
        $newWorksheet.Cells.Item($index,1).Value2 = $value
        $newWorksheet.Cells.Item($index,2).Value2 = $columnYPercent[$index-2]
        $index++
    }
}
$index--

# Generate the output file name with a date postfix
$date = Get-Date -Format "yyyy-MM-dd"
$outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + "-new-" + $date + [System.IO.Path]::GetExtension($fileName)

# Save the new workbook as a CSV file
$newWorkbook.SaveAs("$outputFileName", 6)  # 6 = CSV file format

# Close the new workbook and quit Excel
$newWorkbook.Close()
$newExcel.Quit()

# Release the COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newWorksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newExcel) | Out-Null

