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
$percentFormat = "{0:P0}"   # Set the format for displaying percentages
$columnYPercent = $columnY | ForEach-Object { $percentFormat -f $_ }

# Print the values in column Y as percentages
$columnYPercent | ForEach-Object { Write-Output $_ }

# Close the workbook and quit Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
