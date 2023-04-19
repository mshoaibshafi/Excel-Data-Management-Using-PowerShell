# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Disable alerts and screen updating for a faster operation
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open("C:\Users\shoai\Documents\powershell\excel-data-management-using-powershell\File-A.xlsx")

# Get the first worksheet
$worksheet = $workbook.Sheets.Item(1)

# Get the values in the first column
$column1 = $worksheet.Range("A:A").Value2

# Print the values in the first column
$column1 | ForEach-Object { Write-Output $_ }

# Close the workbook and quit Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
