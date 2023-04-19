# Get the input file name from the command line
$inputFile = 'File-A.xlsx'

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Disable alerts and screen updating for a faster operation
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

# Open the input Excel file
$workbook = $excel.Workbooks.Open($inputFile)

# Get the first worksheet
$worksheet = $workbook.Sheets.Item(1)

# Get the data in columns "Site" and "Availability (MW)"
$data = $worksheet.Range("A:A,B:B").Value2

# Create a new Excel workbook
$newWorkbook = $excel.Workbooks.Add()

# Get the first worksheet in the new workbook
$newWorksheet = $newWorkbook.Sheets.Item(1)

# Paste the data in columns "Site" and "Availability (MW)"
$newWorksheet.Range("A:A,B:B").Value2 = $data

# Get the input file name without extension
$inputFileWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)

# Create the output file name with a "-new" postfix
$outputFile = $inputFileWithoutExtension + "-new.xlsx"

# Save the new workbook
$newWorkbook.SaveAs($outputFile)

# Close the workbooks and quit Excel
$workbook.Close()
$newWorkbook.Close()
$excel.Quit()

# Release the COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newWorksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
