
[CmdletBinding()]
param(
  [Parameter(Mandatory=$false, Position=0)]
  # [string]$FileA = "C:\Users\shoai\Documents\powershell\File-A-mini2.xlsx",

  [string]$FileA = "C:\Users\shoai\Documents\powershell\File-A.xlsx",

  [Parameter(Mandatory=$false, Position=1)]
  [string]$FileB = "C:\Users\shoai\Documents\powershell\File-B.xlsx",

  [Parameter(Mandatory=$false, Position=2)]
  [string]$FileAColumn1 = "Site",

  [Parameter(Mandatory=$false, Position=5)]
  [string]$FileAColumn1Number = "F",

  [Parameter(Mandatory=$false, Position=3)]
  [string]$FileAColumn2 = "Availability (MW)",

  [Parameter(Mandatory=$false, Position=5)]
  [string]$FileAColumn2Number = "Y",

  [Parameter(Mandatory=$false, Position=4)]
  [string]$FileBColumn1Name = "Site",

  [Parameter(Mandatory=$false, Position=5)]
  [string]$FileBColumn1Number = "A",

  [Parameter(Mandatory=$false, Position=6)]
  [string]$FileBColumn2 = "Combo Availability (24 Hrs)",

  [Parameter(Mandatory=$false, Position=6)]
  [string]$FileBColumn2Number = "E"
)



# Load the Create-ExcelWorkbook function from a separate file
. "C:\Users\shoai\Documents\powershell\Create-ExcelWorkbook.ps1"
. "C:\Users\shoai\Documents\powershell\Fill-ExcelWorkbook.ps1"
. "C:\Users\shoai\Documents\powershell\Copy-ExcelTextColumnsToNewWorkbook.ps1"
. "C:\Users\shoai\Documents\powershell\Copy-ExcelNumberColumnsToNewWorkbook.ps1"

. "C:\Users\shoai\Documents\powershell\Extract-ExcelTextAndNumberColumnsToNewWorkbook.ps1"
. "C:\Users\shoai\Documents\powershell\Extract-ExcelTextAndNumberColumnsToNewWorkbook-FileA.ps1"


# $data1 = Read-ExcelColumns -FileName "" -Column1Name "Site" -Column2Name "Availability (MW)"

# Call the Create-ExcelWorkbook function with the provided columns 
# Files are blank
# Create-ExcelWorkbook -WorkbookName $FileA -Column1Name $FileAColumn1 -Column2Name $FileAColumn2

# Create-ExcelWorkbook -WorkbookName $FileB -Column1Name $FileBColumn1 -Column2Name $FileBColumn2

# Now read the File A and B and fill in the new files with columns

# Read-ExcelColumnsToHashtable -WorkbookName $FileB -Column1Name $FileBColumn1 -Column2Name $FileBColumn2

# Copy-ExcelTextColumnsToNewWorkbook -SourceWorkbook $FileB -ColumnNumber $FileBColumn1Number

# Copy-ExcelNumberColumnsToNewWorkbook -SourceWorkbook $FileB -ColumnNumber $FileBColumn2Number

# Extract-ExcelTextAndNumberColumnsToNewWorkbook -SourceWorkbook $FileB -Column1Number $FileBColumn1Number -Column2Number $FileBColumn2Number

Extract-ExcelTextAndNumberColumnsToNewWorkbook-FileA -SourceWorkbook $FileA -Column1Number $FileAColumn1Number -Column2Number $FileAColumn2Number