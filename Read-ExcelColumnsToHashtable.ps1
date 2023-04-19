function Read-ExcelColumnsToHashtable {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$WorkbookName,

    [Parameter(Mandatory=$true)]
    [string]$Column1Name,

    [Parameter(Mandatory=$true)]
    [string]$Column2Name
  )

  # Load the Excel module
  # Import-Module -Name "Microsoft.Office.Interop.Excel"

  # Create a new Excel Application object
  $Excel = New-Object -ComObject Excel.Application

  # Open the workbook
  $Workbook = $Excel.Workbooks.Open($WorkbookName)

  # Get the worksheet
  $Worksheet = $Workbook.Sheets.Item(1)

  # Get the range for the two columns
  $Column1Range = $Worksheet.Range($Column1Name + "1").EntireColumn
  $Column2Range = $Worksheet.Range($Column2Name + "1").EntireColumn

  # Read the data from the two columns into arrays
  $Column1Data = $Column1Range.Value2
  $Column2Data = $Column2Range.Value2

  # Create a hashtable to store the data
  $Data = @{}

  # Loop through the rows and add the data to the hashtable
  for ($i = 2; $i -le $Column1Data.Length; $i++) {
    $Column1Value = $Column1Data[$i, 1]
    $Column2Value = $Column2Data[$i, 1]
    $Data[$Column1Value] = $Column2Value
  }

  # Close the workbook and quit Excel
  $Workbook.Close($false)
  $Excel.Quit()

  # Output the hashtable
  # return $Data
  $Data
}
