# In this version of the script, we've added Write-Debug statements to each major step of the script to print a message indicating what's happening at each step. You can enable debug output by running the script with the -Debug switch, like this:

# This will print all of the debug messages to the console while the script runs, which can help you see what's happening and troubleshoot any issues that arise. If you don't want to see the debug output, you can run the script without the -Debug switch and it will run normally, without printing


# Install the ImportExcel module if it's not already installed
if (-not (Get-Module ImportExcel -ErrorAction SilentlyContinue)) {
  Write-Debug "ImportExcel module not found, installing..."
  Install-Module ImportExcel
}

# Set the paths to the input and output Excel files
$fileA = 'File-A.xlsx'
$fileB = 'File-B.xlsx'
$outputFile = 'outputFile.xlsx'

Write-Debug "Importing data from File A: $fileA"
# Import the data from File A and select the 'Site' and 'Availability (MW)' columns
$dataA = Import-Excel $fileA | Select-Object 'Site', 'Availability (MW)'

Write-Debug "Importing data from File B: $fileB"
# Import the data from File B and select the 'Site' and 'Combo Availability (24 Hrs)' columns
$dataB = Import-Excel $fileB | Select-Object 'Site', 'Combo Availability (24 Hrs)'

Write-Debug "Creating hashtable from data in File B"
# Create a hashtable from the data in File B, keyed by 'Site'
$hashB = @{}
foreach ($rowB in $dataB) {
  $hashB[$rowB.Site] = $rowB.'Combo Availability (24 Hrs)'
}

Write-Debug "Merging data from File A and File B"
# Merge the data from File A and File B, using the hashtable to lookup the 'Combo Availability (24 Hrs)' value
$outputData = @()
$totalRows = $dataA.Count
$currentRow = 0
foreach ($rowA in $dataA) {
  $currentRow++
  Write-Progress -Activity "Merging data" -Status "Processing row $currentRow of $totalRows" -PercentComplete ($currentRow / $totalRows * 100)

  Write-Debug "Merging row $($rowA.Site)"
  $comboAvailability = $hashB[$rowA.Site]
  if ($comboAvailability) {
    $outputData += [PSCustomObject]@{
      'Site'                        = $rowA.Site
      'Availability (MW)'           = $rowA.'Availability (MW)'
      'Combo Availability (24 Hrs)' = $comboAvailability
    }
  }
}

Write-Debug "Filtering data by availability"
# Filter the data to only include rows where 'Availability (MW)' is less than 99.90%
$outputData = $outputData | Where-Object { $_.'Availability (MW)' -lt 99.90 }

Write-Debug "Exporting merged data to Excel file: $outputFile"
# Export the merged data to a new Excel file
$outputData | Export-Excel -Path $outputFile -AutoSize -AutoFilter

Write-Debug "Opening output file: $outputFile"
# Open the output file
Invoke-Item $outputFile
