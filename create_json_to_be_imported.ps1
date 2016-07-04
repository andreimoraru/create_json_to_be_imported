# to run the script, ensure ExecutionPolicy allows that:
# Get-ExecutionPolicy
# Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

# Adding PS Snapin

# This is a common function i am using which will release excel objects

function Release-Ref ($ref) {

([System.Runtime.InteropServices.Marshal]::ReleaseComObject(

[System.__ComObject]$ref) -gt 0)

[System.GC]::Collect()

[System.GC]::WaitForPendingFinalizers()

}

# set the name of sheet from Excel Workbook
$sheetName = "POLL DETAILS"

# set the number of the first row containing the polls - from here the loop iterating through the list of polls will begin
$startRow = 14

# set the title of the excel workbook
$workbookTitle = “Service Transition ASSET Setup.xlsm”

# Directory location where we have our excel files. Note: path should have trailing slash
$ExcelFilesLocation = “some_path_here”

# set the client name as listed in Augmenta
$clientName = "some_value"

# set path to json file
$pathJson = "some_path"

# set name of json file
$titleJson = "some_title.json"

# Creating excel object

$objExcel = New-Object -ComObject Excel.Application 

$objExcel.Visible = $false

# Open our excel file

$UserWorkBook = $objExcel.Workbooks.Open($ExcelFilesLocation + $workbookTitle)

# Here Item(1) refers to sheet 1 of of the workbook. If we want to access sheet 10, we have to modifythe code to Item(10)

$UserSheet = $UserWorkBook.Sheets.Item($sheetName)

# Do-While loop: iterate through row starting from row $startRow up to last row containing any value in column 3 (HOST NAME in the POLL DETAILS sheet of excel workbook)

Do {

# if the cell matching 3rd column in $startRow 
 if ($UserSheet.Cells.Item($startRow,3).Value()) {

    # Reading the third column of the current row - HOST NAME column in POLL DETAILS sheet of excel workbook

    $asset_Name = $UserSheet.Cells.Item($startRow, 3).Value()

    # Reading the 5th column of the current row - POLL FIELD NAME column in POLL DETAILS sheet of excel workbook

    $poll_Name = $UserSheet.Cells.Item($startRow, 5).Value()

    # Reading the 7th column of the current row - POLL DESCRIPTION column in POLL DETAILS sheet of excel workbook

    $pollDescription = $UserSheet.Cells.Item($startRow, 7).Value()

"""$clientName"": {`n`t""$asset_Name" + "-" + "$poll_Name"" :  {`n`t`t""asset_name"": ""$asset_Name"",`n`t`t""poll_name"": ""$poll_Name"",`n`t`t""description"" : ""$pollDescription"",`n`t`t""is_suspended"" : true`n`t}`n}" | Out-File -Append -FilePath $pathJson + $titleJson
  }


           # Move to next row

           $startRow++

           } While ($startRow -lt ($UserSheet.UsedRange).SpecialCells([Microsoft.Office.Interop.Excel.Constants]::xlLastCell).Row)

 

# Exiting the excel object

$objExcel.Quit()

 #Release all the objects used above

$a = Release-Ref($UserSheet)

$a = Release-Ref($UserWorkBook) 

$a = Release-Ref($objExcel)