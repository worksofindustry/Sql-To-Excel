<#
.Synopsis
   Exports SQL to Excel
.DESCRIPTION
    This script template will save the results of a .sql file and save it to Excel. Script can be modified for additional
    formating of Excel.
    To use this executable you'll need to install the module ImportExcel onto your system:
    > Install-Module ImportExcel
    If you don't have administrative access, you can clone the code repo from GitHub @ https://github.com/dfinke/ImportExcel
    and import as so: Import-Module C:\USERS_GIT_REPO_Location\ImportExcel-master\ImportExcel
.NOTES
    Thanks to Matthew Linker for creating this script.
.PARAMETER save_to_folder
    the directory where to save the Excel document
.PARAMETER filename
    Output Excel File Name, do not include file extension .xlsx
.PARAMETER server
    Name of your SQL Server Instance
.PARAMETER database
    Database name, ex. AdventureWorks
.PARAMETER sqlfile
    location of .sql file ex. C:\somedir\my.sql
.PARAMETER worksheetname
    Name of Excel worksheet, Limited to 30 char by Excel
.PARAMETER append_date
    Optional, default false. If you would like to add date of creation to file name, ex. 'Myfile 20200228.xlsx'. Acceptable values: 0,1,$false,$true
.EXAMPLE
   .\HUSA_Order_Creation.ps1 -save_to_folder 'C:\Users\USERNAME\Desktop' -fileName 'My SQL Export' -server localhost -database 'Beta database' `
    -sqlfile 'MySqlScript.sql' -worksheetname Sheet1 -append_date $true
.FUNCTIONALITY
    Excel
#>
[cmdletbinding(ConfirmImpact = 'Medium', SupportsPaging = $true, SupportsShouldProcess = $true)]
param ([string] $save_to_folder, [string] $fileName, [string] $server='SSDSQL01', [string] $database, [string] $sqlfile, [string] $worksheetname, [bool] $append_date=$false)
Write-Verbose "Generating the Excel file now"

Import-Module ImportExcel

if ([bool]([System.Uri]$save_to_folder).IsUnc) 
{
pushd $save_to_folder
}
Set-Location $save_to_folder

$a = (Get-Date).Day 
$b = get-date -format "yyyy"
$date = "$b$a"

if ($append_date -eq $true)
{
$fileName = $fileName+" "+$date+".xlsx"
}
else
{
$fileName = $fileName+".xlsx"
}

$excelfile =(New-Item -Path . -Name $fileName -ItemType "file" -Force)


Invoke-Sqlcmd -Inputfile $sqlfile -ServerInstance $server -Database $database | Export-XLSX -WorkSheetName $worksheetname `
-Path $excelfile -AutoFit -TableStyle None -Force
          

# Format Workbook
$Excel = New-Excel -Path $excelfile

    # Get a Worksheet
    $Worksheet = $Excel | Get-Worksheet -Name $worksheetname 

    # Freeze the top row
    $Worksheet | Set-FreezePane -Row 2

$Excel | Close-Excel -Save

#cleanup temp UNC path mount
popd