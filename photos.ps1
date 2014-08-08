# ==============================================================================================
# 
# Microsoft PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 4.1
# 
# NAME: OrgPhotos.ps1
# 
# UPDATED: Kyle Buzby
# DATE: 24 July 2014
# COMMENT: Changed file paths and combined with the Build-ExcelArray.ps1 script as noted below. Script
#          takes photos from dropbox folder as specified and moves them to photos in specific folder
#          based on excel document of specific folder locations (Inspectors.xlsx)
#
# AUTHOR:  Kim Oppalfens, 
# DATE  : 12/2/2007
# 
# COMMENT: Helps you organise your digital photos into subdirectory, based on the Exif data 
# found inside the picture. Based on the date picture taken property the pictures will be organized into
# c:\RecentlyUploadedPhotos\YYYY\YYYY-MM-DD

#combined with
#Script name: Build-ExcelArray.ps1 
#Created on: Thursday, May 03, 2007 
#Author: Kent Finkle 
#Purpose: How can I use Windows Powershell to Build an Array from a Column of Data in Excel? 
# ============================================================================================== 

$x = [reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 

$dropbox = "C:\Users\kbuzby\Dropbox\"
$zDrive = "Z:\Project Photos\"  
$Inspectors = @()
$JobFolder = @()
$CommonName = @()
 

function Release-Ref ($ref) { 
([System.Runtime.InteropServices.Marshal]::ReleaseComObject( 
[System.__ComObject]$ref) -gt 0) 
[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers() 
} 

function MoveFiles([string]$source, [string]$target, [string]$common) 
{
	$Files = Get-ChildItem $source -recurse -filter *.jpg
	if ($Files -ne $null) 
	{
	foreach ($file in $Files) 
	{
		$foo = New-Object -TypeName system.drawing.bitmap -ArgumentList $file.fullname 
 
		#each character represents an ascii code number 0-10 is date 
		#10th character is space separator between date and time
		#48 = 0 49 = 1 50 = 2 51 = 3 52 = 4 53 = 5 54 = 6 55 = 7 56 = 8 57 = 9 58 = : 
		#date is in YYYY/MM/DD format
		$date = $foo.GetPropertyItem(36867).value[0..9]
		$arYear = [Char]$date[0],[Char]$date[1],[Char]$date[2],[Char]$date[3]
		$arMonth = [Char]$date[5],[Char]$date[6]
		$arDay = [Char]$date[8],[Char]$date[9]
		$strYear = [String]::Join("",$arYear)
		$strMonth = [String]::Join("",$arMonth) 
		$strDay = [String]::Join("",$arDay)
		$DateTaken = $strYear + "." + $strMonth + "." + $strDay
		$TargetPath = $target + "\" + $DateTaken
   
		$foo.dispose()
 
		If (Test-Path $TargetPath)
		{
			Move-Item $file.FullName -destination $TargetPath -force
		}
		Else
		{
			New-Item $TargetPath -Type Directory
			Move-Item $file.FullName -destination $TargetPath -force
		}
	} 
	write-host "$common photos updated."
	}
}
 


$objExcel = new-object -comobject excel.application  
$objExcel.Visible = $False  
$objWorkbook = $objExcel.Workbooks.Open("Z:\Project Photos\Inspectors.xlsx") 
$objWorksheet = $objWorkbook.Worksheets.Item(1) 
 
$i = 1 
 
Do 	
{ 
	$Inspectors += $objWorksheet.Cells.Item($i, 2).Value()
	$JobFolder += $objWorksheet.Cells.Item($i, 3).Value()
	$CommonName += $objWorksheet.Cells.Item($i,4).Value()
    $i++ 
} 
While ($objWorksheet.Cells.Item($i,2).Value() -ne $null) 
 
$a = $objExcel.Quit() 
 
write-host "Inspectors Folders Updated"
 
$a = Release-Ref($objWorksheet) 
$a = Release-Ref($objWorkbook) 
$a = Release-Ref($objExcel) 

$i = 0

foreach ($objItem in $Inspectors)
{
	$source = ($dropbox + $Inspectors[$i])
	$target = ($zDrive + $JobFolder[$i])
	$common = ($CommonName[$i])
	$a = MoveFiles $source $target $common
	$i++
}

write-host "Photos have been updated, press any key to continue..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")


