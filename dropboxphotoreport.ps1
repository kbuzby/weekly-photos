$x = [reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 

#Get Global use variables

function Release-Ref ($ref) { 
([System.Runtime.InteropServices.Marshal]::ReleaseComObject( 
[System.__ComObject]$ref) -gt 0) 
[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers() 
} 

function MoveFiles([string]$source, [string]$target, [string]$common, [string]$engineer) {
	write-host "Checking for $common photos..."
	$Files = Get-ChildItem $source -recurse -filter *.jpg
	if ($Files -ne $null) 
	{
		write-host "Moving $common photos from Dropbox to Z:\ Drive..."
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
			if ($engineer -eq "Jason")
			{
				$DateTaken = "XX " + $strYear + "." + $strMonth + "." + $strDay
			}
			else {$DateTaken = $strYear + "." + $strMonth + "." + $strDay}
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
	}
	else {
		write-host "No $common photos here."
	}
}

$today = Get-Date 
$JobsCommon = @()
$Inspectors = @()
$Engineer = @()
$JobFolder = @()
$CommonName = @()
$PhotoFolders = @()
$WeeklyPhotoPath = "Z:\Project Photos\Photos for d coats project summary reports\" + $(Get-Date $today -format 'yyyyMMdd') + "_WeeklyPhotos"
$dropbox = $env:USERPROFILE+"\Dropbox\"
$zDrive = "Z:\Project Photos\"  

$objExcel = new-object -comobject excel.application  
$objExcel.Visible = $False  
$objWorkbook = $objExcel.Workbooks.Open("Z:\Project Photos\Inspectors.xlsx") 
$dropboxSheet = $objWorkbook.Worksheets.Item(1) 
$photoreportSheet = $objWorkbook.Worksheets.Item(2)
write-host "Checking for new Inspector folders..."

$i = 1 
Do 	
{ 
	$Inspectors += $dropboxSheet.Cells.Item($i, 2).Value()
	$JobFolder += $dropboxSheet.Cells.Item($i, 3).Value()
	$CommonName += $dropboxSheet.Cells.Item($i,4).Value()
	$Engineer += $dropboxSheet.Cells.Item($i, 5).Value()
	$i++ 
} While ($dropboxSheet.Cells.Item($i,2).Value() -ne $null) 
$i = 1 
Do 	
{ 
	$JobsCommon += $photoreportSheet.Cells.Item($i, 1).Value()
	$PhotoFolders += $photoreportSheet.Cells.Item($i, 2).Value()
	$i++ 
} While ($photoreportSheet.Cells.Item($i,2).Value() -ne $null) 

$a = $objExcel.Quit() 
$a = Release-Ref($dropboxSheet)
$a = Release-Ref($photoreportSheet) 
$a = Release-Ref($objWorkbook) 
$a = Release-Ref($objExcel) 
write-host "Inspector and Job Folders Updated"

$i = 0
foreach ($objItem in $Inspectors)
{
	$source = ($dropbox + $Inspectors[$i])
	$target = ($JobFolder[$i])
	$common = ($CommonName[$i])
	$eng = $Engineer[$i]
	$a = MoveFiles $source $target $common $eng
	$i++
}

if ($today.DayOfWeek -eq "Friday")
{
	$sunday = $today.AddDays(-5)
	$i=0
	foreach ($Job in $JobsCommon)
	{
		$RecentFiles = Get-ChildItem $PhotoFolders[$i] -recurse -filter *.jpg | Where-Object {$_.LastWriteTime -gt $sunday}
		
		if ($RecentFiles -ne $null)
		{
			foreach ($file in $RecentFiles)
			{
				If (Test-Path ($WeeklyPhotoPath+"\"+$Job))
				{
					Copy-Item $file.FullName -destination ($WeeklyPhotoPath+"\"+$Job) -force
				}
				Else
				{
					New-Item ($WeeklyPhotoPath+"\"+$Job) -Type Directory
					Copy-Item $file.FullName -destination ($WeeklyPhotoPath+"\"+$Job) -force
				}
			}
		}
		$i++	
	}
	write-host "Photo report compiled..."
	explorer $WeeklyPhotoPath
}

write-host "Photos updated"
	
write-host "Press any key to continue..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")