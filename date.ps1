#######################
# Lance's Movie Dater #
#######################

<#  This script will Look at a folder of movies with an imbedded date in the name in the format (####) and 
  then set the file date to today's month and day + file year.   This is useful for movie indexing programs that have a "new" catagory and when 
  you replace old movies with better versions, and dont want it to appear in new.


Note:  To run a powershell script you must perform the following:
Only Once:
1) Download the PowerShellExecutionPolicy.adm from http://go.microsoft.com/fwlink/?LinkId=131786.
2) Install it
3) open gpedit.msc
4) Under computer configuration, right-click Administrative Templates and then click Add/Remove Templates
5) Add PowerShellExecutionPolicy.adm from %programfiles%\Microsoft Group policy
6) Open Administrative Templates\Classic Administrative Templates\Windows Components\Windows PowerShell
7) Enable the property and allow unsigned scripts to be run

each time:
1) open the power shell prompt
2) from the directory of video files, run the script
#>

<#
ChangeLog
=========
Dec 16th, 2012 - initial creation

#>

### Variables
$HomePath = $MyInvocation.Line | Split-Path

### Working Directories
$SourceDir = "."   #Where the files are to encode

# Get list of files in $SourceDir
$files = get-ChildItem $SourceDir | Sort-Object name

foreach ($file in $files) {
    
	$filename = $file.Name
	$basename = $file.BaseName
	$extension = $file.Extension
	
	#echo "Filename: $filename Basename: $basename Extension: $extension"
	#Check if is a directory, if so skip
	   if (Test-Path -Path $file.FullName -PathType Container )
	{
		#echo $filename "Is a directory, skipping"
		Continue
	}
	
	#echo $filename
	$year = $filename | Select-string "\((?<num>\d{0,9})\)" | %{$_.matches[0].Groups['num'].value}
	#echo $date
	
	$a = Get-Date
	#Make sure date is reasonable (eg, not 1080, or 720, or some stupid date, else set this year
	if (!(([int]$year -gt 1900) -and ([int]$year -lt 2020))){
		echo "bad date $filename $year"
		$year = $a.year
		continue
	}
		$month = $a.Month
	$day = $a.Day
	$time = $a.ToShorttimeString()
	#echo "Year: $year Month: $month Day: $day Time: $time"
	
	#Change date of the file
	$file.LastWriteTime = New-Object DateTime $year,$month,$day
}