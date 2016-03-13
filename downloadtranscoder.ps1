param(
    [string]$reduce = $False,
    [string]$path = ".",
    [string]$codec = "x264"
    )

##########################################################
#                Lances Transcode Script                 #
# https://github.com/fatua-github/LancesTranscodeScript  #
##########################################################

<#  This script will transcode all the files in a directory based up a number of
variables including codec types, frame size, video and audio bitrate, etc.


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

Requires Mediinfo Template:
General;VideoCount=%VideoCount%\r\nAudioCount=%AudioCount%\r\nTextCount=%TextCount%\r\nFileSize=%FileSize%\r\nDuration=%Duration%\r\n
Video;VFormat=%Format%\r\nVCodecID=%Codec ID%\r\nVBitRate=%BitRate%\r\nVWidth=%Width%\r\n
Audio;AFormat=%Format%\r\nACodecID=%Codec ID%\r\nABitRate=%BitRate%\r\nAChannels=%Channels%\r\n
#>

<#
ChangeLog
=========
November 15th, 2012 - Initial Creation
Dec 2, 2012 - Fixed copy, Added initial base for HandbrakeCli, Sorted filelist, And created a Mediainfo Template
Dec 9th, 2012 - Fixed error in Multi detection and moving - missed continue statement + reference to HD720 - only supported in newer versions.
Dec 16th, 2012 - Added PCM audio support, and .avs input file support
Oct 24th, 2015 - Updated ffmpeg launch to Start-Process  and modifed to use the aac expermintal codec instead of gpl unfriendly faac
Nov 1, 2015 - Forked from transcode.ps1.   
            - Included ability to recognize subtitles and increased quality of encodes and allowing 1080p.  
            - Fixed missing continue on bad file detection.
            - Improved status messages to include file names on same line + Timestamp
Nov 24, 2015- changed to 2-pass instead of CRF 720p and 1080p, fixed int64 vs int on large file sizes, added m2ts.  added lock files to allow multiple computers to share a directory 
            - including invoking from command line.  Verifying that you are not running from folder
Feb 27, 2016 - Included switches -x265 to force x265/aac/mp4 -- using medium preset based on http://www.techspot.com/article/1131-hevc-h256-enconding-playback/page7.html
March 12, 2016 - rebuild switch funcationallty, removed handbrakecli options (never fully implemented and ffmpeg is great)
#>

#Set Priority to Low
$a = gps powershell
$a.PriorityClass = "Idle"

$HomePath = Split-Path -parent $MyInvocation.MyCommand.Definition
$CurrentFolder = pwd


# Start Switch evaluation
$switcherror = 0

#Display current switches
echo " " 
echo " " 
echo "-------------------------------------------------"
echo "Switches from Command line and internal variables"
echo "-------------------------------------------------"
echo "-reduce $reduce"
echo "-path $path"
echo "-codec $codec"
echo "-copylocal $copylocal"
echo "Working path: $CurrentFolder"
echo "Script location: $HomePath"
echo " " 
echo " " 

echo "-----------------------------"
echo "Testing the provided switches"
echo "------------------------------"
echo "Testing if $path exists" 
if (Test-Path $path) {
    echo "$path exists"
}
else {
 echo "$path does not exist, try again"
 $switcherror = 1
}
echo " "
echo "Testing chosen codec type"
if ($codec -eq "x265" -or $codec -eq "x264"){
 echo "$codec chosen"
 }
else{
 echo "$codec not supported yet.   To use x264/aac, use -codec x264.  To use x265/aac, use -codec x265"
 $switcherror = 1
}

echo " "    
echo "Testing if reduce is being used"
echo "$reduce $false"

if ($reduce -ne $false) {
 if ($reduce -eq "480p" -or $reduce -eq "720p" -or $reduce -eq "1080p"){
  echo "$reduce reduction chosen" 
 }
 else{
   echo "$reduce not supported yet."
   $switcherror = 1
 }
}
else {
 echo "-reduce not specified, no special reduction will be used"
 $reduce = "none"
}


if ($switcherror -eq 1)  {
 echo "There are errors in your switches" 
 echo "Please use the following format:   downloadtranscoder.ps1 -path <path to source files> -codec [x264|x265] (-reduce [480p|720p|1080p]) (-copylocal)"
 echo "   -path <path to source files>  for example -path c:\filestoencode"
 echo "   -codec [x264|x265] -- optional with x264 as default -- choose your encoder"
 echo "   -reduce [480p|720p|1080p]  -- optional switch to reduce if needed to the chosen size"
}

echo "pausing"
read-host "Press enter"
    
### END OF SWITCH TESTING
    
### Variables
## Choose if you are using FFMPEG or Handbrake
$encoder="ffmpeg"
$passes = 0

###Tool Directories
$mediainfo = "$HomePath\tools\mediaInfo_cli\MediaInfo.exe"
$mediainfotemplate =  "$HomePath\tools\mediaInfo_cli\Transcode.csv"
$ffmpeg = "$HomePath\tools\ffmpeg\ffmpeg.exe"
$Mkvmerge = "C:\Program Files (x86)\MKVToolNix\mkvmerge.exe"

#Base Encoder variables
$GoodExtensions = ".divx",".mov",".mkv",".avi",".AVI",".mp4",".m4v",".mpg",".ogm",".mpeg",".vob",".avs",".m2ts",".wmv"
$SupportedVideoCodecs = "XVID","xvid","avc1","AVC","DX50","DIV3","DivX 4","V_MPEG4/ISO/AVC","MPEG Video","MPEG-4 Visual","Microsoft"
$GoodVideoCodecs = "XVID","xvid","avc1","AVC","DX50","MPEG Video","V_MPEG4/ISO/AVC","MPEG-4 Visual","WVC1"
$BadVideoCodecs = "DIV3","DivX 4"
$SupportedAudioCodecs = "AAC","AC-3","MPEG Audio","WMA","DTS","PCM","WMA"
$GoodAudio = "AAC"
$TranscodeAudio = "AC-3","MPEG Audio","WMA","DTS","PCM"
$SubtitlesExtensions = ".idx",".sub",".srt",".ass",".smi"

#Encoder Strings
$ffmpegvcopy = "-vcodec copy"
#$ffmpeg480p = "-vcodec libx264 -x264opts level=40:b-adapt=1:rc-lookahead=50:ref=5:bframes=16:me=umh:subq=5:deblock=-2,-1:direct=auto -crf 21"
#$ffmpeg720p = "-vcodec libx264 -x264opts level=40:b-adapt=1:rc-lookahead=50:ref=5:bframes=16:me=umh:subq=5:deblock=-2,-1:direct=auto -b:v 1503k"
#$ffmpeg1080p ="-vcodec libx264 -x264opts level=40:b-adapt=1:rc-lookahead=50:ref=5:bframes=16:me=umh:subq=5:deblock=-2,-1:direct=auto -b:v 2200k"  #Dont Reduce to 720p
$ffmpeg480p = "-vcodec libx264 -profile:v high -level 41 -preset slow -crf 21"
$ffmpeg720p = "-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 1503k"
$ffmpeg1080p ="-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 2200k"  #Dont Reduce to 720p
$ffmpegx265_480p = "-vcodec libx265 -preset medium -crf 23 -x265-params `"profile=high10`"" # need to test quality sometime
$ffmpegx265_720p = "-vcodec libx265 -preset medium -b:v 1000k -x265-params `"profile=high10`""
$ffmpegx265_1080p ="-vcodec libx265 -preset medium -b:v 1300k -x265-params `"profile=high10`""

$ffmpegacopy = "-acodec copy" 
$ffmpeg2ch = "-acodec aac -ac 2 -ab 64k -strict -2"
$ffmpeg6ch = "-acodec aac -ac 6 -ab 192k  -strict -2"
$ffmpegxch = "-acodec aac -ac 6 -ab 192k  -strict -2" #Greater than 6 ch audio downmixed to 6ch

if ( $encoder -eq "ffmpeg" -and $codec -eq "x264"){
    echo "ffmepg + x264 chosen"
	$vcopy = $ffmpegvcopy
	$480p = $ffmpeg480p 
	$720p = $ffmpeg720p
	$1080p = $ffmpeg1080p
	$acopy = $ffmpegacopy
	$2ch = $ffmpeg2ch
	$6ch = $ffmpeg6ch
	$xch = $ffmpegxch
}
elseif ( $encoder -eq "ffmpeg" -and $codec -eq "x265"){
    echo "ffmepg + x265 chosen"
	$vcopy = $ffmpegvcopy
	$480p = $ffmpegx265_480p 
	$720p = $ffmpegx265_720p
	$1080p = $ffmpegx265_1080p
	$acopy = $ffmpegacopy
	$2ch = $ffmpeg2ch
	$6ch = $ffmpeg6ch
	$xch = $ffmpegxch
}
else {
	echo "$(get-date -f yyyy-MM-dd) Choose an encoder"
	exit
}

### Working Directories
$SourceDir = $path   #Where the files are to encode
$CompleteDir = "$path\complete" #Where to put encoded files, and original file
$MultiDir = "$path\multi" #where to put files with multiple audio tracks
$BadDir = "$path\bad" #where to put files deemed bad (Old codec, very low bitrate)
$ErrorDir = "$path\error" #where to put files that something bad happened
$UnknownDir = "$path\unknown" #Where to put unknown files

###Clear Variables used throughout Script
$Modifier = ""
$MediainfoOutput = ""
$vcodec = ""
$NumAudioTracks = ""
$vres = ""
$vrestmp = ""
$bittmp = ""
$bit = ""
$vbit = ""
$abit = ""
$acodec = ""
$achannels = ""
$achannelstmp = ""
$ffmpegcommand = ""

# Get list of files in $SourceDir
$files = get-ChildItem $SourceDir | Sort-Object name

#FUNCTIONS

function movefiles ($LastExitCode)
{
    echo "entered movefile fucntion  filename:$filename"
    if ($LASTEXITCODE -eq 0 )
	{
	    if ($modifier)
		{
		    CreateWorkingdir
		    Move-Item "$filename" "$CompleteDir\$basename$extension-orig"
            echo "moved: $filename to $CompleteDir\$basename$extension-orig"
		    Move-Item "$SourceDir\$basename$modifier.mp4" "$CompleteDir\$basename.mp4"
            echo "moved: $SourceDir\$basename$modifier.mp4 to $CompleteDir\$basename.mp4"
		}
		else
		{
		  	CreateWorkingDir
	       	Move-Item "$filename" "$CompleteDir"
		   	Move-Item "$SourceDir\$basename.mp4" "$CompleteDir"
		}
    }
    else
    {
    	CreateWorkingDir
    	Move-Item $filename $ErrorDir
    	Move-Item $SourceDir\$basename$modifier.mp4 $ErrorDir
	} 
    #Remove Lock file
    Remove-Item "$SourceDir\$Basename.$hname.lock"
}
#Test/Create working directories function
Function CreateWorkingDir ()
{
	if (!(test-Path $CompleteDir))
	{
		mkdir $CompleteDir
	}
	if (!(Test-Path $MultiDir))
	{
		mkdir $MultiDir
	}
	if (!(Test-Path $BadDir))
	{
		mkdir $BadDir
	}
	if (!(Test-Path $ErrorDir))
	{
		mkdir $ErrorDir
	}
	if (!(Test-Path $UnknownDir))
	{
		mkdir $UnknownDir
	}
}


#START
foreach ($file in $files) {
    $hname = hostname
	$filename = "$SourceDir\$file"
	$basename = $file.BaseName
	$extension = $file.Extension
	
	# as time passes from first creation of $files, first test if file still exists
    if (!(Test-Path $filename))
     { continue }

    #check if lockfile
	if ("$extension" -eq ".lock")
	{
        echo "$(get-date -f yyyy-MM-dd) Lockfile Found: $extension, skipping"
        Continue
    }

    #Check if is a directory, if so skip
    if (Test-Path -Path $filename -PathType Container )
	{
		echo "$(get-date -f yyyy-MM-dd) $filename Is a directory, skipping"
		Continue
	}
	
    #Check if file is Subtitle.
	if ($SubtitlesExtensions -contains "$extension")
	{
		echo "$(get-date -f yyyy-MM-dd) Subititle file: $filename"
		CreateWorkingDir
		Move-Item $filename $CompleteDir
        Continue
	}

	#Check if file is of supported container type.
	if (!($GoodExtensions -contains "$extension"))
	{
        echo "$(get-date -f yyyy-MM-dd) Unknown Extension: $extension"
		CreateWorkingDir
		Move-Item $filename $UnknownDir
        Continue
    }

	echo "$(get-date -f yyyy-MM-dd) Movie File found: $filename.   Checking for lock file"
    if (Test-Path "$SourceDir\$Basename*.lock")
    {
#        echo "Found lockfile, skipping"
        Continue
    }
    
#    echo "File: $file Filename: $filename Basename: $basename Extension: $extension"
    # Lock file down for further processing
    New-Item "$SourceDir\$Basename.$hname.lock" -type file	| Out-Null

	#Check if .mp4 file already exists
	if (Test-Path "$SourceDir\$Basename.mp4") { $modifier="-transcode" }
	else { $modifier="" }

#    echo "Mediainfo path: $mediainfo"
	#Get Mediainfo
	$MediainfoOutput = &$mediainfo --Output="file://$mediainfotemplate" "$filename" 2>&1
#	echo "Mediainfotemplate: $mediainfotemplate"
#	echo "Mediainfooutput: $MediainfoOutput"

#Get each "variable" into object
	$MediainfoArray = @{}
	switch -regex ($MediainfoOutput) {
    "^\s*([^#].+?)\s*=\s*(.*)" {
      $name,$value = $matches[1..2]
	  #echo $name $value
      $MediainfoArray[$name] = $value.Trim()
   	 }
  	}	
  #$MediainfoArray
 #Note current array entries are: $MediainfoArray["VBitRate"]Name ABitRate AChannels FileSize VCodecID AudioCount ACodecID VWidth TextCount AFormat VideoCount Duration VFormat
   
	#Check if file has multiple audio tracks (Note, track 1 will always be video, so 2 tracks = 1 video + 1 audio.  
	# greater than 2 tracks means manual intervention is required
	#Old Method $NumAudioTracks = $MediainfoOutput | Select-String -Pattern "Format/Info" | Measure-Object -Line
	#echo "Number of audio Tracks:" $NumAudioTracks.Lines
	if ( $MediainfoArray["AudioCount"] -ge 2 )
	{ #Manual handling Required
	  	#Check/Make working folders
	  	CreateWorkingDir
	  	Move-Item $filename $MultiDir
        Remove-Item "$SourceDir\$Basename.$hname.lock"
		Continue
	}
	
	#####Main Processing Starting#####
	#Old Way of doing it: $vcodec = $MediainfoOutput | Select-String -Pattern $SupportedVideoCodecs |  % { $_.Matches } | % { $_.Value } 
	#Old Way of Doing it:$vcodec = $vcodec[1] #Note I am only looking for array entry one in case of multiple video codecs Detected (Eg, AVC and H264 in the same file but as a single track)
	echo $vcodec
	
	### Lets find good video codecs
	if ($GoodVideoCodecs -contains $MediainfoArray["VFormat"]) 
	{
		echo Found good
		#Get Video (and audio) Bitrate
		#Old Way of Doing Things $bittmp = $MediainfoOutput | Select-String -Pattern "Bit rate   " -CaseSensitive -SimpleMatch # find Width
		# $bit = $bittmp -replace "\D" , "" #Replace all non-number characters with nothing (only # left)
		# $vbit = $bit[0]
		# $abit = $bit[1]
		
		#If no bitrate is shown, estimate and set a value
		if (!$MediainfoArray["VBitRate"])
		{
			echo "no VBitRate"
			[int64]$AvgBitRate = ([int64]$MediainfoArray["FileSize"]*8)/$MediainfoArray["Duration"]
			#echo "Total Average Bitrate = $AvgBitRate"
			#acceptable audio bitrate = 64Kbit for 2ch and 256kBit for 6ch
   
			if ($MediainfoArray["AChannels"] -eq 2) { $tmp = 65 }
			else { $tmp = 257 }
			
			$estVideoBitrate = $AvgBitRate - $tmp
			
			#echo $estVideoBitrate
			$MediainfoArray["VBitRate"] = $estVideoBitrate
			$MediainfoArray["ABitrate"] = $tmp
			
		}

		#Get Video Resolution
		# Old method of doing it $vrestmp = $MediainfoOutput | Select-String -Pattern "Width   " -SimpleMatch # find Width
		#$vres = $vrestmp -replace "\D" , "" #Replace all non-number characters with nothing (only # left)
	echo $vres
		
		####### check Video ########
	##	if ($reduce -eq "480")
    ##        { $videoopts = $480p + " -vf scale=-2:480" 
    ##          $passes = 1 }
    ##    elseif ($reduce -eq "720")
    ##        { $videoopts = $720p + " -vf scale=-2:720" 
    ##          $passes = 2 }

    #If video codec is not the codec of the video, force transcode.    Also bitrate is different from video codec or not. 
        if ( [int]$MediainfoArray["VWidth"] -lt 721 -and [int]$MediainfoArray["VBitRate"] -gt 899 ) #480p - Transcode Video
			{ $videoopts = $480p
              $passes = 1 }
		elseif ( [int]$MediainfoArray["VWidth"] -lt 721 -and [int]$MediainfoArray["VBitRate"] -le 899 ) #480p - Copy Video
			{ $videoopts = $vcopy 
              $passes = 1 }  
		elseif ( [int]$MediainfoArray["VWidth"] -lt 1281 -and [int]$MediainfoArray["VBitRate"] -gt 2000 ) #720p - Transcode video
			{ $videoopts = $720p
              $passes = 2}
		elseif ( [int]$MediainfoArray["VWidth"] -lt 1281 -and [int]$MediainfoArray["VBitRate"] -le 2000 ) #720p - Copy video
			{ $videoopts = $vcopy 
              $passes = 1}
		elseif ( [int]$MediainfoArray["VWidth"] -lt 1921 -and [int]$MediainfoArray["VBitRate"] -gt 2000 ) #1080p - Transcode video
			{ $videoopts = $1080p
              $passes = 2 }
		elseif ( [int]$MediainfoArray["VWidth"] -lt 1921 -and [int]$MediainfoArray["VBitRate"] -le 2000 ) #1080p - Copy video
	     	{ $videoopts = $vcopy
              $passes = 1 }
		else ##### Video Resolution didnt Match -- Error #####
		{
			echo "Filename: $filename Good Video - vbitrate: $vbitrate vres: $vres  -- resolution doesnt fit  something has gone wrong"
			CreateWorkingDir
			Move-Item $filename $ErrorDir
            Remove-Item "$SourceDir\$Basename.$hname.lock"
			Continue
		}
	}
	elseif ($BadVideoCodecs -contains $MediainfoArray["VFormat"]) #Found a codec deemed bad
	{
		echo "$(get-date -f yyyy-MM-dd) Unknown video Codec type found in video: $MediainfoArray["VFormat"]"
		CreateWorkingDir
		Move-Item $filename $BadDir
		Continue
	}
	else #Didnt recognize the video codec
	{
		CreateWorkingDir
        echo "Didnt recognize video codec: $vcodec"
		Move-Item $filename $UnknownDir
		Continue
	}
	
	#Lets do audio
	#echo $SupportedAudioCodecs
	#old way of doing it $acodec = $MediainfoOutput | Select-String -Pattern $SupportedAudioCodecs |  % { $_.Matches } | % { $_.Value } 
	#$acodec = $vcodec[1] #Note I am only looking for array entry one in case of multiple video codecs Detected (Eg, AVC and H264 in the same file but as a single track)
	#$acodec = $acodec -replace " " , "" #Remove required leading space to match properly
	
	#remember abit is audio bitrate from before
	# Old way of doing it $achannelstmp = $MediainfoOutput | Select-String -Pattern "Channel(s)   " -SimpleMatch # find Width
	# $achannels = $achannelstmp -replace "\D" , "" #Replace all non-number characters with nothing (only # left)
	echo "$acodec $abit $achannels"
	#AAC Audio
	echo $acodec
	echo TranscodeAudio $TranscodeAudiof
	if ($GoodAudio -contains $MediainfoArray["AFormat"])
	{
	echo "Good audio codec found"
		if ( [int]$MediainfoArray["AChannels"] -le 2 -and [int]$MediainfoArray["ABitrate"] -le 64)
			{ $audioopts = $acopy }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 2 -and [int]$MediainfoArray["ABitrate"] -gt 64)
			{ $audioopts = $2ch }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 6 -and [int]$MediainfoArray["ABitrate"] -le 256)
			{ $audioopts = $acopy }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 6 -and [int]$MediainfoArray["ABitrate"] -gt 256)
			{ $audioopts = $6ch }
		else 
			{ $audioopts = $xch }
	}
	elseif ($TranscodeAudio -contains $MediainfoArray["AFormat"])
	{
    #echo $MediainfoArray["AChannels"]
	#echo "Transcode audio codec found"
        if ( $MediainfoArray["AChannels"] -like '*/*' ) {
            $temp = $MediainfoArray["AChannels"].split("{/}")
            $MediainfoArray["AChannels"] = $temp[0]
        }
		if ( [int]$MediainfoArray["AChannels"] -le "2" )
		{ $audioopts = $2ch }
		elseif ( [int]$MediainfoArray["AChannels"] -le "6" )
		{ $audioopts = $6ch }
		else 
		{ $audioopts = $xch }
		
	}
	else 
	{
		echo "$(get-date -f yyyy-MM-dd) Unknown Audio codec type: $MediainfoArray["AFormat"]"
		CreateWorkingDir
		Move-Item $filename $UnknownDir
		Continue
	}
	# All done finding options
	echo "$(get-date -f yyyy-MM-dd) Modifier $modifier"
	echo "$(get-date -f yyyy-MM-dd) Video options $videooptspass1  Audio Options $audioopts"
	  
        if ($passes -eq 1) { 
        	$ffmpegcommand = "-i `"$filename`" $videoopts $audioopts `"$SourceDir\$basename$modifier.mp4`""
            echo "$(get-date -f yyyy-MM-dd) FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru
            #PS C:\> start-process calc.exe
            #PS C:\> $p = get-wmiobject Win32_Process -filter "Name='calc.exe'"
            #PS C:\> $p.SetPriority(64)
	        #echo "exit:" $LastExitCode
            movefiles($LastExitCode)
         }
        elseif ($passes -eq 2) {
           #PASS 1
        	$ffmpegcommand = "-y -i `"$filename`" -pass 1 $videoopts $audioopts -f MP4 NUL"
            echo "$(get-date -f yyyy-MM-dd) FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru
	        if ($LASTEXITCODE -ne 0)  # IF FFMPEG had an error, dont continue with pass 2
            {
                echo "FFMPEG HAD ERROR $LastExitCode"
                movefiles ($LastExitcode)
                Continue
            }
            #PASS 2
            $ffmpegcommand = "-i `"$filename`" -pass 2 $videoopts $audioopts `"$SourceDir\$basename$modifier.mp4`""
	        echo "$(get-date -f yyyy-MM-dd) FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru
            movefiles($LastExitCode)
        }

}

