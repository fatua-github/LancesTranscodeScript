﻿param(
    [string]$reduce = $False,
    [string]$path = ".\*",
    [string]$codec = "AVC",
    [switch]$turbo = $false,
    [switch]$crf = $false,
    [switch]$copylocal = $false
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

Also requires mkvtools to extract subtitles: https://mkvtoolnix.download/downloads.html#windows 
  Download the portable installaer (32 or 64 bit) and place in tool\mkvtoolnix
#>

<#
ChangeLog
=========
November 15th, 2012 - Initial Creation
Dec 2, 2012 - Fixed copy, Added initial base for HandbrakeCli, Sorted filelist, And created a Mediainfo Template
Dec 9th, 2012 - Fixed error in Multi detection and moving - missed continue statement + reference to HD720 - only supported in newer versions.
Dec 16th, 2012 - Added PCM audio support, and .avs input file support*.
Oct 24th, 2015 - Updated ffmpeg launch to Start-Process  and modifed to use the aac expermintal codec instead of gpl unfriendly faac
Nov 1, 2015 - Forked from transcode.ps1.   
            - Included ability to recognize subtitles and increased quality of encodes and allowing 1080p.  
            - Fixed missing continue on bad file detection.
            - Improved status messages to include file names on same line + Timestamp
Nov 24, 2015- changed to 2-pass instead of CRF 720p and 1080p, fixed int64 vs int on large file sizes, added m2ts.  added lock files to allow multiple computers to share a directory 
            - including invoking from command line.  Verifying that you are not running from folder
Feb 27, 2016 - Included switches -x265 to force x265/aac/mp4 -- using medium preset based on http://www.techspot.com/article/1131-hevc-h256-enconding-playback/page7.html
March 12, 2016 - rebuild switch funcationallty, removed handbrakecli options (never fully implemented and ffmpeg is great)
March 13, 2016 - Switched x265 480p to 2 pass, updated encode logic engine to handle multiple codecs
March 25, 2016 - Adding function to extract MKV Subtitles based on hints from http://www.powershell.amsterdam/2015/06/29/extracting-subtitles-from-mkv-files/ and using mkvtoolsextract
      including tool checking and automatic skipping of extraction if mkvtools is missing
      including filename collision detection in case both internal mkv and external subtitles exist
      only handles srt type
March 14, 2016 - added -copylocal parameter and logic
March 28, 2016 - added -turbo for h264/5 fast 1st pass and updated h265 2nd pass to medium.  
     - added "S_HDMV/PGS" { $subext = "sup" } subtitle extract support
     - added crf option for h265
     - removed "profile" from h265 as it's broken
March 29th, 2016 - Tweaked CRF settings
March 30th, 2016 - Tweaked h265 encode speed, slow just takes too long on my hardware -- eventually i'll create a settings file that users will have the ability to change this
April 16th, 2016 - Increased qualith of video for h265 slightly
May 9th, 2016 - added to the test-path statements to handle square brackets in the filename properly.
May 19th, 2016 - Added HEVC reencode support
May 25th, 2016 - fixed square bracket handling by escapting the filenames instead of using --LiteralPath
June 4th, 2016 - Added creation of log files to "unknown" movements so user can see why the script rejected processing
June 5th, 2016 - Moved Subtitle extractor to after video and audio checks, fixed lockfile and errorlog naming when square brackes are in file name, fixed source video codec array string
Jan 12, 2017 - Copy *.nfo, *.jpg as well
Jan 21st, 2017 - rewrote the order engine so that it encodes the oldest first and repulls folder listing after each encode
Jan 24th, 2018 - fixed "last file endless loop"   + added UHD video support

Note:  
8K 	7,680x4,320
UHD 	3,840x2,160
1080p 	1,920x1,080
720p 	1,280x720 

#>

#Set Priority to Low
$a = gps powershell
$a.PriorityClass = "Idle"

$HomePath = Split-Path -parent $MyInvocation.MyCommand.Definition
$CurrentFolder = pwd



# Start Switch evaluation
$switcherror = 0

#Display current switches
echo "$(get-date) - " 
echo "$(get-date) - " 
echo "$(get-date) - -------------------------------------------------"
echo "$(get-date) - Switches from Command line and internal variables"
echo "$(get-date) - -------------------------------------------------"
echo "$(get-date) - -reduce $reduce"
echo "$(get-date) - -path $path"
echo "$(get-date) - -codec $codec"
echo "$(get-date) - -copylocal $copylocal"
echo "$(get-date) - -crf"
echo "$(get-date) - Working path: $CurrentFolder"
echo "$(get-date) - Script location: $HomePath"
echo " " 
echo " " 

echo "$(get-date) - -----------------------------"
echo "$(get-date) - Testing the provided switches"
echo "$(get-date) - ------------------------------"
echo "$(get-date) - Testing if $path exists" 
if (Test-Path $path) {
    echo "$(get-date) - The path $path exists"
}
else {
 echo "$(get-date) - The path $path does not exist, try again"
 $switcherror = 1
}
echo " "
echo "$(get-date) - Testing chosen codec type"
if ($codec -eq "HEVC" -or $codec -eq "AVC"){
 echo "$(get-date) - $codec chosen"
 }
else{
 echo "$(get-date) - The codec $codec is not supported yet.   To use AVC/aac (H.264), use -codec AVC.  To use HEVC/aac (H.265), use -codec HEVC"
 $switcherror = 1
}

#See if path ends in \* if not, add it
if ($path.substring($path.length - 2) -ne "\*") {
    if ($path.substring($path.length - 1) -eq "*") {
        echo "$(get-date) - Error: Your path ended in * and not \* or just a folder path"
        $switcherror = 1  
    }
    elseif ($path.substring($path.length -1 ) -eq "\") {
        echo "$(get-date) - Note:  Path ended in \,  adding * to the end"
        $path = $path+"*"
    }
    else{
        echo "$(get-date) - Note: Adding \* to the end of your path" 
        $path = $path+"\*"
    }
}

echo " "    
echo "$(get-date) - Testing if reduce is being used"

#All scale options use -2 instead of -1 to ensure final resolution is divisable by 2 as required by x264 and x265
if ($reduce -ne $false) {
 if ($reduce -eq "480p") {
  echo "$(get-date) - $reduce reduction chosen" 
  $ffmpegreduce = "-vf scale=720:-2"
 }
 elseif ($reduce -eq "720p") {
  echo "$(get-date) - $reduce reduction chosen" 
  $ffmpegreduce = "-vf scale=1280:-2"
 }
 elseif ($reduce -eq "1080p") {
  echo "$(get-date) - $reduce reduction chosen" 
  $ffmpegreduce = "-vf scale=1920:-2"
 }
 elseif ($reduce -eq "UHD") {
  echo "$(get-date) - $reduce reduction chosen" 
  $ffmpegreduce = "-vf scale=3840:-2"
 }
 else{
   echo "$(get-date) - $reduce not supported yet. please use 480p, 720p, 1080p, or UHD"
   $switcherror = 1
 }
}
else {
 echo "$(get-date) - -reduce not specified, no special reduction will be used"
 $reduce = "none"
 $ffmpegreduce = ""
}


if ($switcherror -eq 1)  {
 echo "$(get-date) - There are errors in your switches" 
 echo "$(get-date) - Please use the following format:   downloadtranscoder.ps1 -path <path to source files> -codec [AVC|HEVC] (-reduce [480p|720p|1080p]) (-copylocal)"
 echo "$(get-date) -    -path <path to source files>  for example -path c:\filestoencode"
 echo "$(get-date) -    -codec [AVC|HEVC] -- optional with AVC as default -- choose your encoder."
 echo "$(get-date) -    -reduce [480p|720p|1080p|UHD]  -- optional switch to reduce if needed to the chosen size"
 exit
}

    
### END OF SWITCH TESTING
    
### Variables
## Choose if you are using FFMPEG or Handbrake
$encoder="ffmpeg"
$passes = 0

###Tool Directories
$mediainfo = "$HomePath\tools\mediaInfo_cli\MediaInfo.exe"
$mediainfotemplate =  "$HomePath\tools\mediaInfo_cli\Transcode.csv"
$ffmpeg = "$HomePath\tools\ffmpeg\ffmpeg.exe"
$mkvextract = "$HomePath\tools\mkvtoolnix\mkvextract.exe"
$mkvextractexist = 0 #variable to determine if the tool exists
$mkvmerge = "$HomePath\tools\mkvtoolnix\mkvmerge.exe"
$mkvmergeexist = 0 #variable to determine if the tool exists

#Base Encoder variables
$GoodExtensions = "*.divx","*.mov","*.mkv","*.avi","*.mp4","*.m4v","*.mpg","*.ogm","*.mpeg","*.vob","*.avs","*.m2ts","*.wmv"
$SupportedVideoCodecs = "XVID","xvid","avc1","AVC","DX50","DIV3","DivX 4","V_MPEG4/ISO/AVC","MPEG Video","MPEG-4 Visual","Microsoft","VC-1","WMV3","High Efficiency Video Coding","HEVC","V_MPEGH/ISO/HEVC"
$BadVideoCodecs = "DIV3","DivX 4"
$SupportedAudioCodecs = "AAC","AC-3","E-AC-3","MPEG Audio","WMA","DTS","PCM","WMA","Atmos / TrueHD","FLAC"   #Mediainfo Format
$GoodAudio = "AAC"
$SubtitlesExtensions = ".idx",".sub",".srt",".ass",".smi"

# Maximum resolutions for source videos -- maybe switch this to an formula based on source resolution * some factor
if ($codec -eq "AVC") {  #Mediainfo outputs Kbytes*1000 as bytes, not 1024
  $480pVBitRateMax = 899*1000  
  $720pVBitRateMax = 1700*1000
  $1080pVBitRateMax = 2200*1000
  $UHDVBitRateMax = 8000*1000
}
elseif ($codec -eq "HEVC") {
  $480pVBitRateMax = 500*1000
  $720pVBitRateMax = 1000*1000
  $1080pVBitRateMax = 1700*1000 
  $UHDVBitRateMax = 4000*1000
}

#Hardcode Audio codec to AAC, will update later with webm/VP9/OGG
$acodec = "AAC"
if ($acodec -eq "AAC") {
  $2chABitRateMax = 64*1000
  $6chABitRateMax = 256*1000
}

#Encoder Strings
$ffmpegvcopy = "-vcodec copy"
$ffmpeg480p1st = "-vcodec libx264 -profile:v high -level 41 -preset slow -crf 21" # unchanged numbers from 2012  -- No 2nd pass -- Using CRF
$ffmpeg720p2nd = "-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 1503k" # unchanged numbers from 2012
$ffmpeg720p1st = "-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 1503k" # unchanged numbers from 2012
$ffmpeg1080p2nd ="-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 2200k" # unchanged numbers from 2012
$ffmpeg1080p1st ="-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 2200k" # unchanged numbers from 2012
$ffmpegUHD2nd ="-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 7000k" # unchanged numbers from 2012
$ffmpegUHD1st ="-vcodec libx264 -profile:v high -level 41 -preset slow -b:v 7000k" # unchanged numbers from 2012

$ffmpegHEVC_480p2nd = "-vcodec libx265 -preset slow -b:v 303k" # -x265-params `"profile=high10`" # .9 bits per pixel
$ffmpegHEVC_480p1st = "-vcodec libx265 -preset fast -b:v 303k" # -x265-params `"profile=high10`" # .9 bits per pixel
$ffmpegHEVC_720p2nd = "-vcodec libx265 -preset slow -b:v 720k" # -x265-params `"profile=high10`"    #.8 bits per pixel
$ffmpegHEVC_720p1st = "-vcodec libx265 -preset fast -b:v 720k" # -x265-params `"profile=high10`"   #.8 bits per pixel
$ffmpegHEVC_1080p2nd ="-vcodec libx265 -preset slow -b:v 1300k"  # -x265-params `"profile=high10`"   #.64 bits per pixel
$ffmpegHEVC_1080p1st ="-vcodec libx265 -preset fast -b:v 1300k" #  -x265-params `"profile=high10`"   #.64 bits per pixel
$ffmpegHEVC_UHD2nd ="-vcodec libx265 -preset slow -b:v 3500k"  # -x265-params `"profile=high10`"   #.64 bits per pixel
$ffmpegHEVC_UHD1st ="-vcodec libx265 -preset fast -b:v 3500k" #  -x265-params `"profile=high10`"   #.64 bits per pixel

$ffmpegHEVC_480pcrf = "-vcodec libx265 -preset slow -x265-params crf=26" # -x265-params `"profile=high10`"
$ffmpegHEVC_720pcrf = "-vcodec libx265 -preset medium -x265-params crf=24" # -x265-params `"profile=high10`"
$ffmpegHEVC_1080pcrf ="-vcodec libx265 -preset medium -x265-params crf=23" #  -x265-params `"profile=high10`" 
$ffmpegHEVC_UHDcrf ="-vcodec libx265 -preset medium -x265-params crf=21" #  -x265-params `"profile=high10`" 


$ffmpegacopy = "-acodec copy" 
$ffmpeg2ch = "-acodec aac -ac 2 -ab 64k -strict -2"
$ffmpeg6ch = "-acodec aac -ac 6 -ab 192k  -strict -2"
$ffmpegxch = "-acodec aac -ac 6 -ab 192k  -strict -2" #Greater than 6 ch audio downmixed to 6ch

if ( $encoder -eq "ffmpeg" -and $codec -eq "AVC"){
    #echo "ffmepg + AVC chosen"
	$vcopy = $ffmpegvcopy
    $vcopypasses = 1
	$480p1st = $ffmpeg480p1st
    $480ppasses = 1
	$720p1st = $ffmpeg720p1st
	$720p2nd = $ffmpeg720p2nd
    $720ppasses = 2
	$1080p1st = $ffmpeg1080p1st
	$1080p2nd = $ffmpeg1080p2nd
    $1080ppasses = 2
	$UHD1st = $ffmpegUHD1st
	$UHD2nd = $ffmpegUHD2nd
    $UHDpasses = 2
	$acopy = $ffmpegacopy
	$2ch = $ffmpeg2ch
	$6ch = $ffmpeg6ch
	$xch = $ffmpegxch
}
elseif ( $encoder -eq "ffmpeg" -and $codec -eq "HEVC" -and $crf -eq $false){
    #echo "ffmepg + HEVC chosen"
	$vcopy = $ffmpegvcopy
    $vcopypasses = 1
	$480p1st = $ffmpegHEVC_480p1st
	$480p2nd = $ffmpegHEVC_480p2nd
    $480ppasses = 2
	$720p1st = $ffmpegHEVC_720p1st
	$720p2nd = $ffmpegHEVC_720p2nd
    $720ppasses = 2
	$1080p1st = $ffmpegHEVC_1080p1st
	$1080p2nd = $ffmpegHEVC_1080p2nd
    $1080ppasses = 2
	$UHD1st = $ffmpegHEVC_UHD1st
	$UHD2nd = $ffmpegHEVC_UHD2nd
    $UHDpasses = 2
	$acopy = $ffmpegacopy
	$2ch = $ffmpeg2ch
	$6ch = $ffmpeg6ch
	$xch = $ffmpegxch
}
elseif ( $encoder -eq "ffmpeg" -and $codec -eq "HEVC" -and $crf -eq $true){
    #echo "ffmepg + HEVC chosen"
	$vcopy = $ffmpegvcopy
    $vcopypasses = 1
	$480p1st = $ffmpegHEVC_480pcrf
    $480ppasses = 1
	$720p1st = $ffmpegHEVC_720pcrf
    $720ppasses = 1
	$1080p1st = $ffmpegHEVC_1080pcrf
    $1080ppasses = 1
	$UHD1st = $ffmpegHEVC_1080pcrf
    $UHDpasses = 1
	$acopy = $ffmpegacopy
	$2ch = $ffmpeg2ch
	$6ch = $ffmpeg6ch
	$xch = $ffmpegxch
}
else {
	echo "$(get-date) Choose an encoder"
	exit
}


### Working Directories
$SourceDir = $path   #Where the files are to encode
write-host "Sourcedir: $Sourcedir"
$NonWildPath = $path.substring(0,$path.length-2)
$CompleteDir = "$nonwildpath\complete" #Where to put encoded files, and original file
$MultiDir = "$nonwildpath\multi" #where to put files with multiple audio tracks
$BadDir = "$nonwildpath\bad" #where to put files deemed bad (Old codec, very low bitrate)
$ErrorDir = "$nonwildpath\error" #where to put files that something bad happened
$UnknownDir = "$nonwildpath\unknown" #Where to put unknown files
$copylocalDir = "$nonwildpath\copylocal" # where to put files when -copylocal is used

###Clear Variables used throughout Script
$Modifier = ""
$MediainfoOutput = ""
$vcodec = ""
$NumAudioTracks = ""
#$vres = ""
#$vrestmp = ""
$bittmp = ""
$bit = ""
$ffmpegcommand = ""



###########
#FUNCTIONS#
###########

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
		    Move-Item "$NonWildPath\$basename$modifier.mp4" "$CompleteDir\$basename.mp4"
            echo "moved: $NonWildPath\$basename$modifier.mp4 to $CompleteDir\$basename.mp4"
            #Move NFO
            if (test-path "$basename.nfo"){Move-Item "$basename.nfo" "$CompleteDir"} 
            #Move JPG
            if (test-path "$basename.jpg") {Move-Item "$basename.jpg" "$CompleteDir"}
		}
		else
		{
		  	CreateWorkingDir
	       	Move-Item "$filename" "$CompleteDir"
		   	Move-Item "$NonWildPath\$basename.mp4" "$CompleteDir"
		}
    }
    else
    {
    	CreateWorkingDir
    	Move-Item $filename $ErrorDir
    	Move-Item $NonWildPath\$basename$modifier.mp4 $ErrorDir
	} 
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

#Subtitle extract function -- currently only supports .srt from .mkv
#usage:  subextract <fullpathofsource> <name of outpfile without extensions> <type of source container> <destination path>
Function Subextract ($fullpathfile, $videoname, $containertype, $destpath) {
 Write-host "$(Get-date) - Starting subtitle extractor"
 Write-host "$(Get-date) - Video File: $fullpathfile"
 Write-host "$(Get-date) - Container Type: $containertype"

 #if destination path is not there, gracefully extract subtitles to currentdir
 if ($destpath -eq $null) { $destpath = "." }
 Write-host "$(get-date) - Destination path: $destpath"
 
 switch ($containertype)
  {
     ".mkv"  {
        #extension is mkv
        <#
           {
        "codec": "SubRip/SRT",
        "id": 2,
        "properties": {
         "codec_id": "S_TEXT/UTF8",
         "codec_private_length": 0,
         "default_track": false,
         "enabled_track": true,
         "forced_track": false,
         "language": "eng",
         "number": 3,
         "text_subtitles": true,
         "uid": 3112397999577693116
        },
          "type": "subtitles"
      },
     {
       "codec": "SubRip/SRT",
       "id": 3,
       "properties": {
         "codec_id": "S_TEXT/UTF8",
         "codec_private_length": 0,
         "default_track": false,
         "enabled_track": true,
         "forced_track": false,
         "language": "eng",
         "number": 4,
         "text_subtitles": true,
         "track_name": "SDH",
         "uid": 1696238130850447026
       },
       "type": "subtitles"
       #>
     Write-host "$(Get-date) - Starting Mkv Extractor"
     $mkvmergeOutput = &$mkvmerge -i -F json $fullpathfile 2>&1 | ConvertFrom-Json
     #go through all the tracks
     foreach ($track in $mkvmergeOutput.tracks){
        #find the subtitle tracks
        if ($track.type -eq "subtitles") {
            #set the extension of the extracted file based on the type of subtitle stored in the mkv
            switch ($track.codec) {
                "SubRip/SRT" { $subext = "srt" }
                "S_HDMV/PGS" { $subext = "sup" }
            }
            #Set the output file name and include track_name if it exists
            if ($track.properties.track_name -ne $null){
                $outputfile = "$videoname.$($track.properties.track_name).$($track.properties.language).$subext"
            }
            else {
                $outputfile = "$videoname.$($track.properties.language).$subext"
            }

            #check for output file name collisions
            if (test-Path $destpath\$outputfile) {
                #collision detected, renaming
                Move-Item $destpath\$outputfile $destpath\$outputfile-orig
                }


            # Starting of choosing to extract only one langague type, for now, hardcoded to english
            if ($track.properties.language -eq "eng"){
                $command = "$($track.id):$destpath\$outputfile"

                #write-host $command
                # mkvextract tracks source-filename [options] TID1:dest-filename1 [TID2:dest-filename2 ...]
                &$mkvextract tracks `"$fullpathfile`" $command
            }


        }
        else {
          #not a subtitle
        }
     }
     }
  }
}    

#get the next video file
Function getnextfile {
    $files = get-ChildItem -File $SourceDir -Include $GoodExtensions | Sort-Object Creationtime
    write-host FILES found
    write-host $files 
    write-host END FILES found
    Write-host name variable has $file value
    $file = $null 
    Write-host name variable has $file value
    
    foreach ($file in $files) {
        write-host "$(get-date) - Processing $file" 
        #check if lockfile
        $lockfile = $NonWildPath+"\"+$file.basename + "*.lock"
#        write-host "potential lockfile: $lockfile"
	    if (test-path $lockfile)
	    {
         write-host "$(get-date) - Lockfile Found skipping"
         Continue
        }
        else {
            write-host "$(get-date) - No Lockfile found, continuing"
            break
        }
    }
    Write-host name variable has $file value 
    return ($file) 
}        

###############
#END FUNCTIONS#
###############


###################################
#Test that required programs exist#
###################################


write-host "#####################################################################"
Write-Host "$(get-date) - Testing for existance of required tools"
write-host "#####################################################################"
#MkvExtract
if (Test-Path $mkvextract) {
    $mkvtoolnixexist = 1   
    Write-Host "$(get-date) - Found $mkvextract."  -foreground "green"
    Write-Host "$(get-date) - Will extract SRT english subtitles from mkv files" -foreground "green"
}
else {
    $mkvtoolnixexist=0
    Write-Host "$(get-date) - $mkvextract NOT FOUND." -foreground "red"  
    Write-Host "$(get-date) - Will not extract subtitles from mkv files" -foreground "red"
}
#MkvMerge
if (Test-Path $mkvmerge) {
    $mkvmergeexist = 1   
    Write-Host "$(get-date) - Found $mkvmerge."  -foreground "green"
    Write-Host "$(get-date) - Will extract SRT english subtitles from mkv files" -foreground "green"
}
else {
    $mkmergeexist=0
    Write-Host "$(get-date) - $mkvmerge NOT FOUND." -foreground "red"  
    Write-Host "$(get-date) - Will not extract subtitles from mkv files" -foreground "red"
}


##########################################
#  End Testing for existance of programs #
##########################################

###################
#Start Main Script#
###################
#Test if there are files in the folder



#While ($morefiles -eq "yes"){
while (1 -eq 1) {  #infinite loop until break
    write-host getnextfile....
    $file = getnextfile
    write-host file-value is $file

    if ($file.length -eq 0) {
        write-host "No Files found to process, exiting"
        break
    }
   # else {
        #$morefiles = "yes"
   # }

    $hname = hostname
    $filename = $file.FullName

    write-host " " 
    write-host "$(get-date) - Processing $filename"

	$basenameliteral = $file.BaseName
    $lockfileliteral = "$Basenameliteral.$hname.lock"
    $basename = [Management.Automation.WildcardPattern]::Escape($basenameliteral)    #escape the filename to take care of square brackets
    $lockfile = [Management.Automation.WildcardPattern]::Escape($lockfileliteral)    #escape the filename to take care of square brackets
	$extension = $file.Extension
    $filename = [Management.Automation.WildcardPattern]::Escape($filename)    #escape the filename to take care of square brackets

    
	#Check if file is of supported container type.
	if (!($GoodExtensions -contains "*"+"$extension"))
	{
        echo "$(get-date) Unknown Extension: $extension"
		CreateWorkingDir
		Move-Item $filename $UnknownDir
       # Continue
        break
    }

#    echo "File: $file Filename: $filename Basename: $basename Extension: $extension"
    # Lock file down for further processing

    echo "Lockfile: $lockfileliteral"
    New-Item -Path $NonWildPath -Name $lockfileliteral -type file -force	| Out-Null

	#Check if .mp4 file already exists
	if (Test-Path "$NonWildPath\$Basename.mp4") { $modifier="-transcode" }
	else { $modifier="" }

#    echo "Mediainfo path: $mediainfo"
	#Get Mediainfo
$MediainfoOutput = & $mediainfo --Output="file://$mediainfotemplate" "$filename" 2>&1

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
  

  #Example of the hashtable produced by the MediaInfoTemplate
#Name                           Value                                                                                                                                                                                                
#----                           -----                                                                                                                                                                                                
#VBitRate                       8452000                                                                                                                                                                                              
#ABitRate                       107608                                                                                                                                                                                               
#AChannels                      6                                                                                                                                                                                                    
#FileSize                       147399359                                                                                                                                                                                            
#VCodecID                       avc1                                                                                                                                                                                                 
#AudioCount                     1                                                                                                                                                                                                    
#ACodecID                       40                                                                                                                                                                                                   
#VWidth                         1920                                                                                                                                                                                                 
#TextCount                                                                                                                                                                                                                           
#AFormat                        AAC                                                                                                                                                                                                  
#ALangugage                                                                                                                                                                                                                          
#VideoCount                     1                                                                                                                                                                                                    
#Duration                       137252                                                                                                                                                                                               
#VFormat                        AVC     

########################
#   Video Logic Engine #
########################
# Need to make a choice what to do with video.   Data points are: 
#  Test 1 - If current video codec is considered a Bad codec, error out because this will not reencode properly
#  If $codec (user selected codec) is not equal to current video's codec - if not the same as current video, force reencode regardless of bitrate
#  - If the current video codec is unknown, error out because I cant be sure the quality of the output
#  - If $Reduce is set to a resolution lower than the current video, force reencode to the set resolution
#  - If the VBitRateMax is greater than the video bitrate, force reencode
echo "$(get-date) - ==========================="
echo "$(get-date) - Video Logic Engine Starting"
echo "$(get-date) - ==========================="
echo "$(get-date) - User Chosen Codec is $codec"
echo "$(get-date) - Current video Codec is $($MediainfoArray.VFormat)"


#Testing Number of audio tracks

#Note, track 1 will always be video, so 2 tracks = 1 video + 1 audio.  
  # greater than 2 tracks means manual intervention is required
  #Old Method $NumAudioTracks = $MediainfoOutput | Select-String -Pattern "Format/Info" | Measure-Object -Line
  #echo "Number of audio Tracks:" $NumAudioTracks.Lines
	if ( $MediainfoArray["AudioCount"] -ge 2 )
	{ #Manual handling Required
	  	#Check/Make working folders
        write-host  "$(get-date) - Found $($MediainfoArray["AudioCount"]) audio tracks in this video.  It needs to be handled manually and will be moved to the multi folder"
	  	CreateWorkingDir
	  	Move-Item $filename $MultiDir
        Remove-Item -path $(Join-Path $NonWildPath $lockfile)
		Break
	}
#Testing source video codec
	if ($SupportedVideoCodecs -contains $MediainfoArray['VFormat']) 
	{

        write-host "$(get-date) - $($MediainfoArray.VFormat) is an acceptable source codec"
		
		#If no bitrate is shown, estimate and set a value, some source codecs dont supply a bit rate.
		if (!$MediainfoArray["VBitRate"])
		{

			[int64]$AvgBitRate = ([int64]$MediainfoArray["FileSize"]*8)/$MediainfoArray["Duration"]
			
   
			if ($MediainfoArray["AChannels"] -eq 2) { $tmp = 65 }
			else { $tmp = 257 }
			$estVideoBitrate = $AvgBitRate - $tmp
			
			#echo $estVideoBitrate (with conversation from Kb to b, where 1000 = 1k as per mediainfo)
			$MediainfoArray["VBitRate"] = $estVideoBitrate*1000
			$MediainfoArray["ABitrate"] = $tmp*1000
            write-host "$(get-date) - Source codec doesnt supply a Video Bit Rate - using calculated average of $($MediainfoArray.VBitRate)"
            echo "$(get-date) -                              Audio Bit Rate - using calculated average of $($MediainfoArray.ABitRate)"			
		}
        else {
            write-host "$(get-date) - Source codec Video Bit Rate - $($MediainfoArray.VBitRate)"
            write-host "$(get-date) -              Audio Bit Rate - $($MediainfoArray.ABitRate)"		
        }
	}
	elseif ($BadVideoCodecs -contains $MediainfoArray["VFormat"]) #Found a codec deemed bad
	{
		echo "$(get-date) - $($MediainfoArray.VFormat) is considered an unworkable source video codec"
        echo "$(get-date) - $filename will be moved to $BadDir and the next video will be processed"
		CreateWorkingDir
		Move-Item $filename $BadDir
        Remove-Item -path $(Join-Path $NonWildPath $lockfile)
		#Continue   # end this for loop
        break
	}`
	else #Didnt recognize the video codec
	{
		CreateWorkingDir
 		$tmp = "$(get-date) - $($MediainfoArray.VFormat) is an unknown source video codec" 
        write-host $tmp
        $tmp | out-file -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
        $tmp = "$(get-date) - $filename will be moved to $UnknownDir and the next video will be processed"
        write-host $tmp
        $tmp | out-file -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
		Move-Item $filename $UnknownDir
        Remove-Item -path $(Join-Path $NonWildPath $lockfile)
		#Continue
        break
	}


  # Now to determine if transcode of video is required.

  #This Script considers the following resolutions -- Note,  down-scaling, or CRF should produce better quality for the finge cases of in-betwen two resolutions, but for now this works.
#  Name     Height     Width	Width Range
#  480P	    480	       720	    0-1000
#  720P	    720	       1280	    1001-1600
#  1080P	1080	   1920	    1601-2880
#  4K UHD	2160	   3840	    2881-5760
#  8K UHD	4320p	   7680	    5761-

#First if depth is resolution check
# second verification is codec type
# Third verification is codec bitrate

 if ( ([int]$MediainfoArray["VWidth"] -in 0..1000) -or ($reduce -eq "480p") ) {
    echo "$(get-date) - Entering 480p desitnation processing."
       
    if (($reduce -eq "480p") -and ([int]$MediainfoArray.VWidth -ge 1000)) {  # If reduce was specified and source video is larger than this section.. 
        write-host "$(get-date) - Video resolution is larger than 480p Maximum (1000) ($($MediainfoArray.VWidth)), and -reduce specified, will transcode and scale the video"
        $videoopts1st = "$480p1st $ffmpegreduce"
        $videoopts2nd = "$480p2nd $ffmpegreduce"
        $passes = $480ppasses
    }
    else {
        if (($codec -eq $MediainfoArray.VFormat) -and (([int]$MediainfoArray.VBitRate) -le $480pVBitRateMax)) { 
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is less than the maximum ($480pVBitRateMax) and the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs are equal -- Video to be copied"
           $videoopts1st = $vcopy
           $passes = $vcopypasses
        }
        else{
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is greater than the maximum ($480pVBitRateMax) or the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs do not match -- Will Transcode"
           $videoopts1st = $480p1st
           $videoopts2nd = $480p2nd
           $passes = $480ppasses
        }
    }    
 }
 elseif ( ([int]$MediainfoArray["VWidth"] -in 1001..1600) -or ($reduce -eq "720p"))  {
    echo "$(get-date) - Entering 720p desitnation processing."
       
    if (($reduce -eq "720p") -and ([int]$MediainfoArray.VWidth -ge 1600)) {  # If reduce was specified and source video is larger than this section.. 
        write-host "$(get-date) - Video resolution is larger than 720p Maximum (1600) ($($MediainfoArray.VWidth)), and -reduce specified, will transcode and scale the video"
        $videoopts1st = "$720p1st $ffmpegreduce"
        $videoopts2nd = "$720p2nd $ffmpegreduce"
        $passes = $720ppasses
    }
    else {
        if (($codec -eq $MediainfoArray.VFormat) -and (([int]$MediainfoArray.VBitRate) -le $720pVBitRateMax)) { 
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is less than the maximum ($720pVBitRateMax) and the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs are equal -- Video to be copied"
           $videoopts1st = $vcopy
           $passes = $vcopypasses
        }
        else{
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is greater than the maximum ($720pVBitRateMax) or the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs do not match -- Will Transcode"
           $videoopts1st = $720p1st
           $videoopts2nd = $720p2nd
           $passes = $720ppasses
        }
    }    
}
 elseif ( ([int]$MediainfoArray["VWidth"] -in 1601..2880) -or ($reduce -eq "1080p")) {
    echo "$(get-date) - Entering 1080p desitnation processing."
       
    if (($reduce -eq "1080p") -and ([int]$MediainfoArray.VWidth -ge 2880)) {  # If reduce was specified and source video is larger than this section.. 
        write-host "$(get-date) - Video resolution is larger than 1080p Maximum (2880) ($($MediainfoArray.VWidth)), and -reduce specified, will transcode and scale the video"
        $videoopts1st = "$1080p1st $ffmpegreduce"
        $videoopts2nd = "$1080p2nd $ffmpegreduce"
        $passes = $1080ppasses
    }
    else {
        if (($codec -eq $MediainfoArray.VFormat) -and (([int]$MediainfoArray.VBitRate) -le $1080pVBitRateMax)) { 
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is less than the maximum ($1080pVBitRateMax) and the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs are equal -- Video to be copied"
           $videoopts1st = $vcopy
           $passes = $vcopypasses
        }
        else{
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is greater than the maximum ($1080pVBitRateMax) or the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs do not match -- Will Transcode"
           $videoopts1st = $1080p1st
           $videoopts2nd = $1080p2nd
           $passes = $1080ppasses
        }
    }    
}
 elseif ( [int]$MediainfoArray["VWidth"] -in 2881..5760) {
    echo "$(get-date) - Entering UHD desitnation processing."
       
    if (($reduce -eq "UHD") -and ([int]$MediainfoArray.VWidth -ge 2880)) {  # If reduce was specified and source video is larger than this section.. 
        write-host "$(get-date) - Video resolution is larger than 1080p Maximum (2880) ($($MediainfoArray.VWidth)), and -reduce specified, will transcode and scale the video"
        $videoopts1st = "$UHD1st $ffmpegreduce"
        $videoopts2nd = "$UHD2nd $ffmpegreduce"
        $passes = $UHDpasses
    }
    else {
        if (($codec -eq $MediainfoArray.VFormat) -and (([int]$MediainfoArray.VBitRate) -le $UHDVBitRateMax)) { 
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is less than the maximum ($UHDVBitRateMax) and the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs are equal -- Video to be copied"
           $videoopts1st = $vcopy
           $passes = $vcopypasses
        }
        else{
           write-host "$(get-date) - Source Bitrate ($($MediainfoArray.vBitrate)) is greater than the maximum ($UHDVBitRateMax) or the source ($($MediainfoArray.VFormat)) and desination ($codec) codecs do not match -- Will Transcode"
           $videoopts1st = $UHD1st
           $videoopts2nd = $UHD2nd
           $passes = $UHDpasses
        }
    }    
#    $tmp = "$(get-date) - Video has a width of $($MediainfoArray.VWidth) and is considered 4k UHD"
#    write-host $tmp 
#    out-file $tmp -literalpath "$UnknownDir\$Basenameliteral.error.log" -append 
#    $tmp = "$(get-date) - This script is unable to handle 4K UHD video yet, moving $filename to $UnknownDir"
#    write-host $tmp
#    out-file $tmp -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
#    CreateWorkingdir
#    Move-Item $filename $UnknownDir
#    Remove-Item "$NonWildPath\$Basename.$hname.lock"
#    Continue
}
 elseif ( [int]$MediainfoArray["VWidth"] -ge 5761) {
    $tmp = "$(get-date) - Video has a width of $($MediainfoArray.VWidth) and is considered 8k UHD"
    write-host $tmp
    out-file $tmp -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
    $tmp = "$(get-date) - This script is unable to handle 8K UHD video yet, moving $filename to $UnknownDir"
    write-host $tmp
    out-file $tmp -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
    CreateWorkingdir
    Move-Item $filename $UnknownDir
    Remove-Item "$NonWildPath\$Basename.$hname.lock"
    #Continue
    break
}


	#remember abit is audio bitrate from before
	# $achannels = $achannelstmp -replace "\D" , "" #Replace all non-number characters with nothing (only # left)
    echo "$(get-date) - ==========================="
    echo "$(get-date) - Audio Logic Engine Starting"
    echo "$(get-date) - ==========================="
    echo "$(get-date) - Source audio Codec:$($MediainfoArray.AFormat), Channels:$($MediainfoArray.AChannels), Bitrate:$($MediainfoArray.ABitrate)"

	if ($GoodAudio -contains $MediainfoArray["AFormat"])
	{
		if ( [int]$MediainfoArray["AChannels"] -le 2 -and [int]$MediainfoArray["ABitrate"] -le $2chABitRateMax) { 
            write-host "$(get-date) - Will Copy Source audio:$($MediainfoArray.AFormat) is supported, and bitrate ($($MediainfoArray.ABitrate)) is below the maximum ($2chABitRateMax) for the number of channels ($($MediainfoArray.AChannels))"
            $audioopts = $acopy
            }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 2 -and [int]$MediainfoArray["ABitrate"] -gt $2chABitRateMax) { 
            write-host "$(get-date) - Will transcode Source audio:$($MediainfoArray.AFormat) is supported, but bitrate ($($MediainfoArray.ABitrate)) is greater than the maximum ($2chABitRateMax) for the number of channels ($($MediainfoArray.AChannels))"
            $audioopts = $2ch 
            }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 6 -and [int]$MediainfoArray["ABitrate"] -le $6chABitRateMax) {
            write-host "$(get-date) - Will Copy Source audio:$($MediainfoArray.AFormat) is supported, and bitrate ($($MediainfoArray.ABitrate)) is below the maximum ($6chABitRateMax) for the number of channels ($($MediainfoArray.AChannels))"
            $audioopts = $acopy 
            }
		elseif  ( [int]$MediainfoArray["AChannels"] -le 6 -and [int]$MediainfoArray["ABitrate"] -gt $6chABitRateMax){ 
            write-host "$(get-date) - Will transcode Source audio:$($MediainfoArray.AFormat) is supported, but bitrate ($($MediainfoArray.ABitrate)) is greater than the maximum ($6chABitRateMax) for the number of channels ($($MediainfoArray.AChannels))"
            $audioopts = $6ch 
            }
		else { 
            write-host "$(get-date) - Will transcode Source audio:$($MediainfoArray.AFormat) is supported, but bitrate ($($MediainfoArray.ABitrate)) is greater than the maximum ($6chABitRateMax) or the number of channels ($($MediainfoArray.AChannels)) is greater than the maximum (6)"
            $audioopts = $xch 
            }
	}
	elseif ($SupportedAudioCodecs -contains $MediainfoArray["AFormat"])
	{
    #echo $MediainfoArray["AChannels"]
	write-host "$(get-date) - Source Audio codec ($($MediainfoArray.AFormat)) is not supported, will transcode"
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
		$tmp = "$(get-date) Unknown Audio codec type: $($MediainfoArray.AFormat)"
        write-host $tmp
        $tmp | out-file -literalpath "$UnknownDir\$Basenameliteral.error.log" -append
		CreateWorkingDir
		Move-Item $filename $UnknownDir
        Remove-Item -path $(Join-Path $NonWildPath $lockfile)
		#Continue
        break
	}
	# All done finding options

#Subtitle extractor
    Subextract "$NonWildPath\$Basename$extension" "$basename" "$extension" "$CompleteDir"


    echo "$(get-date) - ================="
    echo "$(get-date) - Transcoder Engine"
    echo "$(get-date) - ================="
	echo "$(get-date) - Final Video options $videooptspass1  Audio Options $audioopts"
        if ($passes -eq 1) { 
        	$ffmpegcommand = "-i `"$filename`" $videoopts1st $audioopts `"$NonWildPath\$basename$modifier.mp4`""
	        echo "$(get-date) - Starting Pass 1 of 1"	  
            echo "$(get-date) - FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru
            #PS C:\> start-process calc.exe
            #PS C:\> $p = get-wmiobject Win32_Process -filter "Name='calc.exe'"
            #PS C:\> $p.SetPriority(64)
	        #echo "exit:" $LastExitCode
            movefiles($LastExitCode)
            Remove-Item -path $(Join-Path $NonWildPath $lockfile)
         }
        elseif ($passes -eq 2) {
        	$ffmpegcommand = "-y -i `"$filename`" -pass 1 $videoopts1st $audioopts -f MP4 NUL"
            echo "$(get-date) - Starting Pass 1 of 2"	
            echo "$(get-date) - FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru
	        if ($LASTEXITCODE -ne 0)  # IF FFMPEG had an error, dont continue with pass 2
            {
                echo "FFMPEG HAD ERROR $LastExitCode"
                movefiles ($LastExitcode)
                #Continue
                break
            }
            #PASS 2

            $ffmpegcommand = "-i `"$filename`" -pass 2 $videoopts2nd $audioopts `"$NonWildPath\$basename$modifier.mp4`""
            echo "$(get-date) - Starting Pass 2 of 2"	
	        echo "$(get-date) - FFMPEG Command: $ffmpeg $ffmpegcommand"
            Start-Process -FilePath "$ffmpeg" -ArgumentList "$ffmpegcommand" -Wait -PassThru

            movefiles($LastExitCode)
            Remove-Item -path $(Join-Path $NonWildPath $lockfile)
        }

 #   write-host another round...
 #   #choose next oldest file and restart while
 #   $file = getnextfile
 #   write-host end-file-get $file
#
#    if ($file.length -eq 0) {
#    write-host "No Files found to process, exiting"
#        $morefiles = "no" 
#    }else {
#     $morefiles = "yes" 
#    }

}

Write-host FIN
