# LancesTranscodeScript
Powershell script to kick off analysis and ffmpeg on a folder of videos to transcode.   

I've been using and developing this script for years on  my own and decided that I would share it with the world today

This powershell script currently:
1) reviews a folder for  media
2) finds a media file not yet being transcoded (via a simple lock file check)
3) uses mediainfo to get details about that file and decides if it needs to transcode video, audio or anything
4) kick off ffmpeg with approperiate settings

It is a rigid script that needs some babysitting, but overall it makes the process of trascoding a folder of mixed video of varying quality, resolution, bitrate and audio types and output consistently sized with standard containers.

Eg.     
  Video 1:  1080p video bitrate = 5000 and audio is 2ch aac at 64k
    - will transcode video and copy audio

ToDo
----
 - Make the script less rigid
    - location of ffmpeg and mediainfo
    - ffmpeg encode settings
 - Add Output support for webm(VP9/Ogg) support
 - Add Input support for Webm (vp8/9/ogg) support
 - Add Input support for WMV
 - Ability to detect audio language and drop other langagues
 - Ability to extract specific language subtitles and drop the others
 - adding ability to pull remote files local, transcode locally and push back

Usage:
Note:  To run a powershell script you must perform the following:
Only Once:
1) Download the PowerShellExecutionPolicy.adm from http://go.microsoft.com/fwlink/?LinkId=131786.
2) Install it
3) open gpedit.msc
4) Under computer configuration, right-click Administrative Templates and then click Add/Remove Templates
5) Add PowerShellExecutionPolicy.adm from %programfiles%\Microsoft Group policy
6) Open Administrative Templates\Classic Administrative Templates\Windows Components\Windows PowerShell
7) Enable the property and allow unsigned scripts to be run

