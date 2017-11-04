#variables
$serverName      = "\\192.168.1.23"
$uploadFolder    = "\\192.168.1.23\download\"
$handbrakeFolder = "C:\Program Files\HandBrake"
$mp4Extension    = "mp4"
$mkvExtension    = "mkv"
$username        = "HandBrake"
$password        = "C0py2NAS"

function FileToMp4 ($sourceFile, $targetMp4) {
    $cmd           = "$handbrakeFolder\HandBrakeCLI.exe"
    $presetSwitch  = "-Z"
    $presetValue   = "Fast 1080p30"
    #$presetValue   = "HQ 1080p30 Surround"
    $verboseSwitch = "--verbose=1"
    $sourceFile
    & $cmd $presetSwitch $presetValue -i $sourceFile -o $targetMp4 $verboseSwitch
}

function Copy-SourceTimeStampToTarget ($sourceFile, $targetMp4) {
    $srcTime                  = Get-Item  $sourceFile
    $targetTime               = Get-Item $targetMp4
    $targetTime.LastWriteTime = $srcTime.LastWriteTime
    $targetTime.CreationTime  = $srcTime.CreationTime
}

function Get-FileTypeCount ($folder, $extension) {
    $set      = (Get-ChildItem $folder -Filter "*.$extension") | Measure-Object
    $setCount = $set.Count
    "In folder [$folder], there are [$setCount] files of type [$extension]"
    start-sleep 2
}

function Get-FileNameFromFullPath ($file) {
    Split-Path -Path $file -Leaf
}


function Remove-FileNameFromFullPath ($file) {
    [System.IO.Path]::GetDirectoryName($file)
}

function Get-VideoDuration ($fullPath) {
    $LengthColumn = 27
    $objShell     = New-Object -ComObject Shell.Application 
    $objFolder    = $objShell.Namespace($(Remove-FileNameFromFullPath $fullPath))
    $objFile      = $objFolder.ParseName($(Get-FileNameFromFullPath $fullPath))
    $objFolder.GetDetailsOf($objFile, $LengthColumn)
}

function Convert-FromFileToMp4File ($sourceFile, $mp4target) {
    $source
    $target
    start-sleep 2
    $startTime        = Get-Date
    FileToMp4 -sourceFile $source -targetMp4 $target
    $endTime          = Get-Date
    $secondsToConvert = Round-Number(($endTime - $startTime).TotalSeconds)
    Get-Date | Out-File -Append $logFile
    "Completed conversion of [$source] in [$secondsToConvert] seconds" | Out-File -Append $logFile 
    
    $sourceSizeInMb   = Round-Number((Get-Item -Path $source).Length/1MB)
    $targetSizeInMb   = Round-Number((Get-Item -Path $target).Length/1MB)
    "sourceSizeInMb: [$sourceSizeInMb]" | Out-File -Append $logFile 
    "targetSizeInMb: [$targetSizeInMb]" | Out-File -Append $logFile 

    $conversionRate   =  Round-Number($sourceSizeInMb/$secondsToConvert)
    "Conversion rate was [$conversionRate]MB per second" | Out-File -Append $logFile 
    $compressionRatio =  Round-Number($sourceSizeInMb/$targetSizeInMb)
    "Compression ratio was $compressionRatio ([$sourceSizeInMb]/[$targetSizeInMb])" | Out-File -Append $logFile 
    $fileDuration     =  Get-VideoDuration -fullPath $source
    $mp4Duration      =  Get-VideoDuration -fullPath $target
    "File duration: [$fileDuration]" | Out-File -Append $logFile 
    "MP4 duration: [$mp4Duration] (these should match)" | Out-File -Append $logFile 

    Get-Date | Out-File -Append $logFile

    $baseName | clip.exe 
    start-sleep 5
}

function Round-Number ($number) {
    [System.Math]::Round($number,2)
}

function Copy-File ($uncpath, $source, $destination, $extension) {
    net use $uncpath $password /USER:$username
	
	robocopy $source $destination *.$mp4Extension /s /xo /xc /xn

	net use $uncpath /delete 
}

function Convert-FilesToMP4 ($copyFolder, $videoFolder, $extension) {
   $logFile = "$videoFolder/Conversion.$(Get-Random).log"
   Write-Host "Logfile is at [$logfile]"
   
   Set-Location $videoFolder
   
   Copy-File $servername $copyFolder $videoFolder $extension
   
   Get-FileTypeCount -folder $videoFolder -extension $extension | Out-File -Append $logFile 
   Get-FileTypeCount -folder $videoFolder -extension $mp4Extension | Out-File -Append $logFile 

   $FileList = Get-ChildItem -Include "*.$extension" -Recurse
   
   $FileList | ForEach-Object {
       $currentFile  = $_
       $baseName 	 = $currentFile.BaseName
       $source       = "$videoFolder/$baseName.$extension"
       $target       = "$videoFolder/Done/$baseName.$mp4Extension"

       if (!( Test-Path $target)) {
           Convert-FromFileToMp4File $source $target
       }
       Copy-SourceTimeStampToTarget -sourceFile $source -targetMp4 $target
   }
   Copy-File $servername "$videoFolder/Done" $uploadFolder 
}

Convert-FilesToMP4 -c "\\192.168.1.23\video\movie\War for the Planet of the Apes" -v "C:\Users\nicolas.giunta\Desktop\FileToBeConverted" -e $mkvExtension
