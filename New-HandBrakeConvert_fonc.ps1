###################################
#	FUNCTIONS				
###################################

function New-BalloonTip{ # http://powershell.com/cs/blogs/tips/archive/2011/09/27/displaying-balloon-tip.aspx
	[CmdletBinding(SupportsShouldProcess=$true)]
	param(
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No icon specified. Options are None, Info, Warning, and Error!")] 
    $BalloonTipIcon,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No text specified!")] 
    $BalloonTipText,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No title specified!")] 
    $BalloonTipTitle
	)
  [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
  $balloon = New-Object System.Windows.Forms.NotifyIcon
  $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
  $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
  $balloon.Icon = $icon
  $balloon.BalloonTipIcon = $BalloonTipIcon
  $balloon.BalloonTipText = $BalloonTipText
  $balloon.BalloonTipTitle = $BalloonTipTitle
  $balloon.Visible = $true
  $balloon.ShowBalloonTip(10000)
    
  # Icon options are None, Info, Warning, Error
} # end function New-BalloonTip

function Remove-ScriptVariables($path) {  
	$result = Get-Content $path |  
	ForEach { if ( $_ -match '(\$.*?)\s*=') {      
			$matches[1]  | ? { $_ -notlike '*.*' -and $_ -notmatch 'result' -and $_ -notmatch 'env:'}  
		}  
	}  
	ForEach ($v in ($result | Sort-Object | Get-Unique)){		
		Remove-Variable ($v.replace("$","")) -ErrorAction SilentlyContinue
	}
} # end function Get-ScriptVariables

function Speak	{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param(
		[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true)] 
		[string]$phrase
	)
	$voice = New-Object -com SAPI.SpVoice
	$voice.speak($phrase) | Out-Null
} # end function Speak

function Check-Installation	{
	$handbrakeclipath = (Get-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Handbrake.exe" -erroraction silentlycontinue).'(Default)' -replace '.exe','cli.exe'
	if ($handbrakeclipath -eq $null){
		Write-Host "Handbrake not found on this system. Please install Handbrake and try again." -foregroundcolor red;
		# Speak -phrase "Handbrake not found on this system. Please install Handbrake and try again."
		$ie = new-object -comobject "InternetExplorer.Application"
		$ie.visible = $true
		$ie.navigate("http://www.handbrake.fr")
		exit
	}
	# get the shortpath to the cli file so we can assemble the line to execute. This will be removed once I work out a couple more issues
	$a = New-Object -ComObject Scripting.FileSystemObject
	$handbrakeclishortpath = $a.GetFile($handbrakeclipath).ShortPath
	####################################################
	if ($dir -eq $null){
		# if there is no directory specified 
		$Shell = new-object -com Shell.Application
	  $objFolder=$Shell.BrowseForFolder(0, "Choose a folder that contains the files to be converted", 0, 17)
		if ($objFolder -ne $null) {  
		[string] $dir = $objFolder.self.Path
		}  
	}
}

function Execute-Convertion	{
	# test the path and make sure it exists
	$files = (Get-ChildItem "$dir\*" -include *.avi,*.mkv,*.ogm,*.wmv)
	if ($files.length -ge 1){
		ForEach ($file in $files){
			[string] $justName = $file.name.substring(0,$file.name.length-4)
			if (!(Test-Path "$justname.mp4")){
				# $percentComplete = $filesCompleted * (100/$files.length)
				# Write-Progress -Activity "Working..." -PercentComplete $percentComplete -CurrentOperation "$percentComplete% complete" -Status "Please wait."
				Write-Host "`n*******************************************************************************"
				Write-Host "Processing $file" -foreground green
				Write-Host "*******************************************************************************`n"
				[string] $handbrake = $handbrakeclishortpath + " -i `"$file`" -o `"$justName.mp4` $ConvertAttributes
				Invoke-Expression $handbrake
				# $filesCompleted = $filesCompleted++
				New-BalloonTip -BalloonTipIcon info -BalloontipText "$justname.mp4 finished" -BalloonTipTitle $BalloonTipTitle
			}else{
				Write-Host "$justName.mp4 already exists"
			}
		}
	}else{
		# no files found to process
		# either none existed, or there were matching MP4 files for them
		Write-Host "`n*******************************************************************************"
		Write-Host "No files to process" -foreground yellow
		Write-Host "*******************************************************************************`n"
	}
}