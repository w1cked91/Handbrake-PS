<#  
.SYNOPSIS  
    Powershell Script - Convert file via Handbrake
.DESCRIPTION  
	This script convert file from a directory to the specified container / codecs
.NOTES  
    File Name  : New-HandBrakeConvert_main.ps1  
	Version	   : 1.0
    Author     : Nicolas Giunta
	Email 	   : giunta.nicolas@gmail.com	
	Twitter    : @NicolasGiunta
.LINK
#>

param(
	[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false)] 
	$dir,
	[int] $percentComplete = 0,
	[int] $filesCompleted = 0
)

###################################
#	IMPORTS			
###################################
. ".\New-HandBrakeConvert_var.ps1"
. ".\New-HandBrakeConvert_fonc.ps1" 

####################################################
# Determine if Handbrake is installed and where it is
####################################################
Check-Installation

pushd

Set-Location $dir

##########################
# Execute files convertion
##########################
Execute-Convertion

New-BalloonTip -BalloonTipIcon info -BalloontipText $BalloonTipText -BalloonTipTitle $BalloonTipTitle
# Speak -phrase "Encoding finished."
popd

# clear the variables!
Remove-ScriptVariables($MyInvocation.MyCommand.Name)