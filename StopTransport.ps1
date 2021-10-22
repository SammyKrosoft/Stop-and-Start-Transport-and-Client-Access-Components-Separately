<#
.SYNOPSIS
    This scripts disable Transport component, restarts the Transport service to accelerate the
    Queues drain, and stops the Transport service after 20 seconds.

.PARAMETER ServerName
    Netbios or FQDN name of the server where to stop the Transport Service.

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.EXAMPLE
.\StopTransport.ps1 -ServerName E2016-01
This will stop the Transport component, restart Transport service and stop Transport service 
after 20 seconds on E2016-01


.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $True, Position = 1, ParameterSetName = "NormalRun")][string]$ServerName,
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "0.1"
<# Version changes
v0.1 : first script version
v0.1 -> v0.5 : 
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
$UserDocumentsFolder = "$($env:Userprofile)\Documents"
$OutputReport = "$UserDocumentsFolder\$($ScriptName)_Output_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$UserDocumentsFolder\$($ScriptName)_Logging_$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
Write-Host "Script Log file will be stored as $ScriptLog" -ForegroundColor Green
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function Write-Log
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.
	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.
	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true,position = 0)]
		[string]$Message,
		[Parameter(Mandatory=$false,position = 1)]
        [string]$LogFileName=$ScriptLog,
        [Parameter(Mandatory=$false, position = 2)][switch]$Silent
	)
	
	try
	{
		$DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path $LogFileName
		if (!($Silent)){Write-Host $Message -ForegroundColor Green}
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Write-Log "************************** Script Start **************************"
<# /EXECUTIONS #>

Set-ServerComponentState $ServerName -Component HubTransport -State Draining -Requester Maintenance
Write-Log "Restarting Transport Service to accelerate queues draining..."
Restart-Service MSExchangeTransport
Write-Log "Waiting 20 seconds to drain the queues before stopping Transport Service..."
For ($i=1;$i -lt 20;$i++){
	Write-Host "$i" -ForegroundColor green
	Sleep 1
}
Write-Host "Stopping Transport Service. Any messages left in the queue will be distributed on next server start"
Stop-service MSExchangeTransport -Force

<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
Write-Log "************************** Script End **************************"
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Log $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
