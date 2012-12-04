[cmdletbinding()]
param()
cls
if($host.Runspace.ApartmentState -ne "STA")
{
    "Relaunching in STA mode"
    $file = $myInvocation.MyCommand.Path
    powershell -NoProfile -Sta -File $file
    return
}
function Get-ScriptDirectory {
	[cmdletbinding()]
	param()
	$Invocation = (Get-Variable MyInvocation -Scope 1).Value	
	Split-Path $Invocation.MyCommand.Path
}
Import-Module activedirectory
Import-Module grouppolicy
Import-Module SDM-GPMC

write-verbose "Changing working directory to $(get-scriptdirectory)"
[System.IO.Directory]::SetCurrentDirectory((get-scriptdirectory))
Write-Verbose "Importing functions and global variables from .\ADGPOlib.ps1"
. .\ADGPOlib.ps1
Write-Verbose "Loading HTMLdiff assembly"
LoadHTMLDiff


Write-Verbose "Merging code from Primalforms generated from .\exportcode.ps1 with custom code from .\guicode.ps1"
$code2replace=@"
#----------------------------------------------
#region Generated Form Code
"@
$replaceby=@"
. .\guicode.ps1
#----------------------------------------------
#region Generated Form Code
"@
$scriptcode=[io.file]::ReadAllText(".\exportcode.ps1")
$scriptcode=$scriptcode -replace $code2replace,$replaceby
$scriptblock=$executioncontext.invokecommand.NewScriptBlock($scriptcode)
Write-Verbose "Executing merged code as scriptblock"
&$scriptblock


