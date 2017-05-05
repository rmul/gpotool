[cmdletbinding()]
param()

if($host.Runspace.ApartmentState -ne "STA"){
    "Relaunching in STA mode"
    $file = $myInvocation.MyCommand.Path
    powershell -NoProfile -Sta -File $file
    return
}

Clear-Host

#region Global vars
$oldpath=Get-Location
$binpath=Split-Path $MyInvocation.MyCommand.Path
$global:rootpath=Split-Path $binpath
$scriptname=(Get-Item $MyInvocation.InvocationName).BaseName
$global:verb=$VerbosePreference -eq "Continue"
$runtime=$(Get-Date -Format "yyyyMMddHHmmss")
$RunInIse = ($host.Name -eq 'PowerGUIScriptEditorHost') -or ($host.Name -match 'ISE')
#endregion Global vars

#region Load Modules
if ($global:verb) {
Write-Information (“{0} : {1,-20} :{2,0}{3}” –f (Get-Date -Format "HH:mm:ss"),$(Get-PSCallStack)[0].Command," ","Importing ADGPOlib module") -InformationAction Continue
}
try {
    Import-Module $rootpath\lib\ADGPOlib.psm1 -Verbose:$false -WarningAction SilentlyContinue
} catch {
    Write-Error "Error Loading required module ADGPOlib from $rootpath\lib\ADGPOlib.psm1"
    return 1
}
My-Verbose "Importing ActiveDirectory Module" 
Import-Module activedirectory -Verbose:$false
My-Verbose "Importing GroupPolicy Module"
Import-Module grouppolicy -Verbose:$false
My-Verbose "Loading HTMLDiff"
LoadHTMLDiff $rootpath\lib\htmldiff.dll
#endregion Load Modules

#Import-Module SDM-GPMC

Write-Verbose "Merging code from Primalforms generated from .\exportcode.ps1 with custom code from .\guicode.ps1"
$code2replace=@"
#----------------------------------------------
#region Generated Form Code
"@
$replaceby=@"
. $rootpath\lib\guicode.ps1
#----------------------------------------------
#region Generated Form Code
"@
$scriptcode=[io.file]::ReadAllText("$rootpath\lib\exportcode.ps1")
$scriptcode=$scriptcode -replace $code2replace,$replaceby
$scriptblock=$executioncontext.invokecommand.NewScriptBlock($scriptcode)
Write-Verbose "Executing merged code as scriptblock"
&$scriptblock

#region cleanup
Remove-Variable verb -Scope global
Remove-Module ADGPOlib -Verbose:$false
Remove-Module activedirectory -Verbose:$false
Remove-Module grouppolicy -Verbose:$false
#endregion cleanup