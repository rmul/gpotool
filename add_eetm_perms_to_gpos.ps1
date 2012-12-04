[CmdletBinding()] 
param([string]$InputFile="C:\SBP\Scripts\domains.xml",[Switch]$Dev)

cls

$runtime=$(Get-Date -Format "yyyyMMddHHmmss")

Write-verbose "$(Get-Date -Format "HH:mm:ss") : Importing ActiveDirectory Module"
Import-Module activedirectory -Verbose:$false
Write-verbose "$(Get-Date -Format "HH:mm:ss") : Importing GroupPolicy Module"
Import-Module grouppolicy -Verbose:$false
Write-verbose "$(Get-Date -Format "HH:mm:ss") : Importing ADGPOlib.ps1"
. .\ADGPOlib.ps1

LoadHTMLDiff
$config=LoadConfig $InputFile
$Domains=$config.Domains.Domain
$cssstyle=Get-CSS

#$domain=$Domains[1]
foreach ($domain in $Domains) {
	$domain.Name
	$gpos=get-gpo -domain $domain.name -All 
	$gpos |%{
		#$_.Displayname
		try {
			$perm=Get-GPPermissions $_.displayname -DomainName $domain.name -TargetName "eetm\domain admins" -TargetType Group -ErrorAction SilentlyContinue
		} catch {
			$_.Displayname
			#$perm=set-GPPermissions $_.displayname -DomainName $domain.name -TargetName "eetm\domain admins" -TargetType Group -PermissionLevel GpoEditDeleteModifySecurity
		}
		if (!$perm) {
			"`t$($_.Displayname)"
			$perm=set-GPPermissions $_.displayname -DomainName $domain.name -TargetName "eetm\domain admins" -TargetType Group -PermissionLevel GpoEditDeleteModifySecurity
		}
	}
}