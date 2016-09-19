[CmdletBinding()] 
param([string]$InputFile="C:\SBP\Scripts\domains.xml",[Switch]$Dev)

cls

$runtime=$(Get-Date -Format "yyyyMMddHHmmss")

Write-verbose "$(Get-Date -Format "HH:mm:ss") : Importing ActiveDirectory Module"
Import-Module activedirectory -Verbose:$false
Write-verbose "$(Get-Date -Format "HH:mm:ss") : Importing GroupPolicy Module"
Import-Module grouppolicy -Verbose:$false
$config=LoadConfig $InputFile
$Domains=$config.Domains.Domain

#$domain=$Domains[1]
foreach ($domain in $Domains) {
	$cred=Get-Credential -Message "Your account for $($domain.name)"
	$scriptblock = {
		param($domainname)
		Import-Module activedirectory
		$gpos=get-gpo -domain $domainname -All 
		$gpos |%{
			$_.Displayname
			try {
				$perm=Get-GPPermissions $_.displayname -DomainName $domainname -TargetName "engm\domain admins" -TargetType Group -ErrorAction SilentlyContinue
			} catch {
				$_.Displayname
				#$perm=set-GPPermissions $_.displayname -DomainName $domainname -TargetName "engm\domain admins" -TargetType Group -PermissionLevel GpoEditDeleteModifySecurity
			}
			if (!$perm) {
				"`t$($_.Displayname)"
				$perm=set-GPPermissions $_.displayname -DomainName $domainname -TargetName "engm\domain admins" -TargetType Group -PermissionLevel GpoEditDeleteModifySecurity
			}
		}
	}
	$dc=Get-ADDomainController -DomainName $domain.Name -Discover
	$Session = New-PSSession -ComputerName $dc.Hostname -ConfigurationName microsoft.powershell -Credential $cred 
	Invoke-Command -Session $Session -ScriptBlock $scriptblock -ArgumentList $domain.Name
	"========================="
	Remove-PSSession $Session
	
}