cls
$dc=$(Get-ADDomainController -DomainName "eetp.local" -Discover).Name+"."+"eetp.local"

[xml]$migtable=Get-Content "E:\GPOBackup\Migration Tables\EETA_EETP.migtable"
foreach ($mapping in $($migtable.MigrationTable.Mapping | sort Source)) {
	if ($mapping.DestinationNone -ne $null) {
		Write-Host $mapping.Source,"=>","No Mapping"
	} else {
	$class=$null
	if ($mapping.Type -eq "User") {
		$class="user"
	}
	if ((($mapping.Type -eq "GlobalGroup") -or ($mapping.Type -eq "LocalGroup")) -or ($mapping.Type -eq "UniversalGroup")) {
		$class="group"
	}
	if ($class) {
		$dest=$mapping.Source.Replace("eeta","eetp").Replace("EETA","EETP")
		$dest=$dest.Replace("@eetp.local","")
		$dest=$dest.Replace("EETP\","")
		$target=Get-ADObject -Server $dc -filter {(samaccountname -eq $dest) -and (ObjectClass -eq $class)}
		if ($target -eq $null) {
			Write-Host $mapping.Source,"=>",$dest,"NOT FOUND"
		} else {
			#Write-Host $mapping.Source,"found",$target.DistinguishedName
		}
	}
	}
}