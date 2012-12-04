[CmdletBinding()] 
param([string]$InputFile=".\domains.xml",[Switch]$Dev)

cls

$runtime=$(Get-Date -Format "yyyyMMddHHmmss")

try {
	Start-Transcript adpolice.log
} catch {}
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

#region Backup GPOs
Write-Verbose "$(Get-Date -Format "HH:mm:ss") : Region Backup GPOs"
foreach ($domain in $Domains) {
	Backup_GPOs -ConfigDomain $domain
}
#endregion Backup GPOs

#region Get all GPOReports in XML format
Write-verbose "$(Get-Date -Format "HH:mm:ss") : Region Get all GPOReports in XML format"
$i=0
while ($i -lt $Domains.Count) {
	$GPOXMLReports=Get-GPOXMLReports $Domains[$i].Name	
	$Domains[$i]=$Domains[$i] | Add-Member -MemberType NoteProperty -Name GPOXMLReports -Value $GPOXMLReports -PassThru	
	$i++
}
Remove-Variable GPOXMLReports
Remove-Variable i
#endregion Get all GPOReports in XML format

#region Get Different OUs across domains
$domainous=@()
$neutralous=@()
# Get all OUS
foreach ($domain in $config.Domains.Domain) {
	$dc=$(Get-ADDomainController -DomainName $domain.Name -Discover).Name+"."+$domain.Name
	$dom= Get-ADDomain -Server $dc		
	$adous=Get-ADOrganizationalUnit -Server $dc -Filter * | sort @{Expression={$($_.distinguishedName.split(','))[-1]}},@{Expression={$($_.distinguishedName.split(','))[-2]}},@{Expression={$($_.distinguishedName.split(','))[-3]}},@{Expression={$($_.distinguishedName.split(','))[-4]}},@{Expression={$($_.distinguishedName.split(','))[-5]}},@{Expression={$($_.distinguishedName.split(','))[-6]}}
	$allous=@()
	foreach ($ou in $adous) {
		$allous+=$($ou.distinguishedName -replace "$($dom.distinguishedName)","")
		$neutralous+=$($ou.distinguishedName -replace "$($dom.distinguishedName)","")
	}
	$allous=$allous | sort @{Expression={$($_.split(','))[-2]}},@{Expression={$($_.split(','))[-3]}},@{Expression={$($_.split(','))[-4]}},@{Expression={$($_.split(','))[-5]}},@{Expression={$($_.split(','))[-6]}}
	$domain=$domain | Add-Member -MemberType NoteProperty -Name OUs -Value $allous -PassThru
	$domainous+=$domain
}
$neutralous=$neutralous | sort -Unique | sort @{Expression={$($_.split(','))[-2]}},@{Expression={$($_.split(','))[-3]}},@{Expression={$($_.split(','))[-4]}},@{Expression={$($_.split(','))[-5]}},@{Expression={$($_.split(','))[-6]}}
# Compare OUs per domain
foreach ($domain in $domainous) {
	$comparison=Compare-Object $domain.OUs $neutralous -IncludeEqual
	$comparison=$comparison | sort @{Expression={$($_.InputObject.split(','))[-2]}},@{Expression={$($_.InputObject.split(','))[-3]}},@{Expression={$($_.InputObject.split(','))[-4]}},@{Expression={$($_.InputObject.split(','))[-5]}},@{Expression={$($_.InputObject.split(','))[-6]}}	
	$domainous[[array]::IndexOf($domainous,$domain)]=$domainous[[array]::IndexOf($domainous,$domain)] | Add-Member -MemberType NoteProperty -Name ComparedOUs -Value $comparison -PassThru
}
$table = New-Object system.Data.DataTable "Tabel"
$col = New-Object system.Data.DataColumn OU,([string])
$table.columns.add($col)
foreach ($domain in $domainous) {
	$col = New-Object system.Data.DataColumn $($domain.name),([string])
	$table.columns.add($col)
}
$i=0
while ($i -lt $domainous[0].ComparedOUs.count) {
	$row=$table.NewRow()
	$row.OU=$domainous[0].ComparedOUs[$i].InputObject
	foreach ($domain in $domainous) {
		if ($domain.ComparedOUs[$i].SideIndicator -eq "==") {
			$row.$($domain.name)="OK"			
			$dom=[adsi] "LDAP://$($domain.name)"
			$oustring=$domain.ComparedOUs[$i].InputObject+$dom.distinguishedName.ToString()
			$boundou=New-Object system.DirectoryServices.DirectoryEntry("LDAP://$oustring")
			#Get-ADOrganizationalUnit -Identity $oustring -Server eettdc20.eett.local (Problem, the dc is not known here anymore)
			
			$kids=0
			if (!($boundou.Children -eq $null)) {
				foreach ($kid in $boundou.Children) {
					#if (!(($kid.SchemaClassName -eq "organizationalUnit") -or ($kid.SchemaClassName -eq "container"))) {
						$kids++
					#}
				}
			}
			$row.$($domain.name)="$kids"
		} else {
			$row.$($domain.name)="MISSING"
		}
	}
	$table.Rows.Add($row)
	$i++
}
$testhtml=$table |select *  –ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-HTML -head $cssstyle -Title "EET ADPolice OU Report"
$i=0
while ($i -lt $testhtml.Count) {
	if ($testhtml[$i].toupper().contains("MISSING")) {
		#$testhtml[$i]=$testhtml[$i].replace("<tr>","<tr class=""nok"">")
		$testhtml[$i]=$testhtml[$i].replace("<td>MISSING</td>","<td class=""nok"">MISSING</td>")
	}
	$i++
}
$testhtml | Out-File "$($config.Domains.Reports.OUDiffPath)\OUDiff_$runtime.htm"
if ($dev) {
	ii "$($config.Domains.Reports.OUDiffPath)\OUDiff_$runtime.htm"
} else {
	send-SMTPmail -to $config.Mail.Recipient -from $config.Mail.Sender -attachment "$($config.Domains.Reports.OUDiffPath)\OUDiff_$runtime.htm" -subject "EET AD Police - OUs Report" -smtpserver $config.Mail.Server
}
#endregion Get Different OUs across domains

#region Find bogus (empty, disabled and/or unlinked) GPOs
$unlinkedGPOs=@()
$disabledGPOs=@()
$emptyGPOs=@()
foreach ($domain in $Domains) {
	$unlinkedGPOs+=Get-UnlinkedGPOs $domain
	$disabledGPOs+=Get-DisabledGPOs $domain
	$emptyGPOs+=Get-EmptyGPOs $domain
}
$table=BogusGPOstotable $unlinkedGPOs $disabledGPOs $emptyGPOs
$html=$null| ConvertTo-HTML -head $cssstyle -Title "EET ADPolice Obsolete GPOs Report" -Body $table
$html| Out-File "$($config.Domains.Reports.ObsoletedGPOsPath)\ObsoletedGPOs_$runtime.htm"
if ($dev) {
	ii "$($config.Domains.Reports.ObsoletedGPOsPath)\ObsoletedGPOs_$runtime.htm"
} else {
	send-SMTPmail -to $config.Mail.Recipient -from $config.Mail.Sender -attachment "$($config.Domains.Reports.ObsoletedGPOsPath)\ObsoletedGPOs_$runtime.htm" -subject "EET AD Police - Obsolete GPOs Report" -smtpserver $config.Mail.Server
}
#endregion Find bogus (empty, disabled and/or unlinked) GPOs

#region Get GPOs across domains
Write-verbose "$(Get-Date -Format "HH:mm:ss") : Region Get GPOs across domains"
$neutralGPOs=@()
# Get all GPOs
foreach ($domain in $Domains) {
	foreach ($gpo in $domain.GPOXMLReports) {
		$neutralGPOs+=$($gpo.Name -replace "^$($domain.GPOPrefix)","?_")
	}
	Remove-Variable gpo
}
Remove-Variable domain
$neutralGPOs=$neutralGPOs | sort
$neutralGPOs=$neutralGPOs | sort -Unique
## Compare OUs per domain
foreach ($domain in $Domains) {
	$comparison=Compare-Object $($domain.GPOXMLReports |foreach {$_.Name -replace "^$($domain.GPOPrefix)","?_"} | sort) $neutralGPOs -IncludeEqual
	$comparison=$comparison | sort InputObject
	$domains[[array]::IndexOf($domains,$domain)]=$domains[[array]::IndexOf($domains,$domain)] | Add-Member -MemberType NoteProperty -Name ComparedGPOs -Value $comparison -PassThru
}
#$comparison | ft -AutoSize
$table = New-Object system.Data.DataTable "Tabel"
$col = New-Object system.Data.DataColumn GPO,([string])
$table.columns.add($col)
foreach ($domain in $domains) {
	$col = New-Object system.Data.DataColumn $($domain.name),([string])
	$table.columns.add($col)
}
$i=0
while ($i -lt $domains[0].ComparedGPOs.count) {
	$row=$table.NewRow()
	$row.GPO=$domains[0].ComparedGPOs[$i].InputObject
	foreach ($domain in $domains) {
		if ($domain.ComparedGPOs[$i].SideIndicator -eq "==") {
			$row.$($domain.name)="OK"			
		} else {
			$row.$($domain.name)="MISSING"
		}
	}
	$table.Rows.Add($row)
	$i++
}
$testhtml=$table |select *  –ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-HTML -head $cssstyle -Title "EET ADPolice GPO Report"
$i=0
while ($i -lt $testhtml.Count) {
	if ($testhtml[$i].toupper().contains("MISSING")) {
		#$testhtml[$i]=$testhtml[$i].replace("<tr>","<tr class=""nok"">")
		$testhtml[$i]=$testhtml[$i].replace("<td>MISSING</td>","<td class=""nok"">MISSING</td>")
	}
	$i++
}
$testhtml | Out-File "$($config.Domains.Reports.GPDiffPath)\GPDiff_$runtime.htm"
if ($dev) {
	ii "$($config.Domains.Reports.GPDiffPath)\GPDiff_$runtime.htm"
	#send-SMTPmail -to $config.Mail.Recipient -from $config.Mail.Sender -attachment "$($config.Domains.Reports.GPDiffPath)\GPDiff_$runtime.htm" -subject "EET AD Police - GPOs Report" -smtpserver $config.Mail.Server
} else {
	send-SMTPmail -to $config.Mail.Recipient -from $config.Mail.Sender -attachment "$($config.Domains.Reports.GPDiffPath)\GPDiff_$runtime.htm" -subject "EET AD Police - GPOs Report" -smtpserver $config.Mail.Server
}
#endregion

#region Get GPOLinks per Domain
foreach ($domain in $Domains) { 
	get-OU-Report $domain
}
#endregion

try {
	Stop-Transcript
} catch {}