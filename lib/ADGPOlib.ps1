[cmdletbinding()]
param()

#region Functions
Write-Verbose (“{0} : {1,-20} :{2,4}{3}” –f (Get-Date -Format "HH:mm:ss"),$(Get-PSCallStack)[0].Command," ","Defining Functions")

#Write-Verbose (“{0} : {1,-20} :{2,8}{3}” –f (Get-Date -Format "HH:mm:ss"),$(Get-PSCallStack)[0].Command," ","My-Verbose - Formatted Verbose Logging")
Function My-Verbose {
	[cmdletbinding()]
	param ([string]$message,[int]$indent=0)
    $callstack=Get-PSCallStack
    $spacing=4 * $($callstack.count - (3 - $indent))
    Write-Verbose (“{0} : {1,-20} :{2,$spacing}{3}” –f (Get-Date -Format "HH:mm:ss"),$callstack[1].Command," ",$message)

}

#My-Verbose "My-Error - Formatted Error Logging" 1
Function My-Error {
	[cmdletbinding()]
	param ([string]$message,[int]$indent=0)
    $callstack=Get-PSCallStack
    $spacing=4 * $($callstack.count - (3 - $indent))
    Write-Error (“{0} : {1,-25} :{2,$spacing}{3}” –f (Get-Date -Format "HH:mm:ss"),$callstack[1].Command," ",$message)
}

#My-Verbose "LoadHTMLDiff - Loads htmldiff.dll from current directory" 1
function LoadHTMLDiff {
	[cmdletbinding()]
	param($file='.\HtmlDiff.dll')

	try {
        My-Verbose "Loading: $file"
		$rc=[Reflection.Assembly]::LoadFrom($file)
		My-Verbose "Loaded: $rc"
	} catch {
		My-Error "$file could not be loaded!"
	}
}

#My-Verbose "LoadConfig - Loads specified xmlfile and returns configuration-section as xml-object" 1
Function LoadConfig {
	[cmdletbinding()]
	param ($filename)
	My-Verbose "Loading $filename"
	try {
		[xml]$config=Get-Content $filename
		My-Verbose "$(Get-Date -Format "HH:mm:ss") : Loaded config from $filename"
	} catch {
		Write-Error "Error loading config $filename"
	}
	$config.Configuration
}

#My-Verbose "CreateConfig - Scans for Trusted Domains and returns configuration-section for these domains as xml-object" 1
Function CreateConfig {
	[cmdletbinding()]
	param ($filename)
	$xmldoc=new-object System.Xml.XmlDocument
	$xmlconfig=$xmldoc.CreateElement("Configuration")
	$xmldomains=$xmldoc.CreateElement("Domains")
	
    $domain= Get-ADDomain
    $xmldomain=$xmldoc.CreateElement("Domain")		
	$xmldomain.SetAttribute("Name",$domain.DNSRoot)
	$xmldomain.SetAttribute("GPOPrefix",$domain.Name.Substring(3,1)+'_')
	$xmldomain.SetAttribute("ShortName",$domain.Name.Split('.')[0])
	$xmldomain.SetAttribute("GPOBackupPath","c:\GPOBackup\$($domain.Name)\Backups" )
	$xmldomain.SetAttribute("GPOReportPath","C:\GPOBackup\$($domain.Name)\Reports")
	$xmldomain.SetAttribute("GPOLinkReportPath","C:\GPOBackup\$($domain.Name)\GPOLinkReports")
	$xmldomain.SetAttribute("OUReportPath","C:\GPOBackup\$($domain.Name)\OUReports")
	$xmldomains.AppendChild($xmldomain)
	Remove-Variable xmldomain

	$trusteddomains=Get-ADTrust -Filter *
	foreach ($trusteddomain in $trusteddomains) {
		$xmldomain=$xmldoc.CreateElement("Domain")		
		$xmldomain.SetAttribute("Name",$trusteddomain.Name)
		$xmldomain.SetAttribute("GPOPrefix",$trusteddomain.Name.Substring(3,1)+'_')
		$xmldomain.SetAttribute("ShortName",$trusteddomain.Name.Split('.')[0])
		$xmldomain.SetAttribute("GPOBackupPath","c:\GPOBackup\$($trusteddomain.Name)\Backups" )
		$xmldomain.SetAttribute("GPOReportPath","C:\GPOBackup\$($trusteddomain.Name)\Reports")
		$xmldomain.SetAttribute("GPOLinkReportPath","C:\GPOBackup\$($trusteddomain.Name)\GPOLinkReports")
		$xmldomain.SetAttribute("OUReportPath","C:\GPOBackup\$($trusteddomain.Name)\OUReports")
		$xmldomains.AppendChild($xmldomain)
		Remove-Variable xmldomain
	}
	$xmlReports=$xmldoc.CreateElement("Reports")
	$xmlReports.SetAttribute("OUDiffPath","C:\GPOBackup\OUDiff")
	$xmlReports.SetAttribute("GPDiffPath","C:\GPOBackup\GPDiff")
	$xmlReports.SetAttribute("ObsoletedGPOsPath","C:\GPOBackup\Obsoleted GPOs")
	$xmlReports.SetAttribute("GPLinkDiffPath","C:\GPOBackup\GPLinkDiff")
	$xmldomains.AppendChild($xmlReports)
	$xmlconfig.AppendChild($xmldomains)
	
	$xmlmail=$xmldoc.CreateElement("Mail")
	$xmlmail.SetAttribute("SmtpHost","mail")
	$xmlmail.SetAttribute("recipient","rmul@schubergphilis.com")
	$xmlmail.SetAttribute("sender","ENG_AD_Police_DEV@engmdc11.engm.local")
	$xmlGPOChangeReport=$xmldoc.CreateElement("GPOChangeReport")
	$xmlGPOChangeReport.SetAttribute("Recipient","rmul@schubergphilis.com")
	$xmlGPOChangeReport.SetAttribute("Sender","ENG_AD_Police_DEV@engmdc11.engm.local")
	$xmlGPOChangeReport.SetAttribute("Subject","ENG GPO Change Report")
	$xmlmail.AppendChild($xmlGPOChangeReport)
	$xmlOUChangedReport=$xmldoc.CreateElement("OUChangedReport")
	$xmlOUChangedReport.SetAttribute("Recipient","rmul@schubergphilis.com")
	$xmlOUChangedReport.SetAttribute("Sender","ENG_AD_Police_DEV@engmdc11.engm.local")
	$xmlOUChangedReport.SetAttribute("Subject","ENG OU Change Report")
	$xmlmail.AppendChild($xmlOUChangedReport)  
	$xmlconfig.AppendChild($xmlmail)
	
	$xmlGPOBackup=$xmldoc.CreateElement("GPOBackup")
	$xmlGPOBackup.SetAttribute("Recipient","rmul@schubergphilis.com")
	$xmlGPOBackup.SetAttribute("Sender","ENG_AD_Police_DEV@engmdc11.engm.local")
	$xmlGPOBackup.SetAttribute("Subject","GPO Change Report")
	$xmlGPOBackup.SetAttribute("SaveDiffReports","true")
	$xmlGPOBackup.SetAttribute("AttachDiffReports","true")
	$xmlGPOBackup.SetAttribute("SendResult","true")
	$xmlconfig.AppendChild($xmlGPOBackup)  

	$xmlGPOLinkReport=$xmldoc.CreateElement("GPOLinkReport")
	$xmlGPOLinkReport.SetAttribute("Recipient","rmul@schubergphilis.com")
	$xmlGPOLinkReport.SetAttribute("Sender","ENG_AD_Police_DEV@engmdc11.engm.local")
	$xmlGPOLinkReport.SetAttribute("Subject","GPOLinks Change Report")
	$xmlGPOLinkReport.SetAttribute("SaveDiffReports","true")
	$xmlGPOLinkReport.SetAttribute("AttachDiffReports","true")
	$xmlGPOLinkReport.SetAttribute("SendResult","true")
	$xmlconfig.AppendChild($xmlGPOLinkReport)  

	$xmlconfig
}

Function Send-SMTPmail($to, $from, $subject, $body, $attachment, $cc, $bcc, $port, $timeout, $smtpserver, [switch] $html, [switch] $alert) {
    if ($smtpserver -eq $null) {$smtpserver = "mail"}
    $mailer = new-object Net.Mail.SMTPclient($smtpserver)
    if ($port -ne $null) {$mailer.port = $port}
    if ($timeout -ne $null) {$mailer.timeout = $timeout}
    $msg = new-object Net.Mail.MailMessage($from,$to,$subject,$body)
    if ($html) {$msg.IsBodyHTML = $true}
    if ($cc -ne $null) {$msg.cc.add($cc)}
    if ($bcc -ne $null) {$msg.bcc.add($bcc)}
    if ($alert) {$msg.Headers.Add("message-id", "<3bd50098e401463aa228377848493927-1>")}
    if ($attachment -ne $null) {
        $attachment = new-object Net.Mail.Attachment($attachment)
        $msg.attachments.add($attachment)
    }

    $mailer.send($msg)
} 

Function Get-OU-Report {
	[cmdletbinding()]
	param($ConfigDomain,$Config)
	$changes=$false
	if ($config.GPOLinkReport.SendResult -eq "true") {
		if ($config.Mail.SmtpHost -eq $null) {$smtpserver = "mail"} else {$smtpserver = $config.Mail.SmtpHost}
    	$mailer = new-object Net.Mail.SMTPclient($smtpserver)
    	$msg = new-object Net.Mail.MailMessage($config.GPOLinkReport.Sender,$config.GPOLinkReport.Recipient)
	}
	$myruntime=$(Get-Date -Format "yyyyMMddHHmmss")
	
	New-Item -ItemType Directory -Force -Path $ConfigDomain.GPOLinkReportPath
	$dc=$(Get-ADDomainController -DomainName $ConfigDomain.Name -Discover).Name+"."+$ConfigDomain.Name
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Found DC $dc"
	$DomainDN="DC="+$ConfigDomain.Name.split('.')[0]+",DC="+$ConfigDomain.Name.split('.')[1]
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting GPOLinks for $DomainDN"
	[Array]$gpinheritance=get-gpinheritance -target $DomainDN -Domain $ConfigDomain.Name
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting all OUs"
	$adous=Get-ADOrganizationalUnit -Server $dc -Filter * | sort @{Expression={$($_.distinguishedName.split(','))[-1]}},@{Expression={$($_.distinguishedName.split(','))[-2]}},@{Expression={$($_.distinguishedName.split(','))[-3]}},@{Expression={$($_.distinguishedName.split(','))[-4]}},@{Expression={$($_.distinguishedName.split(','))[-5]}},@{Expression={$($_.distinguishedName.split(','))[-6]}}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting GPOLinks for all $($adous.count) OUs"
	$gpinheritance+=$adous|%{(Get-GPInheritance -Target $_ -Domain $ConfigDomain.Name)}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Searching most recent GPOLinkreport in $($ConfigDomain.GPOLinkReportPath)"
	[array]$oldreports=get-childitem "$($ConfigDomain.GPOLinkReportPath)\*" -Include "GpoLinkReport_*.html" | sort CreationTime
	if ($oldreports.count -eq 0) {
		[array]$oldreports=get-childitem "..\var\*" -Include "Empty.html"
	}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting content from $($oldreports[-1])"
	$oldreport=Get-Content $oldreports[-1]

	$report=OutputOUGPOLinkHtmlHeader $ConfigDomain
	$report+=OutputOUGPOLinkTableHeader
	$report+=$gpinheritance | foreach {OutputGPOLinkOU $_}
	$report+=OutputGPOLinkTableFooter
	$report+=OutputGPOLinkHtmlFooter
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Comparing $($oldreports[-1]) with current report"
	$compare=Compare-Object $oldreport $report
	if ($compare) {
		My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.name): GPO Links changed, Saving current report to $($ConfigDomain.GPOLinkReportPath)\GpoLinkReport_$myruntime.html"		
		$report | Out-File "$($ConfigDomain.GPOLinkReportPath)\GpoLinkReport_$myruntime.html"
		if ($config.GPOLinkReport.SaveDiffReports -eq "true") {
			My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.name): Creating Diff report $($ConfigDomain.GPOLinkReportPath)\GpoLinkReportDiff_$myruntime.html"
			[Helpers.HtmlDiff] $diff=new-object helpers.HtmlDiff($oldreport, $report)
			$html=$diff.Build()
			$style="<style type=""text/css"">"+$(Get-Content "..\var\GpoLinkReport.css")+"</style>"
			$html=$html.Replace("<link rel=""stylesheet"" type=""text/css"" href=""GpoLinkReport.css"" />",$style)
			$html=$html.Replace("</h1>","</h1><h2>Generated on $now</h2>")
			$html | Out-File "$($ConfigDomain.GPOLinkReportPath)\GpoLinkReportDiff_$myruntime.html"
			if ($Dev) {
				ii "$($ConfigDomain.GPOLinkReportPath)\GpoLinkReport_$myruntime.html"
				ii "$($ConfigDomain.GPOLinkReportPath)\GpoLinkReportDiff_$myruntime.html"
				#if ($ReportRecipient) {send-SMTPmail -to $notify -from $notifier -attachment "$($ConfigDomain.GPOLinkReportPath)\GpoLinkReportDiff_$myruntime.html" -subject "$($ConfigDomain.Name) GPOLink Change Report - $now" -smtpserver $smtpserver -html -body $html}
			}
			if ($config.GPOLinkReport.AttachDiffReports -eq "true") {
				$attachment = new-object Net.Mail.Attachment("$($ConfigDomain.GPOLinkReportPath)\GpoLinkReportDiff_$myruntime.html")
        		$msg.attachments.add($attachment)
			}
		}
		if ($config.GPOLinkReport.SendResult -eq "true") {
			$attachment = new-object Net.Mail.Attachment("$($ConfigDomain.GPOLinkReportPath)\GpoLinkReport_$myruntime.html")
        	$msg.attachments.add($attachment)
			$msg.body="GPO Links in domain $($configdomain.Name) have changed."
			$msg.Subject="$($configdomain.Name) - $($config.GPOLinkReport.Subject)"
    		$mailer.send($msg)
		}
	} else {
		My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.name): GPO Links Not changed since last report $($oldreports[-1].CreationTime)"
	}
	$gpinheritance
}

Function Get-PreviousReport($ConfigDomain=$(throw "ConfigDomain required."),$gponame,$currentreport){
	$gporeports=Get-ChildItem -Path "$($ConfigDomain.GPOReportPath)\*" -Include "$($gponame)_[0-9]*.html"
	$tempje=$currentreport.split('_')
	$currentreportdate=[DateTime]::ParseExact($tempje[$tempje.Count-2],"M-d-yyyy",[System.Globalization.CultureInfo]::InvariantCulture)
	$previousreport="..\var\Empty.html"
	$datediff=New-TimeSpan
	if ($gporeports) {
		foreach ($report in $gporeports) {
			$reportdate=[DateTime]::ParseExact($report.name.split('_')[$report.name.split('_').Count-2],"M-d-yyyy",[System.Globalization.CultureInfo]::InvariantCulture)
			$thisdiff=New-TimeSpan -Start $reportdate -End $currentreportdate
			if ($thisdiff.days -gt 0) {
				if ($datediff.days.equals(0)) {
					$datediff=$thisdiff
					$previousreport=$report.Fullname			
				} else {
					if ($thisdiff.days -lt $datediff.days) {
						$datediff=$thisdiff
						$previousreport=$report.Fullname
					}
				}
			}
		}
	}
	return $previousreport
}

Function Add-Node { 
	[cmdletbinding()]
    param (
		$selectedNode, 
        $name, 
        $tag 
    ) 
    $newNode = new-object System.Windows.Forms.TreeNode  
	$newNode.ImageKey="Blank"
    $newNode.Name = $name 
    $newNode.Text = $name 
    $newNode.Tag = $tag 
    $selectedNode.Nodes.Add($newNode) | Out-Null 
    return $newNode 
} 

Function ou-treenodes {
	[cmdletbinding()]
	param ($node,$dn,$svr)	
	My-Verbose "$(Get-Date -Format "HH:mm:ss") : `tOU-Treenodes for $dn"
	Get-ADOrganizationalUnit -Server $dc -Filter * -SearchScope 1 -SearchBase $dn | %{
		$_.distinguishedname
		$newnode=add-node $node $_.Name $_.Name
		ou-treenodes $newnode $_.distinguishedname $svr | Out-Null
	}
}

Function Compare-OU-Trees($domains2compare){
	My-Verbose "$(Get-Date -Format "HH:mm:ss") : Comparing $domains2compare"
	$domainous=New-Object System.Windows.Forms.TreeNode
	$domainous.Name = "Compared" 
    $domainous.Text = "All Domains" 
    $domainous.Tag = "Compared" 
	$domainous.ImageKey = "Blank"
	
	foreach ($domain in $domains2compare) {
		My-Verbose "$(Get-Date -Format "HH:mm:ss") : `tAdding $($domain.Name) to tree"
		$ous=New-Object System.Windows.Forms.TreeNode
		$dc=$(Get-ADDomainController -DomainName $domain.Name -Discover).Name+"."+$domain.Name
		$dom= Get-ADDomain -Server $dc		
		$domainnode=add-node $domainous $domain.Name $domain.Name
		$domainnode.imagekey="Main"
		ou-treenodes -node $domainnode -dn $dom.DistinguishedName -svr $dc | Out-Null		
		#$domainous.nodes.addrange($rootnode)
	}
	$alldomainous=New-Object System.Windows.Forms.TreeNode
	foreach ($n in $domainous.Nodes) {
		My-Verbose "$(Get-Date -Format "HH:mm:ss") : `tDoing sometihing $($n.name)"
		if (!($alldomainous.nodes[$n.name])) { $alldomainous.nodes.add($n.name,$n.name) }
		# if (!($x.Tag -band [Math]::Pow(2,$y))) {$x.Tag+= [Math]::Pow(2,$y)} else {write-host $x.Tag}
		My-Verbose "$($alldomainous.nodes[$n.name].Tag)"
		if (!($alldomainous.nodes[$n.name].Tag -band [Math]::Pow(2,[array]::IndexOf($domains2compare,$domain)))) {$alldomainous.nodes[$n.name].Tag+= [Math]::Pow(2,[array]::IndexOf($domains2compare,$domain))}
		My-Verbose $alldomainous.nodes[$n.name].Tag
	}
	$alldomainous
	
#	$OUtreeView1.Nodes.Clear()
#	$OUtreeView1.BeginUpdate()
#	
#	$comparedous=New-Object System.Windows.Forms.TreeNode
#	$comparedous.Name = "Compared" 
#    $comparedous.Text = "All Domains Compared" 
#    $comparedous.Tag = "Compared" 
#	$comparedous.ImageKey = "Main"
#	$OUtreeView1.nodes.add($comparedous)
#	$OUtreeView1.nodes.addrange($domainous.nodes)
#	$OUtreeView1.EndUpdate()
#
#	$domainous=@()
#	$neutralous=@()
#	# Get all OUS
#	foreach ($domain in $config.Domains.Domain) {
#		$dc=$(Get-ADDomainController -DomainName $domain.Name -Discover).Name+"."+$domain.Name
#		$dom= Get-ADDomain -Server $dc		
#		$adous=Get-ADOrganizationalUnit -Server $dc -Filter * | sort @{Expression={$($_.distinguishedName.split(','))[-1]}},@{Expression={$($_.distinguishedName.split(','))[-2]}},@{Expression={$($_.distinguishedName.split(','))[-3]}},@{Expression={$($_.distinguishedName.split(','))[-4]}},@{Expression={$($_.distinguishedName.split(','))[-5]}},@{Expression={$($_.distinguishedName.split(','))[-6]}}
#		$allous=@()
#		foreach ($ou in $adous) {
#			$allous+=$($ou.distinguishedName -replace "$($dom.distinguishedName)","")
#			$neutralous+=$($ou.distinguishedName -replace "$($dom.distinguishedName)","")
#		}
#		$allous=$allous | sort @{Expression={$($_.split(','))[-2]}},@{Expression={$($_.split(','))[-3]}},@{Expression={$($_.split(','))[-4]}},@{Expression={$($_.split(','))[-5]}},@{Expression={$($_.split(','))[-6]}}
#		$domain=$domain | Add-Member -MemberType NoteProperty -Name OUs -Value $allous -PassThru
#		$domainous+=$domain
#	}
#	$neutralous=$neutralous | sort -Unique | sort @{Expression={$($_.split(','))[-2]}},@{Expression={$($_.split(','))[-3]}},@{Expression={$($_.split(','))[-4]}},@{Expression={$($_.split(','))[-5]}},@{Expression={$($_.split(','))[-6]}}
#	# Compare OUs per domain
#	foreach ($domain in $domainous) {
#		$comparison=Compare-Object $domain.OUs $neutralous -IncludeEqual
#		$comparison=$comparison | sort @{Expression={$($_.InputObject.split(','))[-2]}},@{Expression={$($_.InputObject.split(','))[-3]}},@{Expression={$($_.InputObject.split(','))[-4]}},@{Expression={$($_.InputObject.split(','))[-5]}},@{Expression={$($_.InputObject.split(','))[-6]}}	
#		$domainous[[array]::IndexOf($domainous,$domain)]=$domainous[[array]::IndexOf($domainous,$domain)] | Add-Member -MemberType NoteProperty -Name ComparedOUs -Value $comparison -PassThru
#	}
#	$table = New-Object system.Data.DataTable "Tabel"
#	$col = New-Object system.Data.DataColumn OU,([string])
#	$table.columns.add($col)
#	foreach ($domain in $domainous) {
#		$col = New-Object system.Data.DataColumn $($domain.name),([string])
#		$table.columns.add($col)
#	}
#	$i=0
#	while ($i -lt $domainous[0].ComparedOUs.count) {
#		$row=$table.NewRow()
#		$row.OU=$domainous[0].ComparedOUs[$i].InputObject
#		foreach ($domain in $domainous) {
#			if ($domain.ComparedOUs[$i].SideIndicator -eq "==") {
#				$row.$($domain.name)="OK"			
#				$dom=[adsi] "LDAP://$($domain.name)"
#				$oustring=$domain.ComparedOUs[$i].InputObject+$dom.distinguishedName.ToString()
#				$boundou=New-Object system.DirectoryServices.DirectoryEntry("LDAP://$oustring")
#				#Get-ADOrganizationalUnit -Identity $oustring -Server eettdc20.eett.local (Problem, the dc is not known here anymore)
#				
#				$kids=0
#				if (!($boundou.Children -eq $null)) {
#					foreach ($kid in $boundou.Children) {
#						#if (!(($kid.SchemaClassName -eq "organizationalUnit") -or ($kid.SchemaClassName -eq "container"))) {
#							$kids++
#						#}
#					}
#				}
#				$row.$($domain.name)="$kids"
#			} else {
#				$row.$($domain.name)="MISSING"
#			}
#		}
#		$table.Rows.Add($row)
#		$i++
#	}
#	$testhtml=$table |select *  –ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-HTML -head $cssstyle -Title "EET ADPolice OU Report"
#	$i=0
#	while ($i -lt $testhtml.Count) {
#		if ($testhtml[$i].toupper().contains("MISSING")) {
#			#$testhtml[$i]=$testhtml[$i].replace("<tr>","<tr class=""nok"">")
#			$testhtml[$i]=$testhtml[$i].replace("<td>MISSING</td>","<td class=""nok"">MISSING</td>")
#		}
#		$i++
#	}
#	$testhtml | Out-File "$($config.Domains.Reports.OUDiffPath)\OUDiff_$runtime.htm"
}

Function OutputOUGPOLinkHtmlHeader {
	[cmdletbinding()]
	param($ConfigDomain)
	$htmlheader="<html>"
	$htmlheader+="<head><link rel=""stylesheet"" type=""text/css"" href=""GpoLinkReport.css"" />"
	$htmlheader+="</head><body>"
	$htmlheader+="<h1>GPO Link Report - $($Configdomain.Name)</h1>"
	$htmlheader
}

Function OutputGPOLinkHtmlHeader {
	[cmdletbinding()]
	param()
	$htmlheader="<html>"
	$htmlheader+="<head>"
	$htmlheader+=Get-CSS
	#$htmlheader+="<head><link rel=""stylesheet"" type=""text/css"" href=""ADPolice.css"" />"
	$htmlheader+="</head><body>"
	$htmlheader+="<h1>GPO Link Differences Report</h1>"
	$htmlheader
}

Function OutputGPOLinkHtmlFooter {
	$htmlfooter="</body></html>"
	$htmlfooter
}
function OutputOUGPOLinkTableHeader {
	$tableheader="<table border=""2"" width=""100%"">"
	$tableheader+="<tr><th width=""35%"">OU</th><th width=""5%"">Inheritance Blocked</th><th width=""60%""><table align=""center"" width=""100%"" border=""0""><Caption>Effective GPO's</Caption><tr><th width=""30%"">GPO</th><th width=""60%"">LinkedTo</th><th width=""10%"">Order</th></tr></table></th></tr>"
	$tableheader
}
function OutputGPOLinkTableHeader {
	$tableheader="<table border=""2"" width=""100%"">"
	$tableheader+="<tr><th width=""25%"">OU</th><th width=""25%"">GPO</th><th width=""10%"">Enabled</th><th width=""10%"">Order</th>"
	foreach ($domain in $domains) {
		$tableheader+="<th>$($domain.ShortName)</th>"
	}
	$tableheader+="</tr>"
	$tableheader
}
function OutputGPOLinkTableFooter {
	$tablefooter="</table>"
	$tablefooter
}
function OutputGPOLinkOU ([Microsoft.GroupPolicy.Som]$ouinfo) {
	$tablerow="<tr>"
	$tablerow+="<td>$($ouinfo.Path)</td>"
	$tablerow+="<td>$($ouinfo.GpoInheritanceBlocked)</td>"
	$tablerow+="<td><table class=""internal"" width=""100%"" border=""1"">"
	$ouinfo.InheritedGPoLinks |foreach {
		if ($_.Target -eq $ouinfo.Path) {
			$tablerow+="<tr><td  class=""internal"" width=""30%"">$($_.DisplayName)</td><td class=""internal"" width=""60%"">$($_.Target)</td><td class=""internal"" width=""10%"">$($_.Order)</td></tr>"
		} else {
			$tablerow+="<tr><td  class=""inherited"" width=""30%"">$($_.DisplayName)</td><td class=""inherited"" width=""60%"">$($_.Target)</td><td class=""inherited"" width=""10%"">$($_.Order)</td></tr>"
		}
	}
	$tablerow+="</table></td></tr>"
	$tablerow
}
function OutputGPOLinkOU_withoutinheritance ([Microsoft.GroupPolicy.Som]$ouinfo) {
	$tablerow="<tr>"
	$tablerow+="<td>$($ouinfo.Path)</td>"
	$tablerow+="<td>$($ouinfo.GpoInheritanceBlocked)</td>"
	$tablerow+="<td><table class=""internal"" width=""100%"" border=""1"">"
	$ouinfo.InheritedGPoLinks |foreach {
		if ($_.Target -eq $ouinfo.Path) {
			$tablerow+="<tr><td  class=""internal"" width=""30%"">$($_.DisplayName)</td><td class=""internal"" width=""60%"">$($_.Target)</td><td class=""internal"" width=""10%"">$($_.Order)</td></tr>"
		} else {
			#$tablerow+="<tr><td  class=""inherited"" width=""30%"">$($_.DisplayName)</td><td class=""inherited"" width=""60%"">$($_.Target)</td><td class=""inherited"" width=""10%"">$($_.Order)</td></tr>"
		}
	}
	$tablerow+="</table></td></tr>"
	$tablerow
}
function OutputGPOLinkTable ($gpolinksarray) {
	$tablerow=""
	$gpolinksarray |foreach {
		$tablerow+="<tr><td  class=""internal"" width=""30%"">$($_.Target)</td><td class=""internal"" width=""50%"">$($_.DisplayName)</td><td class=""internal"" width=""10%"">$($_.Enabled)</td><td class=""internal"" width=""10%"">$($_.Order)</td>"
		foreach ($domain in $domains) {
			if ($_.($domain.ShortName)) {
				$tablerow+="<td class=""internal"">$($_.($domain.ShortName))</td>"
			} else {
				$tablerow+="<td class=""nok"">$($_.($domain.ShortName))</td>"
			}
		}
		$tablerow+="</tr>"
	}
	$tablerow
}
function Get-CSS([string]$StyleSheet="..\var\ADPolice.css") {
	<#  
	.SYNOPSIS  
    	Returns css style string loaded from inputfile  
	.DESCRIPTION  
    	This function reads the content of the inputfile (.css)
		and returns it as a string.
	#>
	$style="<style type=""text/css"">"+$(Get-Content $StyleSheet)+"</style>"
	$style
}
function Get-GPOXMLReports($DomainName){
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($DomainName): Getting XML GPOReports"
	[array]$array=Get-GPO -Domain $DomainName -all |Get-GPOReport -Domain $DomainName -ReportType xml | %{
		([xml]$_).gpo | select name,@{n="SOMName";e={$_.LinksTo | % {$_.SOMName}}},@{n="SOMPath";e={$_.LinksTo | %{$_.SOMPath}}},@{n="Computer";e={$_.Computer.ExtensionData}},@{n="User";e={$_.User.ExtensionData}},@{n="ComputerEnabled";e={$_.Computer.Enabled}},@{n="UserEnabled";e={$_.User.Enabled}}		
	}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $DomainName : Retrieved $($array.count) XML GPOReports"
	$array
}
function Get-UnlinkedGPOs{
	[cmdletbinding()]
	param($ConfigDomain)
	foreach ($xmlreport in ($ConfigDomain.GPOXMLReports | where {$_.SomName -eq $null})) {
        Get-GPO -Domain $ConfigDomain.Name -Name $xmlreport.name
    }
}
function Get-GPO-Without-Prefix{
	[cmdletbinding()]
	param($ConfigDomain)
	$strippedgpos=@()
	foreach ($gpo in $(Get-GPO -Domain $ConfigDomain.Name -All)) {
		$strippedgpos+=$($gpo.DisplayName -replace "^$($ConfigDomain.GPOPrefix)","X_")
	}
	$strippedgpos
}
function Get-EmptyGPOs{
	[cmdletbinding()]
	param($ConfigDomain)
	foreach ($xmlreport in ($ConfigDomain.GPOXMLReports | where {(($_.Computer -eq $null) -and ($_.User -eq $null))} )) {
        Get-GPO -Domain $ConfigDomain.Name -Name $xmlreport.name
    }
}
function Get-DisabledGPOs{
	[cmdletbinding()]
	param($ConfigDomain)
	foreach ($xmlreport in ($ConfigDomain.GPOXMLReports | where {
        (($_.Computer -eq $null) -and ($_.UserEnabled -eq "false") -and ($_.ComputerEnabled -eq "true")) `
        -or (($_.User -eq $null) -and ($_.ComputerEnabled -eq "false") -and ($_.UserEnabled -eq "true")) `
        -or (($_.ComputerEnabled -eq "false") -and ($_.UserEnabled -eq "false"))
    })) {
        Get-GPO -Domain $ConfigDomain.Name -Name $xmlreport.name
    }
}
function BogusGPOstotable{
	param ([array]$unlinked,[array]$disabled,[array]$empty)
	$allbogusGPOs=@{}
	foreach ($gpo in $unlinked) {
		if ($gpo -ne $null) {
			$allbogusGPOs+=@{$gpo.Id=@{Domain=$gpo.DomainName;GPO=$gpo.DisplayName;Unlinked="&#10004;";Disabled="";Empty=""}}
		}
	}
	foreach ($gpo in $disabled) {
		if ($gpo -ne $null) {
			if ($allbogusGPOs.Keys -contains $gpo.Id) {
				$allbogusGPOs.Item($gpo.Id).Disabled="&#10004;"
			} else {
				$allbogusGPOs+=@{$gpo.Id=@{Domain=$gpo.DomainName;GPO=$gpo.DisplayName;Unlinked="";Disabled="&#10004;";Empty=""}}
			}
		}
	}
	foreach ($gpo in $empty) {
		if ($gpo -ne $null) {
			if ($allbogusGPOs.Keys -contains $gpo.Id) {
				$allbogusGPOs.Item($gpo.Id).Empty="&#10004;"
			} else {
				$allbogusGPOs+=@{$gpo.Id=@{Domain=$gpo.DomainName;GPO=$gpo.DisplayName;Unlinked="";Disabled="";Empty="&#10004;"}}
			}
		}
	}
	$HTMLtable="<table align=""center"" width=""100%"" border=""0""><Caption>Not Applicable GPO's</Caption><tr><th width=""20%"">Domain</th><th width=""50%"">GPO</th><th width=""10%"">Unlinked</th><th width=""10%"">Non-empty Settings<br>Disabled</th><th width=""10%"">Enabled Settings<br>Empty</th></tr>"
	#foreach ($row in $allbogusGPOs.Keys) {
	foreach($row in $allbogusGPOs.GetEnumerator()| sort {$_.value.gpo.substring(2)}) {
		#$HTMLtable+="<tr><td>$($allbogusGPOs.Item($row).Domain)</td><td>$($allbogusGPOs.Item($row).GPO)</td><td>$($allbogusGPOs.Item($row).Unlinked)</td><td>$($allbogusGPOs.Item($row).Disabled)</td><td>$($allbogusGPOs.Item($row).Empty)</td></tr>`n"
		$HTMLtable+="<tr><td>$($row.Value.Domain)</td><td>$($row.Value.GPO)</td><td>$($row.Value.Unlinked)</td><td>$($row.Value.Disabled)</td><td>$($row.Value.Empty)</td></tr>`n"
	}
	$HTMLtable+="</table>"
	$HTMLtable
}
function Get-GPOBackup {
	[cmdletbinding()]
	Param(
		[Parameter(Position=0)]$ConfigDomain=$(throw "ConfigDomain required."), 
		#[Parameter(Position=0,Mandatory=$true,HelpMessage="What is the path to the GPO backup folder?")]
		#[ValidateNotNullOrEmpty()]
		#[string]$Path=$global:GPBackupPath,
		[Parameter(Position=1)]
		[string]$Name,
		[switch]$Latest
	)
	#validate $Path
	if (-Not $ConfigDomain.GPOBackupPath) {
	  throw "GPOBackup path not defined"
	 }
	Try 
	{
	    My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Validating $($ConfigDomain.GPOBackupPath)"
	    if (-Not (Test-Path $ConfigDomain.GPOBackupPath)) { Throw }
	}
	Catch 
	{
	    My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Failed to find GPOBackupPath $($ConfigDomain.GPOBackupPath)"
	    Break
	}

	#get each folder that looks like a GUID
	[regex]$regex="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"

	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Enumerating folders under $($ConfigDomain.GPOBackupPath)"

	#define an array to hold each backup object
	$Backups=@()

	#find all folders named with a GUID
	Get-ChildItem -Path $ConfigDomain.GPOBackupPath | Where {$_.name -Match $regex -AND $_.PSIsContainer} |
	foreach {

	  #import the Bkupinfo.xml file
	  $file=Join-Path $_.FullName -ChildPath "bkUpinfo.xml"
	  My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Importing $file"
	  [xml]$data=Get-Content -Path $file 
	  
	  #parse the xml file for data
	  $GPO=$data.BackupInst.GPODisplayName."#cdata-section"
	  $GPOGuid=$data.BackupInst.GPOGuid."#cdata-section"
	  $ID=$data.BackupInst.ID."#cdata-section"
	  $Comment=$data.BackupInst.comment."#cdata-section"
	  #convert backup time to a date time object
	  [datetime]$Backup=$data.BackupInst.BackupTime."#cdata-section"
	  $GPODomain=$data.BackupInst.GPODomain."#cdata-section"

	  #write a custom object to the pipeline
	  $Backups+=New-Object -TypeName PSObject -Property @{
	    Name=$GPO
	    Comment=$Comment
	    #strip off the {} from the Backup ID GUID
	    BackupID=$ID.Replace("{","").Replace("}","")
	    #strip off the {} from the GPO GUID
	    Guid=$GPOGuid.Replace("{","").Replace("}","")
	    Backup=$Backup
	    Domain=$GPODomain
	    Path=$Path
	 }
	 } #foreach	 
	 #if searching by GPO name, then filter and get just those GPOs
	 if ($Name)
	 {
	    My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Filtering for GPO: $Name"
	    $Backups=$Backups | where {$_.Name -like $Name}	 
	 }	 
	 My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Found $($Backups.Count) GPO Backups"	 
	 #if -Latest then only write out the most current version of each GPO
	 if ($Latest) 
	 {
	    My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting Latest Backups"
	    $grouped=$Backups | Sort-Object -Property GUID | Group-Object -Property GUID
	    $grouped | Foreach {
	        $_.Group | Sort-Object -Property Backup | Select-Object -Last 1
	    }
	 }
	 else
	 {
	    $Backups
	 }	 
	 My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Ending function Get-GPOBackup"
}
Function Backup_GPOs {
[cmdletbinding()]
param(
	[string]$ReportRecipient,
	[string]$smtpserver="mail",
	[Parameter(Position=0)]$ConfigDomain=$(throw "ConfigDomain required."),
	$Config
)
	$changes=$false
	if ($config.GPOBackup.SendResult -eq "true") {
		if ($config.Mail.SmtpHost -eq $null) {$smtpserver = "mail"} else {$smtpserver = $config.Mail.SmtpHost}
    	$mailer = new-object Net.Mail.SMTPclient($smtpserver)
    	$msg = new-object Net.Mail.MailMessage($config.GPOBackup.Sender,$config.GPOBackup.Recipient)
	}
	[String]$result="GPO Change Reports`r`n==============`r`n"
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Creating Report Path $($ConfigDomain.GPOReportPath) if not exist"	
	if (!(Test-Path -path $ConfigDomain.GPOReportPath)) {New-Item $ConfigDomain.GPOReportPath -type directory}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Creating Backup Path $($ConfigDomain.GPOBackupPath) if not exist"
	if (!(Test-Path -path $ConfigDomain.GPOBackupPath)) {New-Item $ConfigDomain.GPOBackupPath -type directory}
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting all GPOs"
	$AllGPOs=Get-GPO -All -Domain $ConfigDomain.Name
	My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Getting old Backups from $($ConfigDomain.GPOBackupPath)"
	$GPMBackupDir=$GPM.GetBackupDir($Configdomain.GPOBackupPath)
	$GPMSearchCriteria = $GPM.CreateSearchCriteria()
	$GPMSearchCriteria.Add($Constants.SearchPropertyBackupMostRecent, $Constants.SearchOpEquals, $true)
	$Backups=$GPMBackupDir.SearchBackups($GPMSearchCriteria)
	Foreach ($GPO in $AllGPOs) {
		$needbackup=$false
		My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): Checking $($GPO.DisplayName)"
		$LastBackup=$Backups | where {$_.GPOID -eq "{$($GPO.Id)}" }
		if ($LastBackup) {
			My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tLast Backup found from $($LastBackup.Timestamp)"
			if ($GPO.ModificationTime -gt $LastBackup.Timestamp) {
				$result+="$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): $($GPO.DisplayName), Last Backup found from $($LastBackup.Timestamp), Last modified $($GPO.ModificationTime), so created backup`r`n"
				$result+=get-gpochange-auditevent -gpo $gpo -StartTime $LastBackup.Timestamp
				My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tLast modified $($GPO.ModificationTime), so need backup"
				$needbackup=$true
			} else {
				My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tLast modified $($GPO.ModificationTime), so already in backup"
			}
		} else {
			$result+="$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): $($GPO.DisplayName), No previous backup found, so created backup`r`n"
			#$result+=get-gpochange-auditevent -gpo $gpo
			My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tNo previous backup found, so need backup"
			$needbackup=$true
		}
		if ($needbackup) {
			$changes=$true
			$Description="Automated backup"
    		$GPOBackup = Backup-GPO -Guid $GPO.Id -Path $ConfigDomain.GPOBackupPath -Comment $Description -Domain $ConfigDomain.Name
			My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tBacked up"	    		
			$ReportPath = $ConfigDomain.GPOReportPath + "\"+$GPO.Displayname + "_" + $GPO.ModificationTime.Month + "-"+ $GPO.ModificationTime.Day + "-" + $GPO.ModificationTime.Year + "_" + $GPOBackup.Id + ".html" 
			Get-GPOReport -Name $GPO.DisplayName -path $ReportPath -ReportType HTML -Domain $ConfigDomain.Name
			My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tSaved Report"
			if ($config.GPOBackup.SaveDiffReports -eq "true") {
				$DiffReportPath = $ConfigDomain.GPOReportPath + "\Diff_"+$GPO.Displayname + "_" + $GPO.ModificationTime.Month + "-"+ $GPO.ModificationTime.Day + "-" + $GPO.ModificationTime.Year + "_" + $GPOBackup.Id + ".html" 
				$previousreport=Get-PreviousReport -ConfigDomain $ConfigDomain -gponame $GPO.Displayname -currentreport $ReportPath
				Create-GPOReportDiff -gpo1 $(get-content $previousreport) -gpo2 $(get-content $ReportPath) | Out-File $DiffReportPath
				My-Verbose "$(Get-Date -Format "HH:mm:ss") $($ConfigDomain.Name): `tSaved DiffReport"
			}
			if ($config.GPOBackup.SendResult -eq "true") {
				if ($config.GPOBackup.AttachDiffReports -eq "true") {
					$attachment = new-object Net.Mail.Attachment($DiffReportPath)
        			$msg.attachments.add($attachment)
				}
			}						
		}    	
	}
	if (($config.GPOBackup.SendResult -eq "true")-and $changes) {
		$msg.body=$result.tostring()
		$msg.Subject="$($configdomain.Name) - $($config.GPOBackup.Subject)"
    	$mailer.send($msg)
	}
}
function SEND-ZIP ($zipfilename, $filename) { 
	# The $zipHeader variable contains all the data that needs to sit on the top of the  
	# Binary file for a standard .ZIP file 
	$zipHeader=[char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 
	# Check to see if the Zip file exists, if not create a blank one 
	If ( (TEST-PATH $zipfilename) -eq $FALSE ) { Add-Content $zipfilename -value $zipHeader } 
	$ExplorerShell=NEW-OBJECT -comobject 'Shell.Application' 
	$SendToZip=$ExplorerShell.Namespace($zipfilename.tostring()).MoveHere($filename.ToString()) 
} 
Function Create-GPOReportDiff($gpo1,$gpo2) {
$oldhtml=$gpo1
$newhtml=$gpo2
[Helpers.HtmlDiff] $diff=new-object helpers.HtmlDiff($oldhtml, $newhtml)
$html=$diff.Build()
$newstyle="<style type=""text/css""> ins { color: #FFFF00; background-color: #0000FF; text-decoration: none; } del { color: #0000FF; background-color:#FFFF00; text-decoration: none; } "
$html=$html.Replace("<style type=""text/css"">",$newstyle)
return $html
}
Function Get-GPOChange-AuditEvent {
	[cmdletbinding()]
	param(
		[Parameter(Position=0)]$GPO=$(throw "GPO required."),
		[DateTime]$StartTime=$GPO.ModificationTime - (New-TimeSpan -Minutes 5),
		[DateTime]$EndTime=$GPO.ModificationTime + (New-TimeSpan -Minutes 5)
	)
	$ADInfo = Get-ADDomain $gpo.DomainName
	$ADDomainReadOnlyReplicaDirectoryServers = $ADInfo.ReadOnlyReplicaDirectoryServers
	$ADDomainReplicaDirectoryServers = $ADInfo.ReplicaDirectoryServers
	$DomainControllers = $ADDomainReadOnlyReplicaDirectoryServers + $ADDomainReplicaDirectoryServers
	$events=@()
	foreach ($dc in $DomainControllers) {
		$events+=Get-WinEvent -ComputerName $dc -FilterHashtable @{ProviderName="Microsoft-Windows-Security-Auditing";ID=@(5136..5137);StartTime=$StartTime;EndTime=$EndTime} -ErrorAction SilentlyContinue
    }
    $events=$events | sort -Property TimeCreated
	$events|?{($_.properties[8].value -match "CN={$($gpo.Id)},") -and ($_.properties[14].value -eq "%%14674")}|%{
		"$($_.TimeCreated) : $($_.properties[3].value) saved version $($_.properties[13].value) of $($GPO.DisplayName)`r`n"
	}
	#$events | ft TimeCreated,Id,@{expression={$_.properties[3].value}},@{expression={$_.properties[8].value}},@{expression={$_.properties[14].value}},@{expression={$_.properties[10].value}},@{expression={$_.properties[11].value}},@{expression={$_.properties[13].value}} -AutoSize
	# id:
	# 5137 create ad object
	# 5136 modify ad object
	# properties:
	# 3 = username
	# 4 = userdomain
	# 8 = distinguishedname of object
	# 10 = type of object (groupPolicyContainer,organizationalunit, etc)
	# 11 = valuename
	# 13 = value
	# 14 = %%14674 (value added), %%14675 (value deleted)
}
Function Domain-Neutral([string]$domainspecificstring,$configdomains) {
	[string]$domainneutralstring=$domainspecificstring
	foreach ($dom in $configdomains) {
		$domainneutralstring=$domainneutralstring -replace "^$($dom.GPOPrefix)","X_"
		$domainneutralstring=$domainneutralstring -replace "$($dom.Name)","domainXroot"
		$domainneutralstring=$domainneutralstring -replace "$($dom.ShortName)","domainX"
	}
	$domainneutralstring
}
Function Domain-Specific([string]$domainneutralstring,$configdomain) {
	[string]$domainspecificstring=$domainneutralstring
		$domainspecificstring=$domainspecificstring -replace "^X_","$($configdomain.GPOPrefix)"
		$domainspecificstring=$domainspecificstring -replace "domainXroot","$($configdomain.Name)"
		$domainspecificstring=$domainspecificstring -replace "domainX","$($configdomain.ShortName)"
	$domainspecificstring
}
#endregion

#region Global variables
My-Verbose "Creating Global Variables" -1
$GPM = New-Object -comobject GPMgmt.GPM
My-Verbose "`$GPM as the Group Policy COM-object"
$constants=$gpm.GetConstants()
My-Verbose "`$constants as the Group Policy COM-object constants"
#endregion Global variables