cls
$rc=[Reflection.Assembly]::LoadFile(“C:\SBP\Scripts\htmldiff.dll”)
Import-Module activedirectory
Import-Module grouppolicy

$reportfolder="E:\GPOLinkInfo"
$smtpserver="mail"
$notify="rmul@schubergphilis.com"
$notifier=$env:COMPUTERNAME+"@"+[System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name

$domain=Get-ADDomain
$now=get-date
$runtime=$(Get-Date $now -Format "yyyyMMddhhmmss")

function Send-SMTPmail($to, $from, $subject, $body, $attachment, $cc, $bcc, $port, $timeout, $smtpserver, [switch] $html, [switch] $alert) {
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
function OutputHtmlHeader {
	$htmlheader="<html>"
	$htmlheader+="<head><link rel=""stylesheet"" type=""text/css"" href=""GpoLinkReport.css"" />"
	$htmlheader+="</head><body>"
	$htmlheader+="<h1>GPO Link Report - $($domain.NetBIOSName)</h1>"
	$htmlheader
}
function OutputHtmlFooter {
	$htmlfooter="</body></html>"
	$htmlfooter
}
function OutputTableHeader {
	$tableheader="<table border=""2"" width=""100%"">"
	$tableheader+="<tr><th width=""35%"">OU</th><th width=""5%"">Inheritance Blocked</th><th width=""60%""><table align=""center"" width=""100%"" border=""0""><Caption>Effective GPO's</Caption><tr><th width=""30%"">GPO</th><th width=""60%"">LinkedTo</th><th width=""10%"">Order</th></tr></table></th></tr>"
	$tableheader
}
function OutputTableFooter {
	$tablefooter="</table>"
	$tablefooter
}
function OutputOU ([Microsoft.GroupPolicy.Som]$ouinfo) {
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

$adous=Get-ADOrganizationalUnit -Filter * | sort @{Expression={$($_.distinguishedName.split(','))[-1]}},@{Expression={$($_.distinguishedName.split(','))[-2]}},@{Expression={$($_.distinguishedName.split(','))[-3]}},@{Expression={$($_.distinguishedName.split(','))[-4]}},@{Expression={$($_.distinguishedName.split(','))[-5]}},@{Expression={$($_.distinguishedName.split(','))[-6]}}

[Array]$gpinheritance=get-gpinheritance -target $domain.DistinguishedName
$gpinheritance+=$adous|%{ (Get-GPInheritance -Target $_)}

[array]$oldreports=get-childitem "$reportfolder\*" -Include "GpoLinkReport_*.html" | sort CreationTime
$oldreport=Get-Content $oldreports[-1]

$report=OutputHtmlHeader
$report+=OutputTableHeader
$report+=$gpinheritance | foreach {OutputOU $_}
$report+=OutputTableFooter
$report+=OutputHtmlFooter
$compare=Compare-Object $oldreport $report
if ($compare) {
	$report | Out-File "$reportfolder\GpoLinkReport_$runtime.html"
	[Helpers.HtmlDiff] $diff=new-object helpers.HtmlDiff($oldreport, $report)
	$html=$diff.Build()
	$style="<style type=""text/css"">"+$(Get-Content "$reportfolder\GpoLinkReport.css")+"</style>"
	$html=$html.Replace("<link rel=""stylesheet"" type=""text/css"" href=""GpoLinkReport.css"" />",$style)
	$html=$html.Replace("</h1>","</h1><h2>Generated on $now</h2>")
	$html | Out-File "$reportfolder\GpoLinkReportDiff_$runtime.html"
	send-SMTPmail -to $notify -from $notifier -attachment "$reportfolder\GpoLinkReportDiff_$runtime.html" -subject "$($domain.NetBIOSName) GPOLink Change Report - $now" -smtpserver $smtpserver -html -body $html
} else {
	"Nothing changed"
}



