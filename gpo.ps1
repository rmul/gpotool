cls
[xml]$xml=Get-Content ".\Default Domain Policy.xml"
$extension1=$xml.GPO.Computer.ExtensionData[1].Extension
$childnodes=$extension1.ChildNodes
$childnodes | ft -Property Name

