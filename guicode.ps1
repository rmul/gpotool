#region Custom code
$CompareButton.enabled=$false
$ConfigDataset=New-Object System.Data.DataSet
$wb=New-Object system.Windows.Forms.WebBrowser
$ReportTabPage.Controls.Add($wb)
$wb.Dock=5
$OUtreeView1.ImageList=new-object system.Windows.Forms.ImageList
$System_Drawing_Size = New-Object System.Drawing.Size 
$System_Drawing_Size.Width = 48 
$System_Drawing_Size.Height = 16 
$OUtreeView1.ImageList.ImageSize = $System_Drawing_Size
$OUtreeView1.Imagelist.Images.Add("Main",[System.Drawing.Image]::FromFile(".\Images\eye.gif"))
$OUtreeView1.Imagelist.Images.Add("Blank",[System.Drawing.Image]::FromFile(".\Images\blank.gif"))
$OUtreeView1.Imagelist.Images.Add("TAP",[System.Drawing.Image]::FromFile(".\Images\TAP.gif"))
$OUtreeView1.Imagelist.Images.Add("xAP",[System.Drawing.Image]::FromFile(".\Images\xAP.gif"))
$OUtreeView1.Imagelist.Images.Add("TxP",[System.Drawing.Image]::FromFile(".\Images\TxP.gif"))
$OUtreeView1.Imagelist.Images.Add("TAx",[System.Drawing.Image]::FromFile(".\Images\TAx.gif"))
$OUtreeView1.Imagelist.Images.Add("xxP",[System.Drawing.Image]::FromFile(".\Images\xxP.gif"))
$OUtreeView1.Imagelist.Images.Add("xAx",[System.Drawing.Image]::FromFile(".\Images\xAx.gif"))
$OUtreeView1.Imagelist.Images.Add("Txx",[System.Drawing.Image]::FromFile(".\Images\Txx.gif"))


#Extra Functions
function PopulateConfigGridView {
	[cmdletbinding()]
	param ($filename)
	$ConfigDataset.Clear()
	$ConfigDataset.ReadXml($filename)
	$ConfigurationdataGridView.Datasource=$ConfigDataset.Tables["Domain"].Defaultview
	$ConfigurationdataGridView.AutoResizeColumns()
	$ConfigurationdataGridView1.Datasource=$ConfigDataset.Tables["Reports"].Defaultview
	$ConfigurationdataGridView1.AutoResizeColumns()
	$ConfigurationdataGridView2.Datasource=$ConfigDataset.Tables["Mail"].Defaultview
	$ConfigurationdataGridView2.AutoResizeColumns()
	$ConfigDataset.WriteXml("c:\SBP\Scripts\outxml.xml")
}
function Add-Node { 
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
function ou-treenodes {
	[cmdletbinding()]
	param ($node,$dn,$svr)	
	Get-ADOrganizationalUnit -Server $dc -Filter * -SearchScope 1 -SearchBase $dn | %{
		$_.distinguishedname
		$newnode=add-node $node $_.Name "OU"
		ou-treenodes $newnode $_.distinguishedname $svr
	}
}


#Custom Handlers

$handler_OtherGPOlistBox_SelectedIndexChanged= 
{
	if ($GPOlistBox.selectedindex -ge 0) {
		$CompareGPOButton.enabled=$true
	}
}

$handler_DomainlistBox_SelectedIndexChanged= 
{
	$OtherGPOlistBox.items.clear()
	$domainindex=$DomainListBox.selectedindex
	$selecteddomain=$configuration.Domains.Domain[$domainindex]
	$selectedneutralgpos=Get-GPO-Without-Prefix $selecteddomain
	$selectedneutralgpos=$selectedneutralgpos | sort
	$selectedneutralgpos=$selectedneutralgpos | sort -Unique
	$OtherGPOlistBox.Items.AddRange($selectedneutralgpos)
}

$handler_CompareGPObutton_Click= 
{
	$domainindex=$DomainListBox.selectedindex
	$a_policy=$($GPOlistBox.SelectedItem -replace "^X_","$($configuration.Domains.Domain[$domainindex].GPOPrefix)")
	$b_policy=$($OtherGPOlistBox.SelectedItem -replace "^X_","$($configuration.Domains.Domain[$domainindex].GPOPrefix)")
	$a_report=Get-GPOReport -Name $a_policy -ReportType HTML -Domain $configuration.Domains.Domain[$domainindex].Name
	$b_report=Get-GPOReport -Name $b_policy -ReportType HTML -Domain $configuration.Domains.Domain[$domainindex].Name
	$wb.DocumentText=Create-GPOReportDiff -gpo1 $a_report -gpo2 $b_report
	$Maintabcontrol.SelectedTab=$reportTabPage

}

$handler_MaintabControl_SelectedIndexChanged= 
{
#TODO: Place custom script here
	if ($Maintabcontrol.SelectedTab -eq $ReportTabPage) {
		$SaveButton.Visible=$true
	} else {
		$SaveButton.Visible=$false
	}
	if ($Maintabcontrol.SelectedTab -eq $OUtabPage) {
		. $handler_OUbutton_Click
	}

}

$SaveButton_OnClick= 
{
#TODO: Place custom script here
	$savefiledialog1.CreatePrompt=$false
	$savefiledialog1.OverwritePrompt=$true
	$savefiledialog1.ShowDialog()

}

$handler_saveFileDialog1_FileOk= 
{
#TODO: Place custom script here
	$wb.DocumentText | Out-File $savefiledialog1.FileName

}

$handler_ConfigLoadbutton_Click= 
{
#TODO: Place custom script here
	$openFileDialog1.ShowDialog()
}

$handler_CompareButton_Click= 
{
	$source=$SourceListBox.selectedindex
	$dest=$DestListBox.selectedindex
	$a_policy=$($GPOlistBox.SelectedItem -replace "^X_","$($configuration.Domains.Domain[$source].GPOPrefix)")
	$b_policy=$($GPOlistBox.SelectedItem -replace "^X_","$($configuration.Domains.Domain[$dest].GPOPrefix)")
	$a_report=Get-GPOReport -Name $a_policy -ReportType HTML -Domain $configuration.Domains.Domain[$source].Name
	$b_report=Get-GPOReport -Name $b_policy -ReportType HTML -Domain $configuration.Domains.Domain[$dest].Name
	$wb.DocumentText=Create-GPOReportDiff -gpo1 $a_report -gpo2 $b_report
	$Maintabcontrol.SelectedTab=$reportTabPage
}

$handler_ConfigFile_OK= 
{
#TODO: Place custom script here
	$configfile=$openFileDialog1.FileName
	$global:configuration=LoadConfig $configfile
	#$ConfigLoadbutton.Text=$configfile
	$neutralgpos=@()
	$GPOlistBox.Items.Clear()
	$SourceListBox.Items.clear()
	$DestListBox.Items.clear()
	$DomainListBox.Items.clear()
	foreach ($domain in $configuration.Domains.Domain) {
		$SourceListBox.Items.Add($domain.Name)
		$DestListBox.Items.Add($domain.Name)
		$DomainListBox.Items.Add($domain.Name)
		$neutralgpos+=Get-GPO-Without-Prefix $domain
	}
	$SourceListBox.setSelected(0,$true)
	$DestListBox.setSelected(1,$true)
	$neutralGPOs=$neutralGPOs | sort
	$neutralGPOs=$neutralGPOs | sort -Unique
	$GPOlistBox.Items.AddRange($neutralgpos)
	$GPOlistBox.Visible=$true
	$tabControl1.Visible=$true
	PopulateConfigGridView $configfile
	
	
}

$handler_ExitButton_Click= 
{
#TODO: Place custom script here
	$form1.close()
}

$handler_GPOlistBox_SelectedIndexChanged=
{
	$CompareButton.enabled=$true
	$GPOHistoryButton.enabled=$true
		if ($OtherGPOlistBox.selectedindex -ge 0) {
		$CompareGPOButton.enabled=$true
	}
}
$handler_GPOHistoryButton_Click= 
{
	$GPOHistoryTextBox.Text=$null
	$GPOHistoryDataGridView.Rows.Clear()
	$GPOHistoryTextBox.Text=$GPOlistBox.SelectedItem
	foreach ($Configdomain in $configuration.Domains.Domain) {
		$gponame=$GPOlistBox.SelectedItem -replace "^X_",$($Configdomain.GPOPrefix)
		$GPOHistoryTextBox.Text+="==================`r`n"
		$GPOHistoryTextBox.Text+="$($Configdomain.Name) => $gponame`r`n"
		$GPOHistoryTextBox.Text+="------------------`r`n"
		$GPMBackupDir=$GPM.GetBackupDir($Configdomain.GPOBackupPath)
		$GPMSearchCriteria = $GPM.CreateSearchCriteria()
		$GPMSearchCriteria.Add($Constants.SearchPropertyGPODisplayName, $Constants.SearchOpEquals, $gponame)
		$Backups=$GPMBackupDir.SearchBackups($GPMSearchCriteria)
		
		foreach ($backup in $Backups) {
			[xml]$xmlreport=$backup.GenerateReport($constants.ReportXML).result
			$GPOHistoryTextBox.Text+="Modified: `t$(get-date $xmlreport.GPO.ModifiedTime)`r`n"
			$GPOHistoryTextBox.Text+="Backed up:`t$($backup.Timestamp)`r`n"
			$GPOHistoryTextBox.Text+="`r`n"
			$row=@($ConfigDomain.Name,$(get-date $xmlreport.GPO.ModifiedTime),$backup.Timestamp,$backup.ID)
			$GPOHistoryDataGridView.Rows.Add($row)
		}
	}
	$GPOHistoryDataGridView.Sort($GPOHistoryDataGridView.Columns[2],'Descending')
}

$handler_GPOHistoryReportButton_Click= 
{
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.index)`r`n"
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.cells[3].value)`r`n"
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.cells[0].value)`r`n"
	$GPMBackupDir=$gpm.GetBackupDir(($configuration.Domains.Domain | ?{$_.Name -eq $GPOHistoryDataGridView.CurrentRow.cells[0].value}).GPOBackupPath)
	$GPMBackup=$GPMBackupDir.GetBackup($GPOHistoryDataGridView.CurrentRow.cells[3].value)
	$wb.DocumentText=$GPMBackup.GenerateReport($constants.ReportHTML).result
	$Maintabcontrol.SelectedTab=$reportTabPage
}

$handler_GPOBackupDiffbutton_Click= 
{
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.index)`r`n"
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.cells[3].value)`r`n"
	$GPOHistoryTextBox.Text+="$($GPOHistoryDataGridView.CurrentRow.cells[0].value)`r`n"
	$GPMBackupDir=$gpm.GetBackupDir(($configuration.Domains.Domain | ?{$_.Name -eq $GPOHistoryDataGridView.CurrentRow.cells[0].value}).GPOBackupPath)
	$GPMBackup=$GPMBackupDir.GetBackup($GPOHistoryDataGridView.CurrentRow.cells[3].value)

	$GPMSearchCriteria = $GPM.CreateSearchCriteria()
	$GPMSearchCriteria.Add($Constants.SearchPropertyGPODisplayName, $Constants.SearchOpEquals, $GPMBackup.GPODisplayName)
	$Backups=$GPMBackupDir.SearchBackups($GPMSearchCriteria) | sort Timestamp
	$index=0
	while ($index -lt $Backups.Count) {
		if ($Backups[$index].ID -eq $GPMbackup.ID) { $index; break }
		$index++
	}
	if ($index -ne 0) {
		$a_report=$GPMBackup.GenerateReport($constants.ReportHTML).result
		$b_report=$Backups[$index-1].GenerateReport($constants.ReportHTML).result
		$wb.DocumentText=Create-GPOReportDiff -gpo1 $a_report -gpo2 $b_report
		$Maintabcontrol.SelectedTab=$reportTabPage
	}
}

$handler_GPOHistoryDataGridView_SelectionChanged={
	$GPOHistoryReportButton.enabled=$true
	$GPMBackupDir=$gpm.GetBackupDir(($configuration.Domains.Domain | ?{$_.Name -eq $GPOHistoryDataGridView.CurrentRow.cells[0].value}).GPOBackupPath)
	$GPMBackup=$GPMBackupDir.GetBackup($GPOHistoryDataGridView.CurrentRow.cells[3].value)

	$GPMSearchCriteria = $GPM.CreateSearchCriteria()
	$GPMSearchCriteria.Add($Constants.SearchPropertyGPODisplayName, $Constants.SearchOpEquals, $GPMBackup.GPODisplayName)
	$Backups=$GPMBackupDir.SearchBackups($GPMSearchCriteria) | sort Timestamp
	$index=0
	while ($index -lt $Backups.Count) {
		if ($Backups[$index].ID -eq $GPMbackup.ID) { $index; break }
		$index++
	}
	if ($index -ne 0) {
		$GPOBackupDiffbutton.enabled=$true
	} else {
		$GPOBackupDiffbutton.enabled=$false
	}
}
$handler_OUbutton_Click= 
{
	$tn=Compare-OU-Trees $configuration.Domains.Domain
	$domainous=New-Object System.Windows.Forms.TreeNode
	$domainous.Name = "Compared" 
    $domainous.Text = "All Domains" 
    $domainous.Tag = "Compared" 
	$domainous.ImageKey = "Blank"
	$OUtreeView1.Nodes.Clear()
	$OUtreeView1.BeginUpdate()
	foreach ($domain in $configuration.Domains.Domain) {
		$ous=New-Object System.Windows.Forms.TreeNode
		$dc=$(Get-ADDomainController -DomainName $domain.Name -Discover).Name+"."+$domain.Name
		$dom= Get-ADDomain -Server $dc		
		$rootnode=add-node $ous $domain.Name "root"
		$rootnode.imagekey="Main"
		ou-treenodes -node $rootnode -dn $dom.DistinguishedName -svr $dc
		$domainous.nodes.addrange($rootnode)
	}
	$comparedous=New-Object System.Windows.Forms.TreeNode
	$comparedous.Name = "Compared" 
    $comparedous.Text = "All Domains Compared" 
    $comparedous.Tag = "Compared" 
	$comparedous.ImageKey = "Main"
	$OUtreeView1.nodes.add($comparedous)
	$OUtreeView1.nodes.addrange($domainous.nodes)
	$OUtreeView1.EndUpdate()
}



$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
	if (Test-Path $openFileDialog1.FileName) {
		. $handler_ConfigFile_OK
	}
}

#endregion Custom code