Add-Type -AssemblyName System.Windows.Forms

$DefaultFont = 'Times New Roman,12'

# Define the initial form
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.ClientSize = '300,150'
$MainForm.Text = 'GPO Management'
$MainForm.BackColor = 'white'
$MainForm.Font = $DefaultFont

# Define buttons for the main form
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = 'Backup'
$btnBackup.AutoSize = $true
$btnBackup.Location = New-Object System.Drawing.Point(10,10)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = 'Import'
$btnImport.AutoSize = $true
$btnImport.Location = New-Object System.Drawing.Point(100,10)

$btnLink = New-Object System.Windows.Forms.Button
$btnLink.Text = 'Link'
$btnLink.AutoSize = $true
$btnLink.Location = New-Object System.Drawing.Point(190,10)

# Add click event handlers for each button
$btnBackup.Add_Click({ ShowBackupForm })
$btnImport.Add_Click({ ShowImportForm })
$btnLink.Add_Click({ ShowLinkForm })

$MainForm.Controls.AddRange(@($btnBackup, $btnImport, $btnLink))

# Function definitions for each process
function ShowBackupForm {
    # Paste the content of backupGPO.ps1 here
    Add-Type -AssemblyName System.Windows.Forms

    $FormObject=[System.Windows.Forms.Form]
    $LabelObject=[System.Windows.Forms.Label]
    $CheckedListBoxObject=[System.Windows.Forms.CheckedListBox]
    $ButtonObject=[System.Windows.Forms.Button]

    $DefaultFont='Times New Roman,12'

    # Set up base form
    $GPOForm=New-Object $FormObject
    $GPOForm.ClientSize='700,200'
    $GPOForm.Text='Export GPOs'
    $GPOForm.BackColor='white'
    $GPOForm.Font=$DefaultFont
    $GPOForm.WindowState = 'Maximized'  # Set the window state to Maximized

    # Building the form
    $lblGPO=New-Object $LabelObject
    $lblGPO.Text='GPOs:'
    $lblGPO.Autosize=$true
    $lblGPO.Location=New-Object System.Drawing.Point(20,20)

    $ddlGPO=New-Object $CheckedListBoxObject
    $ddlGPO.Size = New-Object System.Drawing.Size(950,600)
    $ddlGPO.Location=New-Object System.Drawing.Point(70,20)

    $btnGPO=New-Object $ButtonObject
    $btnGPO.Text='Export'
    $btnGPO.AutoSize=$true
    $btnGPO.Location=New-Object System.Drawing.Point(400,650)

    $btnSelectAll=New-Object $ButtonObject
    $btnSelectAll.Text='Select All'
    $btnSelectAll.AutoSize=$true
    $btnSelectAll.Location=New-Object System.Drawing.Point(600,650)

    Get-GPO -all | ForEach-Object {$ddlGPO.Items.Add($_.DisplayName)}

    $btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $ddlGPO.Items.Count; $i++) {
            $ddlGPO.SetItemChecked($i, $true)
        }
    })

    $btnGPO.Add_Click({ 
        $selectedItems = $ddlGPO.CheckedItems | ForEach-Object {$_.ToString()}  
        Write-Host "Selected GPOs: $($selectedItems -join ',')"

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog 
        $dialog.Description = "Select a folder to save the backup files"

        if ($dialog.ShowDialog() -eq 'OK') {
            $backupPath = $dialog.SelectedPath
        
            foreach ($gpoName in $selectedItems) {
                Backup-GPO -Name $gpoName -Path $backupPath
                Write-Host "Backup of GPO '$gpoName' created at '$backupPath'"
            }

            $numItems = (Get-ChildItem -Path $backupPath | Measure-Object).Count
            if ($numItems -eq $ddlGPO.CheckedItems.Count) {
                [System.Windows.Forms.MessageBox]::Show("Selected GPO(s) backed up successfully!", "Backup Complete", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Failed to back up some or all selected GPO(s)", "Backup Incomplete", "OK", "Error")
            }
        }
    })

    $GPOForm.Controls.AddRange(@($lblGPO, $ddlGPO, $btnGPO, $btnSelectAll))

    # Display the form
    $GPOForm.ShowDialog()

    # Clean up the form
    $GPOForm.Dispose()
}

function ShowImportForm {
    # Paste the content of importGPO2.ps1 here
    Add-Type -AssemblyName System.Windows.Forms

    $FormObject2=[System.Windows.Forms.Form]
    $LabelObject2=[System.Windows.Forms.Label]
    $CheckedListBoxObject2=[System.Windows.Forms.CheckedListBox]
    $ButtonObject2=[System.Windows.Forms.Button]

    $DefaultFont2='Times New Roman,12'

    # Set up base form #2
    $GPOForm2=New-Object $FormObject2
    $GPOForm2.ClientSize='700,200'
    $GPOForm2.Text='Import GPOs'
    $GPOForm2.BackColor='white'
    $GPOForm2.Font=$DefaultFont2
    $GPOForm2.WindowState = 'Maximized'  # Set the window state to Maximized

    # Building the form #2
    $lblGPO2=New-Object $LabelObject2
    $lblGPO2.Text='GPOs:'
    $lblGPO2.Autosize=$true
    $lblGPO2.Location=New-Object System.Drawing.Point(20,20)

    $ddlGPO2=New-Object $CheckedListBoxObject2
    $ddlGPO2.Size = New-Object System.Drawing.Size(950,600)
    $ddlGPO2.Location=New-Object System.Drawing.Point(70,20)

    $btnGPO2=New-Object $ButtonObject2
    $btnGPO2.Text='Import'
    $btnGPO2.AutoSize=$true
    $btnGPO2.Location=New-Object System.Drawing.Point(400,650)

    $btnSelectAll2=New-Object $ButtonObject2
    $btnSelectAll2.Text='Select All'
    $btnSelectAll2.AutoSize=$true
    $btnSelectAll2.Location=New-Object System.Drawing.Point(600,650)

    $dialog2 = New-Object System.Windows.Forms.FolderBrowserDialog 
    $dialog2.Description = "Go to location where backup folder(s) is/are saved"

    if ($dialog2.ShowDialog() -eq 'OK') {
        # Load the XML file for form #2
        $backupPath2 = $dialog2.SelectedPath+"\manifest.xml"
        $xml = [xml](Get-Content $backupPath2)
    }

    # Get all BackupInst nodes for form #2
    $backupInstNodes = $xml.Backups.BackupInst

    # Iterate through each BackupInst node and fetch the GPODisplayName for form #2
    foreach ($node in $backupInstNodes) {
        $displayName = $node.GPODisplayName.'#cdata-section'
        $ddlGPO2.Items.Add($displayName)
    }

    # Iterate through each BackupInst node and fetch the BackupID for form #2
    foreach ($node in $backupInstNodes) {
        $backupID = $node.ID.'#cdata-section'
        $backupID = $backupID.TrimStart('{').TrimEnd('}')
    }

    $btnSelectAll2.Add_Click({
        for ($i = 0; $i -lt $ddlGPO2.Items.Count; $i++) {
            $ddlGPO2.SetItemChecked($i, $true)
        }
    })

    $btnGPO2.Add_Click({ 
        $selectedItems2 = $ddlGPO2.CheckedItems | ForEach-Object {$_.ToString()} 
        Write-Host "Selected GPOs: $($selectedItems2 -join ', ')"

        $gpoHashtable = @{}
        Get-GPO -All | ForEach-Object {
            $gpoHashtable[$_.DisplayName] = $true
        }

        $successCount = 0
        foreach ($displayName in $selectedItems2) {
            if ($gpoHashtable.ContainsKey($displayName)) {
                $successCount++
                $result = [System.Windows.Forms.MessageBox]::Show("Duplicate GPO found: $($displayName) ... Do you still want to import it?", "Duplicate GPO", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-Host "$($displayName) successfully imported"
                    Import-GPO -ID $backupID -TargetName $displayName -CreateIfNeeded -Path $dialog2.SelectedPath
                } else {
                    Write-Host "$($displayName) not imported"
                }
            }
            else {
                Write-Host "$($displayName) successfully imported"
                Import-GPO -ID $backupID -TargetName $displayName -CreateIfNeeded -Path $dialog2.SelectedPath
            }
        }

        if ($successCount -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("All selected GPO(s) imported successfully!", "Import Complete", "OK", "Information")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Duplicate GPO(s) found!", "Import Complete", "OK", "Information") 
        }
    })


    Read-Host "Press Enter to continue..."


    $GPOForm2.Controls.AddRange(@($lblGPO2, $ddlGPO2, $btnGPO2, $btnSelectAll2))

    # Display the form #2
    $GPOForm2.ShowDialog()

    # Clean up the form #2
    $GPOForm2.Dispose()
}

function ShowLinkForm {
    # Paste the content of linkGPO.ps1 here
    Add-Type -AssemblyName System.Windows.Forms

    $FormObject=[System.Windows.Forms.Form]
    $LabelObject=[System.Windows.Forms.Label]
    $CheckedListBoxObject=[System.Windows.Forms.CheckedListBox]
    $ButtonObject=[System.Windows.Forms.Button]

    $DefaultFont='Times New Roman,12'

    # Set up base form
    $GPOForm=New-Object $FormObject
    $GPOForm.ClientSize='700,200'
    $GPOForm.Text='Link GPOs'
    $GPOForm.BackColor='white'
    $GPOForm.Font=$DefaultFont
    $GPOForm.WindowState = 'Maximized'  # Set the window state to Maximized

    # Building the form
    $lblGPO=New-Object $LabelObject
    $lblGPO.Text='GPOs:'
    $lblGPO.Autosize=$true
    $lblGPO.Location=New-Object System.Drawing.Point(20,20)

    $ddlGPO=New-Object $CheckedListBoxObject
    $ddlGPO.Size = New-Object System.Drawing.Size(950,300)
    $ddlGPO.Location=New-Object System.Drawing.Point(70,20)

    $lblGPO2=New-Object $LabelObject
    $lblGPO2.Text='OUs:'
    $lblGPO2.Autosize=$true
    $lblGPO2.Location=New-Object System.Drawing.Point(20,330)
 
    $ddlGPO2=New-Object $CheckedListBoxObject
    $ddlGPO2.Size = New-Object System.Drawing.Size(950,300)
    $ddlGPO2.Location=New-Object System.Drawing.Point(70,330)

    $btnGPO1=New-Object $ButtonObject
    $btnGPO1.Text='Link'
    $btnGPO1.AutoSize=$true
    $btnGPO1.Location=New-Object System.Drawing.Point(400,650)

    $btnGPO2=New-Object $ButtonObject
    $btnGPO2.Text='Unlink'
    $btnGPO2.AutoSize=$true
    $btnGPO2.Location=New-Object System.Drawing.Point(600,650)

    Get-GPO -all | ForEach-Object {$ddlGPO.Items.Add($_.DisplayName)}
    $domainDNSRoot = (Get-ADDomain).DNSRoot
    $ddlGPO2.Items.Add($domainDNSRoot)
    $OUs = Get-ADOrganizationalUnit -Filter * | Select-Object Name
    $OUs = ($OUs).Name
    foreach ($OU in $OUs) {
            $ddlGPO2.Items.Add($OU)
        }

    $btnGPO1.Add_Click({ 
        $selectedGPOs = $ddlGPO.CheckedItems 
        $selectedOU = $ddlGPO2.CheckedItems 
        Write-Host $ddlGPO2.CheckedItems 
        $distinguishedOU = Get-ADOrganizationalUnit -Filter "Name -eq '$selectedOU'" | Select-Object DistinguishedName 
        foreach ($gpo in $selectedGPOs) { 
            New-GPLink -Name $gpo -Target $distinguishedOU.DistinguishedName -Enforced Yes 
        } 
            [System.Windows.Forms.MessageBox]::Show("GPOs linked to selected OU successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }) 
        
    $btnGPO2.Add_Click({ 
        $selectedGPOs = $ddlGPO.CheckedItems 
        $selectedOU = $ddlGPO2.CheckedItems 
        Write-Host $ddlGPO2.CheckedItems 
        $distinguishedOU = Get-ADOrganizationalUnit -Filter "Name -eq '$selectedOU'" | Select-Object DistinguishedName 
        foreach ($gpo in $selectedGPOs) { 
            Remove-GPLink -Name $gpo -Target $distinguishedOU.DistinguishedName 
        } 
            [System.Windows.Forms.MessageBox]::Show("GPOs unlinked to selected OU successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) })
        
    $GPOForm.Controls.AddRange(@($lblGPO, $ddlGPO, $btnGPO1, $lblGPO2, $ddlGPO2, $btnGPO2))

    # Display the form
    $GPOForm.ShowDialog()

    # Clean up the form
    $GPOForm.Dispose()
}

# Display the main form
$MainForm.ShowDialog()