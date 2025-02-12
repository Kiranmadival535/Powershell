﻿# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Function to save credentials to a file
function Save-Credentials {
    param(
        [System.Management.Automation.PSCredential]$Credential,
        [string]$FilePath
    )
    $Credential | Export-Clixml -Path $FilePath
}
$credentialFilePath = "C:\Temp\credential.xml"  # Change this path as needed

# Function to load credentials from a file
function Load-Credentials {
    param(
        [string]$FilePath
    )
    if (Test-Path $FilePath) {
        return Import-Clixml -Path $FilePath
    } else {
        return $null
    }
}

function Get-AdminsAndUsers {
    param(
        [string]$Server,
        [System.Management.Automation.PSCredential]$Credential  # Add this line
    )

    # Get local administrators on the server
    $admins = Invoke-Command -ComputerName $Server -Credential $Credential -ScriptBlock {
        $admins = Get-LocalGroupMember -Group Administrators | Select-Object -ExpandProperty Name
        $adminsRoles = foreach ($admin in $admins) {
            [PSCustomObject]@{
                Role = "Administrator"
                Name = $admin
            }
        }
        $users = Get-LocalGroupMember -Group 'Remote Desktop Users' | Select-Object -ExpandProperty Name
        $usersRoles = foreach ($user in $users) {
            [PSCustomObject]@{
                Role = "Remote Desktop User"
                Name = $user
            }
        }
        return $adminsRoles + $usersRoles
    }
    return $admins
}



# Function to clear DataGridView
function ClearTable {
    param (
        [System.Windows.Forms.DataGridView]$dataGridView
    )
    $dataGridView.Rows.Clear()
}



# Define the function to export DataGridView data to CSV, excluding last two columns
function Export-ToCSV {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView
    )

    # Get default save location and name using SaveFileDialog
    $saveFileDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV File"
    $saveFileDialog.ShowDialog() | Out-Null
    $filePath = $saveFileDialog.FileName

    if (-not [string]::IsNullOrWhiteSpace($filePath)) {
        try {
            # Open the file for writing
            $fileStream = [System.IO.StreamWriter]::new($filePath)

            # Write column headers to CSV
            $columnHeaders = @()
            foreach ($column in $DataGridView.Columns) {
                $columnHeaders += '"' + $column.HeaderText.Replace('"', '""') + '"'
            }
            $fileStream.WriteLine(($columnHeaders -join ","))

            # Write data rows to CSV
            foreach ($row in $DataGridView.Rows) {
                $rowData = @()
                foreach ($cell in $row.Cells) {
                    $cellValue = $cell.Value
                    if ($cellValue -is [string]) {
                        $cellValue = '"' + $cellValue.Replace('"', '""') + '"'
                    }
                    $rowData += $cellValue
                }
                $fileStream.WriteLine(($rowData -join ","))
            }

            # Close the file stream
            $fileStream.Close()
            
            Write-Host "CSV file saved to: $filePath"
        }
        catch {
            Write-Host "Error occurred while saving CSV file: $_"
        }
    } else {
        Write-Host "No file selected. Data was not saved."
    }
}

# Function to export DataGridView data to Excel
function Export-ToExcel {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView
    )

    try {
        # Create Excel application object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false

        # Create a SaveFileDialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $saveFileDialog.Title = "Save Excel File"
        
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $saveFileDialog.FileName

            # Add a workbook
            $workbook = $excel.Workbooks.Add()

            # Get the first worksheet
            $worksheet = $workbook.Worksheets.Item(1)

            # Export column headers
            for ($col = 0; $col -lt $DataGridView.ColumnCount; $col++) {
                $worksheet.Cells.Item(1, $col + 1) = $DataGridView.Columns[$col].HeaderText
            }

            # Export data rows
            for ($row = 0; $row -lt $DataGridView.Rows.Count; $row++) {
                for ($col = 0; $col -lt $DataGridView.Columns.Count; $col++) {
                    $worksheet.Cells.Item($row + 2, $col + 1) = $DataGridView.Rows[$row].Cells[$col].Value
                }
            }

            # Auto-resize columns
            $range = $worksheet.UsedRange
            $range.EntireColumn.AutoFit() | Out-Null

            # Save and close the workbook without prompting
            $workbook.SaveAs($FilePath)
            $workbook.Close($true)
            $excel.Quit()

            Write-Output "Exported data to $FilePath"
        }
    }
    catch {
        Write-Error "Error occurred: $_"
    }
}


# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server Access Management Tool"
$form.Size = New-Object System.Drawing.Size(1000, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::DarkSlateGray
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Create DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 250)
$dataGridView.Location = New-Object System.Drawing.Point(200, 50)
$dataGridView.Anchor = "Top","Bottom", "Left", "Right"
$dataGridView.BackgroundColor = [System.Drawing.SystemColors]::InactiveCaption
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.RowHeadersVisible = $false

# Configure DataGridView column headers
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::BurlyWood
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$dataGridView.ColumnHeadersHeight = 25;

# Create a Font object with size 10 and bold style
$font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

# Set the font for column headers
$dataGridView.ColumnHeadersDefaultCellStyle.Font = $font

# Add columns to the DataGridView
$serverNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$serverNameColumn.HeaderText = "Server Name"
$serverNameColumn.Name = "ServerName"
$serverNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serverNameColumn)


$roleColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$roleColumn.HeaderText = "Role"
$roleColumn.Name = "Role"
$roleColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($roleColumn)

$usersColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$usersColumn.HeaderText = "Users"
$usersColumn.Name = "UsersList"
$usersColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($usersColumn)

# Center align all text in cells
$dataGridView.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

# Add DataGridView to the form
$form.Controls.Add($dataGridView)

# Add server label and textbox
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(20, 50)
$serverLabel.Size = New-Object System.Drawing.Size(180,25)
$serverLabel.Text = "Enter Servers Name"
$serverLabel.Anchor = "TOP","LEFT"
$serverLabel.AutoSize = $false
$serverLabel.BackColor = [System.Drawing.Color]::BurlyWood
$serverLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$serverLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$serverLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($serverLabel)

$serverTextbox = New-Object System.Windows.Forms.TextBox
$serverTextbox.Location = New-Object System.Drawing.Point(20, 75)
$serverTextbox.Size = New-Object System.Drawing.Size(180, 225)
$serverTextbox.Anchor = "Left","Top","Bottom"
$serverTextbox.Multiline = $true
$serverTextbox.ScrollBars = "Vertical"
$serverTextbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$serverTextbox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($serverTextbox)

## Create a label control for the watermark
$watermarkLabel = New-Object System.Windows.Forms.Label
$watermarkLabel.Text = "Created by Kiran Madival "
$watermarkLabel.AutoSize = $true
$watermarkLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)
$watermarkLabel.BackColor = [System.Drawing.Color]::White
$watermarkLabel.Location = New-Object System.Drawing.Point(850, 495)

# Set anchor to bottom-right so it stays in place when resizing the form
$watermarkLabel.Anchor = "Bottom","Right"

# Add CHECK button
$button1 = New-Object System.Windows.Forms.Button
$button1.Location = New-Object System.Drawing.Point(250, 310)
$button1.Size = New-Object System.Drawing.Size(100, 30)
$button1.Text = "CHECK"
$button1.BackColor = [System.Drawing.SystemColors]::ActiveCaption
$button1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button1.FlatAppearance.BorderSize = 3
$button1.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$button1.Anchor = "Bottom"

$button1.Add_Click({
    $credential = Load-Credentials -FilePath $credentialFilePath

    if (-not $credential) {
        $credential = Get-Credential
        Save-Credentials -Credential $credential -FilePath $credentialFilePath
    }

    $serverNames = $serverTextbox.Text -split "`r`n" | ForEach-Object { $_.Trim() }
    ClearTable -dataGridView $dataGridView
    
    foreach ($server in $serverNames) {
        $adminsAndUsers = Get-AdminsAndUsers -Server $server -Credential $credential
        
        $serverNameAdded = $false
        
        foreach ($entry in $adminsAndUsers) {
            if (-not $serverNameAdded) {
                $dataGridView.Rows.Add($server, $entry.Role, $entry.Name)
                $serverNameAdded = $true
            } else {
                $dataGridView.Rows.Add("", $entry.Role, $entry.Name)
            }
        }
    }
})


$form.Controls.Add($button1)

# Add CLEAR button
$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Point(380, 310)
$button2.Size = New-Object System.Drawing.Size(100, 30)
$button2.Text = "CLEAR"
$button2.BackColor = [System.Drawing.SystemColors]::ButtonFace
$button2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button2.FlatAppearance.BorderSize = 3
$button2.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$button2.Anchor = "Bottom"
$button2.Add_Click({
    ClearTable -dataGridView $dataGridView
    $serverTextbox.Text = ""
    $textboxServers.Text = ""
    $textboxUsers.Text = ""
    

    # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }
})


$form.Controls.Add($button2)

# Add EXIT button
$button3 = New-Object System.Windows.Forms.Button
$button3.Location = New-Object System.Drawing.Point(510, 310)
$button3.Size = New-Object System.Drawing.Size(100, 30)
$button3.Text = "EXIT"
$button3.BackColor = [System.Drawing.Color]::LightCoral
$button3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button3.FlatAppearance.BorderSize = 3
$button3.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$button3.Anchor = "Bottom"

$button3.Add_Click({
    # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }

    $form.Close()
})


$form.Controls.Add($button3)

$textboxserverLabel = New-Object System.Windows.Forms.Label
$textboxserverLabel.Location = New-Object System.Drawing.Point(20, 355)
$textboxserverLabel.Size = New-Object System.Drawing.Size(180,25)
$textboxserverLabel.Text = "Enter Servers Name"
$textboxserverLabel.Anchor = "BOTTOM"
$textboxserverLabel.AutoSize = $false
$textboxserverLabel.BackColor = [System.Drawing.Color]::BurlyWood
$textboxserverLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$textboxserverLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$textboxserverLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($textboxserverLabel)


# Multiline text box for server names
$textboxServers = New-Object System.Windows.Forms.TextBox
$textboxServers.Multiline = $true
$textboxServers.Anchor  = "BOTTOM" 
$textboxServers.Size = New-Object System.Drawing.Size(180, 100)
$textboxServers.Location = New-Object System.Drawing.Point(20, 380)
$form.Controls.Add($textboxServers)

$textboxuserLabel = New-Object System.Windows.Forms.Label
$textboxuserLabel.Location = New-Object System.Drawing.Point(220, 355)
$textboxuserLabel.Size = New-Object System.Drawing.Size(180,25)
$textboxuserLabel.Text = "Enter Users Name"
$textboxuserLabel.Anchor = "BOTTOM"
$textboxuserLabel.AutoSize = $false
$textboxuserLabel.BackColor = [System.Drawing.Color]::BurlyWood
$textboxuserLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$textboxuserLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$textboxuserLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($textboxuserLabel)

# Multiline text box for user names
$textboxUsers = New-Object System.Windows.Forms.TextBox
$textboxUsers.Multiline = $true
$textboxUsers.Anchor  = "BOTTOM"
$textboxUsers.Size = New-Object System.Drawing.Size(180, 100)
$textboxUsers.Location = New-Object System.Drawing.Point(220, 380)
$form.Controls.Add($textboxUsers)

$DomainselectionLabel = New-Object System.Windows.Forms.Label
$DomainselectionLabel.Location = New-Object System.Drawing.Point(420, 355)
$DomainselectionLabel.Size = New-Object System.Drawing.Size(180,25)
$DomainselectionLabel.Text = "Select the Domain"
$DomainselectionLabel.Anchor = "BOTTOM"
$DomainselectionLabel.AutoSize = $false
$DomainselectionLabel.BackColor = [System.Drawing.Color]::BurlyWood
$DomainselectionLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$DomainselectionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$DomainselectionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($DomainselectionLabel)

# Dropdown box for domain selection
$dropdownDomain = New-Object System.Windows.Forms.ComboBox
$dropdownDomain.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dropdownDomain.Items.Add("TW")
$dropdownDomain.Items.Add("WMAD")
$dropdownDomain.Items.Add("MGTDMZ")
$dropdownDomain.Items.Add("DOMAIN0")
$dropdownDomain.Items.Add("DOMAIN1")
$dropdownDomain.Items.Add("TESTEIS")
$dropdownDomain.Anchor = "BOTTOM"
$dropdownDomain.Size = New-Object System.Drawing.Size(180, 20)
$dropdownDomain.Location = New-Object System.Drawing.Point(420, 380)
#$dropdownDomain.SelectedIndex = 0  # Select the first domain by default
$form.Controls.Add($dropdownDomain)

$acessselectionLabel = New-Object System.Windows.Forms.Label
$acessselectionLabel.Location = New-Object System.Drawing.Point(620, 355)
$acessselectionLabel.Size = New-Object System.Drawing.Size(180,25)
$acessselectionLabel.Text = "Select the level of access"
$acessselectionLabel.Anchor = "BOTTOM"
$acessselectionLabel.AutoSize = $false
$acessselectionLabel.BackColor = [System.Drawing.Color]::BurlyWood
$acessselectionLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$acessselectionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$acessselectionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($acessselectionLabel)
# Dropdown box for access level
$dropdownAccess = New-Object System.Windows.Forms.ComboBox
$dropdownAccess.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dropdownAccess.Items.Add("Admin")
$dropdownAccess.Items.Add("RDP")
$dropdownAccess.Anchor = "BOTTOM"
$dropdownAccess.Size = New-Object System.Drawing.Size(180, 20)
$dropdownAccess.Location = New-Object System.Drawing.Point(620, 380)
$form.Controls.Add($dropdownAccess)


# Button for adding access
$buttonAdd = New-Object System.Windows.Forms.Button
$buttonAdd.Text = "Add"
$buttonAdd.Size = New-Object System.Drawing.Size(80, 30)
$buttonAdd.Location = New-Object System.Drawing.Point(840, 380)
$buttonAdd.Anchor = "BOTTOM"
$buttonAdd.BackColor = [System.Drawing.Color]::MediumSeaGreen
$buttonAdd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonAdd.FlatAppearance.BorderSize = 3
$buttonAdd.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$buttonAdd.Add_Click({
    $credential = Load-Credentials -FilePath $credentialFilePath

    if (-not $credential) {
        $credential = Get-Credential
        Save-Credentials -Credential $credential -FilePath $credentialFilePath
    }

    $results = Execute-Actions -actionType "Add" -Credential $credential
    [System.Windows.Forms.MessageBox]::Show($results -join "`r`n", "Add Access Results")
})


$form.Controls.Add($buttonAdd)

# Button for removing access
$buttonRemove = New-Object System.Windows.Forms.Button
$buttonRemove.Text = "Remove"
$buttonRemove.Size = New-Object System.Drawing.Size(80, 30)
$buttonRemove.Location = New-Object System.Drawing.Point(840, 420)
$buttonRemove.Anchor = "BOTTOM"
$buttonRemove.BackColor = [System.Drawing.Color]::Coral
$buttonRemove.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRemove.FlatAppearance.BorderSize = 3
$buttonRemove.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$buttonRemove.Add_Click({
    $credential = Load-Credentials -FilePath $credentialFilePath

    if (-not $credential) {
        $credential = Get-Credential
        Save-Credentials -Credential $credential -FilePath $credentialFilePath
    }

    $results = Execute-Actions -actionType "Remove" -Credential $credential
    [System.Windows.Forms.MessageBox]::Show($results -join "`r`n", "Remove Access Results")
})


$form.Controls.Add($buttonRemove)

function Execute-Actions {
    param(
        [string]$actionType,
        [System.Management.Automation.PSCredential]$Credential  # Add this line
    )

    # Collect results
    $results = @()

    $servers = $textboxServers.Text -split "`r?`n" | Where-Object { $_ -match '\S' }
    $users = $textboxUsers.Text -split "`r?`n" | Where-Object { $_ -match '\S' }
    $accessLevel = $dropdownAccess.SelectedItem.ToString()
    $selectedDomain = $dropdownDomain.SelectedItem.ToString()

    foreach ($server in $servers) {
        foreach ($user in $users) {
            # Construct the full username with selected domain
            $fullUsername = "$selectedDomain\$user"

            # Define command based on action type
            switch ($actionType) {
                "Add" {
                    if ($accessLevel -eq "Admin") {
                        $command = "Add-LocalGroupMember -Group 'Administrators' -Member $fullUsername"
                    }
                    elseif ($accessLevel -eq "RDP") {
                        $command = "Add-LocalGroupMember -Group 'Remote Desktop Users' -Member $fullUsername"
                    }
                    $actionVerb = "added"
                }
                "Remove" {
                    if ($accessLevel -eq "Admin") {
                        $command = "Remove-LocalGroupMember -Group 'Administrators' -Member $fullUsername"
                    }
                    elseif ($accessLevel -eq "RDP") {
                        $command = "Remove-LocalGroupMember -Group 'Remote Desktop Users' -Member $fullUsername"
                    }
                    $actionVerb = "removed"
                }
                Default {
                    $results += "Failed: Invalid action type '$actionType'."
                    continue
                }
            }

            if ($command) {
                try {
                    # Execute the command on the server
                    Invoke-Command -ComputerName $server -Credential $Credential -ScriptBlock {
                        param($cmd)
                        Invoke-Expression $cmd
                    } -ArgumentList $command -ErrorAction Stop

                    $results += "Success: $accessLevel access $actionVerb for $fullUsername on $server."
                } catch {
                    $results += "Failed: $accessLevel access for $fullUsername on $server. Error: $_"
                }
            }
        }
    }

    return $results
}



# Create a new MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# Set background color
$menuStrip.BackColor = [System.Drawing.SystemColors]::ActiveCaption

# Set font size
$font = New-Object System.Drawing.Font("Segoe UI", 9) # Adjust the font size as needed
$menuStrip.Font = $font

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$editMenu.Text = "Edit"

$toolBoxMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolBoxMenu.Text = "Toolbox"

$tool1MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$tool1MenuItem.Text = "Tool 1"

$tool2MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$tool2MenuItem.Text = "Tool 2"

$toolBoxMenu.DropDownItems.AddRange(@($tool1MenuItem, $tool2MenuItem))

$fileSubMenu1 = New-Object System.Windows.Forms.ToolStripMenuItem
$fileSubMenu1.Text = "New"
$fileSubMenu2 = New-Object System.Windows.Forms.ToolStripMenuItem
$fileSubMenu2.Text = "Open"
$fileSubMenu3 = New-Object System.Windows.Forms.ToolStripMenuItem
$fileSubMenu3.Text = "Save"
$fileSubMenu3.Add_Click({
    Export-ToCSV $dataGridView
})


$fileMenu.DropDownItems.AddRange(@($fileSubMenu1, $fileSubMenu2, $fileSubMenu3))

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"

# Create the "About" submenu
$helpSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpSubMenu.Text = "About"

# Add event handler for the "About" submenu click event
$helpSubMenu.Add_Click({
    # Display a message box with the note or additional information
    [System.Windows.Forms.MessageBox]::Show("Kiran created this tool. Check user access information and add or removeÂ access to the servers.", "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Add the "About" submenu to the "Help" menu
$helpMenu.DropDownItems.Add($helpSubMenu)


$helpMenu.DropDownItems.Add($helpSubMenu)

$menuStrip.Items.AddRange(@($fileMenu, $editMenu, $toolBoxMenu, $helpMenu))

# Add MenuStrip to Form
$form.Controls.Add($menuStrip)

# Create a new ToolStrip
$toolStrip = New-Object System.Windows.Forms.ToolStrip

# Set background color
$toolStrip.BackColor = [System.Drawing.SystemColors]::ActiveCaption
        

# Create Save Button
$saveButton = New-Object System.Windows.Forms.ToolStripButton
$saveButton.Text = "Save"
$saveButton.ToolTipText = "Save DataGridView to Excel"
$saveButton.ForeColor = [System.Drawing.Color]::Black
$saveButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)

# Create Separator
$separator = New-Object System.Windows.Forms.ToolStripSeparator

# Create Print Button
$printButton = New-Object System.Windows.Forms.ToolStripButton
$printButton.Text = "Print"
$printButton.ToolTipText = "Print DataGridView"
$printButton.ForeColor = [System.Drawing.Color]::Black
$printButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)

# Add Items to ToolStrip
$toolStrip.Items.AddRange(@($saveButton, $separator, $printButton))


# Add ToolStrip to Form
$form.Controls.Add($toolStrip)

# Add Click Event Handlers for Save and Print Buttons
$saveButton.Add_Click({
   Export-ToExcel -DataGridView $dataGridView
})

$printButton.Add_Click({
    try {
        $printDialog = New-Object System.Windows.Forms.PrintDialog
        if ($printDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $printDocument = New-Object System.Drawing.Printing.PrintDocument
            $printDocument.PrinterSettings = $printDialog.PrinterSettings

            $printDocument.add_PrintPage({
                param($sender, $e)
                $font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)
                $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
                $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black)

                # Validate if DataGridView is not null
                if ($null -eq $dataGridView) {
                    throw "DataGridView is null or not assigned."
                }

                $x = 20  # Initial x position
                $y = 20  # Initial y position

                # Define cell dimensions
                $cellWidth = 250
                $cellHeight = 20

                # Watermark text
                $watermarkText = "Made By Kiran Madival"
                $watermarkFont = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Regular)
                $watermarkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightGray)

                # Draw watermark
                $watermarkSize = $e.Graphics.MeasureString($watermarkText, $watermarkFont)
                $watermarkX = ($e.PageBounds.Width - $watermarkSize.Width) / 2
                $watermarkY = ($e.PageBounds.Height - $watermarkSize.Height) / 2
                $e.Graphics.DrawString($watermarkText, $watermarkFont, $watermarkBrush, $watermarkX, $watermarkY)

                # Draw table headers
                for ($col = 0; $col -lt $dataGridView.ColumnCount; $col++) {
                    $e.Graphics.DrawRectangle($pen, $x, $y, $cellWidth, $cellHeight)
                    $e.Graphics.DrawString($dataGridView.Columns[$col].HeaderText, $font, $brush, ($x + 5), ($y + 5))
                    $x += $cellWidth
                }

                $x = 20  # Reset x position
                $y += $cellHeight  # Move to the next row

                # Draw table data
                for ($row = 0; $row -lt $dataGridView.Rows.Count; $row++) {
                    for ($col = 0; $col -lt $dataGridView.Columns.Count; $col++) {
                        $cellValue = $dataGridView.Rows[$row].Cells[$col].Value
                        if ($cellValue -ne $null) {
                            $e.Graphics.DrawRectangle($pen, $x, $y, $cellWidth, $cellHeight)
                            $e.Graphics.DrawString($cellValue.ToString(), $font, $brush, ($x + 5), ($y + 5))
                        }
                        $x += $cellWidth
                    }
                    $x = 20  # Reset x position
                    $y += $cellHeight  # Move to the next row
                }
            })

            $printDocument.Print()
        }
    } catch {
        Write-Error "An error occurred while printing: $_"
    }
})






## Create a label control for the watermark
$watermarkLabel = New-Object System.Windows.Forms.Label
$watermarkLabel.Text = "Created By : Kiran Madival "
$watermarkLabel.AutoSize = $true
$watermarkLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$watermarkLabel.BackColor = [System.Drawing.Color]::White
$watermarkLabel.Location = New-Object System.Drawing.Point(800, 492)

# Set anchor to bottom-right so it stays in place when resizing the form
$watermarkLabel.Anchor = "Bottom","Right"

# Add the watermark label to the form
$form.Controls.Add($watermarkLabel)

# Create StatusStrip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = "Bottom"
$statusStrip.BackColor = [System.Drawing.Color]::White
$statusStrip.Height = 20
$form.Controls.Add($statusStrip)


# Show the form
$form.ShowDialog() | Out-Null
