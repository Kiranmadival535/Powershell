Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
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



function AddServerNamesToGridViewAndCheck {
    param (
        [System.Windows.Forms.TextBox]$ServerTextBox,
        [System.Windows.Forms.DataGridView]$DataGridView,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $serverNames = $ServerTextBox.Lines | ForEach-Object { $_.Trim() }
    
    foreach ($serverName in $serverNames) {
        if (-not [string]::IsNullOrWhiteSpace($serverName)) {
            $row = $DataGridView.Rows[$DataGridView.Rows.Add()]
            $row.Cells["ServerName"].Value = $serverName
            
            try {
                # Attempt to retrieve services with credentials
                $services = Get-WmiObject -Class Win32_Service -ComputerName $serverName -Credential $Credential -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
                $comboBoxCell = New-Object System.Windows.Forms.DataGridViewComboBoxCell
                $comboBoxCell.Items.AddRange($services)
                $row.Cells["Services"] = $comboBoxCell
            } catch {
                $row.Cells["Services"] = New-Object System.Windows.Forms.DataGridViewComboBoxCell
                $row.Cells["Services"].Items.Add("Error retrieving services: " + $_.Exception.Message)
            }
        }
    }
}

function PopulateServicesDropdown {
    param (
        [System.Management.Automation.PSCredential]$Credential
    )
    
    foreach ($row in $dataGridView.Rows) {
        $selectedServer = $row.Cells["ServerName"].Value
        if (-not [string]::IsNullOrWhiteSpace($selectedServer)) {
            try {
                $services = Get-WmiObject -Class Win32_Service -ComputerName $selectedServer -Credential $Credential -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
                $comboBoxCell = New-Object System.Windows.Forms.DataGridViewComboBoxCell
                $comboBoxCell.Items.AddRange($services)
                $row.Cells["Services"] = $comboBoxCell
            } catch {
                $row.Cells["Services"] = New-Object System.Windows.Forms.DataGridViewComboBoxCell
                $row.Cells["Services"].Items.Add("Error retrieving services: " + $_.Exception.Message)
            }
        }
    }
}

function CheckServiceStatus {
    param (
        [System.Management.Automation.PSCredential]$Credential
    )
    
    foreach ($row in $dataGridView.Rows) {
        $serverName = $row.Cells["ServerName"].Value
        if (-not [string]::IsNullOrWhiteSpace($serverName)) {
            $serviceName = $row.Cells["Services"].Value
            if (-not [string]::IsNullOrWhiteSpace($serviceName)) {
                try {
                    $services = Get-WmiObject -Class Win32_Service -ComputerName $serverName -Credential $Credential -ErrorAction Stop | Where-Object { $_.DisplayName -eq $serviceName }
                    if ($services) {
                        $status = $services.State
                        $row.Cells["Status"].Value = $status
                    } else {
                        $row.Cells["Status"].Value = "Service '$serviceName' not found on server '$serverName'."
                    }
                } catch {
                    $row.Cells["Status"].Value = "Error: " + $_.Exception.Message
                }
            } else {
                $row.Cells["Status"].Value = "Service name is empty for server '$serverName'."
            }
        }
    }
}

function StartService {
    param (
        [string]$serverName,
        [string]$serviceName,  # This should be the display name
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Look for the service by display name
        $service = Get-WmiObject -Class Win32_Service -ComputerName $serverName -Credential $Credential -ErrorAction Stop | Where-Object { $_.DisplayName -eq $serviceName }
        if ($service) {
            if ($service.State -eq 'Stopped') {
                $service.StartService()  # Use StartService method
                [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' on server '$serverName' started successfully.")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' on server '$serverName' is already running.")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' not found on server '$serverName'.")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error starting service: $_")
    }
}


function StopService {
    param (
        [string]$serverName,
        [string]$serviceName,  # This should be the display name
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Look for the service by display name
        $service = Get-WmiObject -Class Win32_Service -ComputerName $serverName -Credential $Credential -ErrorAction Stop | Where-Object { $_.DisplayName -eq $serviceName }
        if ($service) {
            if ($service.State -eq 'Running') {
                $service.StopService()  # Use StopService method
                [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' on server '$serverName' stopped successfully.")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' on server '$serverName' is already stopped.")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Service '$serviceName' not found on server '$serverName'.")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error stopping service: $_")
    }
}

# Function to clear the DataGridView
function ClearTable {
    $dataGridView.Rows.Clear()
}

function Export-ToExcel {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView
    )

    # Check if Excel is running
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        # Create a new Excel application if it's not running
        $excel = New-Object -ComObject Excel.Application
    }
    
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)

    # Export column headers excluding last two columns
    $columnIndex = 1
    foreach ($header in $DataGridView.Columns) {
        if ($header.Name -notin @("Start Service", "Stop Service") -and $columnIndex -le ($DataGridView.Columns.Count - 2)) {
            $sheet.Cells.Item(1, $columnIndex) = $header.HeaderText
            $columnIndex++
        }
    }

    # Export data rows excluding rows with empty server names
    $rowIndex = 2
    foreach ($row in $DataGridView.Rows) {
        $serverName = $row.Cells["ServerName"].Value
        if (-not [string]::IsNullOrWhiteSpace($serverName)) {
            $columnIndex = 1
            foreach ($cell in $row.Cells) {
                # Exclude last two columns
                if ($DataGridView.Columns[$columnIndex - 1].Name -notin @("Start Service", "Stop Service") -and $columnIndex -le ($DataGridView.Columns.Count - 2)) {
                    if ($DataGridView.Columns[$columnIndex - 1].Name -eq "Status") {
                        $statusValue = $cell.Value
                        switch ($statusValue) {
                            1 { $statusString = "Stopped" }
                            4 { $statusString = "Running" }
                            default { $statusString = "Unknown" }
                        }
                        $sheet.Cells.Item($rowIndex, $columnIndex) = $statusString
                    } else {
                        $sheet.Cells.Item($rowIndex, $columnIndex) = $cell.Value
                    }
                }
                $columnIndex++
            }
            $rowIndex++
        }
    }

    # Autofit columns
    $range = $sheet.UsedRange
    if ($range -ne $null) {
        $range.EntireColumn.AutoFit() | Out-Null
    }

    # Get default save location and name using SaveFileDialog
    $saveFileDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    $saveFileDialog.Title = "Save Excel File"
    $saveFileDialog.ShowDialog() | Out-Null
    $filePath = $saveFileDialog.FileName

    if (-not [string]::IsNullOrWhiteSpace($filePath)) {
        # Save the workbook
        $workbook.SaveAs($filePath)
        Write-Host "Excel file saved to: $filePath"
    } else {
        Write-Host "No file selected. Excel data was not saved."
    }

    # Close Excel objects
    $workbook.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null

    # Safely quit Excel if there are no workbooks left
    if ($excel.Workbooks.Count -eq 0) {
        $excel.Quit()
    }
    
    # Clean up COM objects
    if ($excel -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Optionally, use this to remove the variable safely
    Remove-Variable -Name excel -ErrorAction SilentlyContinue
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
        # Open the file for writing
        $fileStream = [System.IO.StreamWriter]::new($filePath)

        # Write column headers to CSV, excluding last two columns
        $columnHeaders = @()
        foreach ($column in $DataGridView.Columns) {
            if ($column.Index -lt ($DataGridView.Columns.Count - 2)) {
                $columnHeaders += $column.HeaderText
            }
        }
        $fileStream.WriteLine(($columnHeaders -join ","))

        # Write data rows to CSV, excluding last two columns
        foreach ($row in $DataGridView.Rows) {
            $rowData = @()
            foreach ($cell in $row.Cells) {
                if ($cell.ColumnIndex -lt ($DataGridView.Columns.Count - 2)) {
                    $rowData += $cell.Value
                }
            }
            $fileStream.WriteLine(($rowData -join ","))
        }

        # Close the file stream
        $fileStream.Close()
        
        Write-Host "CSV file saved to: $filePath"
    } else {
        Write-Host "No file selected. Data was not saved."
    }
}



# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server Service Management Tool"
$form.Size = New-Object System.Drawing.Size(1000, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::CadetBlue
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Create DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 215)
$dataGridView.Location = New-Object System.Drawing.Point(200, 60)
$dataGridView.Anchor = "Top","Bottom", "Left", "Right"
$dataGridView.BackgroundColor = [System.Drawing.SystemColors]::InactiveCaption
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.RowHeadersVisible = $false
$dataGridView.AllowUserToAddRows = $False

# Configure DataGridView column headers
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::wheat
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$dataGridView.ColumnHeadersHeight = 25;




# Create a Font object with size 10 and bold style
$font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)

# Set the font for column headers
$dataGridView.ColumnHeadersDefaultCellStyle.Font = $font

# Add columns to the DataGridView
$serverNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$serverNameColumn.HeaderText = "Server Name"
$serverNameColumn.Name = "ServerName"
$serverNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serverNameColumn)

$serviceNameColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$serviceNameColumn.HeaderText = "Select The Service"
$serviceNameColumn.Name = "Services"
$serviceNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serviceNameColumn)

$serviceStatusColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$serviceStatusColumn.HeaderText = "Status"
$serviceStatusColumn.Name = "Status"
$serviceStatusColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serviceStatusColumn)

# Create Start button column
$startButtonColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn

$startButtonColumn.UseColumnTextForButtonValue = $true
$startButtonColumn.FlatStyle = 'Flat'  # Set FlatStyle to Flat
$startButtonColumn.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightSeaGreen
$startButtonColumn.Text = "START"
$startButtonColumn.DefaultCellStyle.Font = New-Object Drawing.Font("Microsoft Sans Serif", 10, [Drawing.FontStyle]::Bold)
$dataGridView.Columns.Add($startButtonColumn)


# Create Stop button column
$stopButtonColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn

$stopButtonColumn.UseColumnTextForButtonValue = $true
$stopButtonColumn.FlatStyle = 'Flat'  # Set FlatStyle to Flat
$stopButtonColumn.DefaultCellStyle.BackColor = [System.Drawing.Color]::IndianRed
$stopButtonColumn.Text = "STOP"
$stopButtonColumn.DefaultCellStyle.Font = New-Object Drawing.Font("Microsoft Sans Serif", 10, [Drawing.FontStyle]::Bold)

$dataGridView.Columns.Add($stopButtonColumn)

# Add CellClick event handler for the DataGridView
$dataGridView.add_CellClick({
    $rowIndex = $_.RowIndex
    $columnIndex = $_.ColumnIndex
    
    if ($columnIndex -eq $startButtonColumn.Index) {
        # Start button clicked
        $serverName = $dataGridView.Rows[$rowIndex].Cells["ServerName"].Value
        $serviceName = $dataGridView.Rows[$rowIndex].Cells["Services"].Value
        if (-not [string]::IsNullOrWhiteSpace($serverName) -and -not [string]::IsNullOrWhiteSpace($serviceName)) {
            StartService -serverName $serverName -serviceName $serviceName -Credential $credential
        }
    } elseif ($columnIndex -eq $stopButtonColumn.Index) {
        # Stop button clicked
        $serverName = $dataGridView.Rows[$rowIndex].Cells["ServerName"].Value
        $serviceName = $dataGridView.Rows[$rowIndex].Cells["Services"].Value
        if (-not [string]::IsNullOrWhiteSpace($serverName) -and -not [string]::IsNullOrWhiteSpace($serviceName)) {
            StopService -serverName $serverName -serviceName $serviceName -Credential $credential
        }
    }
})


# Add DataGridView to the form
$form.Controls.Add($dataGridView)

# Add server label and textbox
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(20, 60)
$serverLabel.Size = New-Object System.Drawing.Size(180,25)
$serverLabel.Text = "Enter Servers Name"
$serverLabel.Anchor = "TOP","LEFT"
$serverLabel.AutoSize = $false
$serverLabel.BackColor = [System.Drawing.Color]::Wheat
$serverLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$serverLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$serverLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($serverLabel)

$serverTextbox = New-Object System.Windows.Forms.TextBox
$serverTextbox.Location = New-Object System.Drawing.Point(20, 85)
$serverTextbox.Size = New-Object System.Drawing.Size(180, 190)
$serverTextbox.Anchor = "Left","Top","Bottom"
$serverTextbox.Multiline = $true
$serverTextbox.ScrollBars = "Vertical"
$serverTextbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$serverTextbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($serverTextbox)

# Add ADD SERVERS button
$buttonAddServers = New-Object System.Windows.Forms.Button
$buttonAddServers.Location = New-Object System.Drawing.Point(50, 300)
$buttonAddServers.Size = New-Object System.Drawing.Size(100, 30)
$buttonAddServers.Text = "ENTER"
$buttonAddServers.BackColor = [System.Drawing.Color]::Silver
$buttonAddServers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonAddServers.FlatAppearance.BorderSize = 3
$buttonAddServers.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$buttonAddServers.Anchor = "Bottom"
$buttonAddServers.Add_Click({
    $credential = Get-Credential "Enter credentials for remote servers"
    Save-Credentials -Credential $credential -FilePath $credentialFilePath
    AddServerNamesToGridViewAndCheck -ServerTextBox $serverTextbox -DataGridView $dataGridView -Credential $credential
    PopulateServicesDropdown -Credential $credential
})
$form.Controls.Add($buttonAddServers)

# Add CHECK STATUS button
$buttonCheckStatus = New-Object System.Windows.Forms.Button
$buttonCheckStatus.Location = New-Object System.Drawing.Point(300, 300)
$buttonCheckStatus.Size = New-Object System.Drawing.Size(100, 30)
$buttonCheckStatus.Text = "CHECK"
$buttonCheckStatus.BackColor = [System.Drawing.Color]::SteelBlue
$buttonCheckStatus.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCheckStatus.FlatAppearance.BorderSize = 3
$buttonCheckStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckStatus.Anchor = "Bottom"
# Example for button click events
$buttonCheckStatus.Add_Click({
    CheckServiceStatus -Credential $credential
})

$form.Controls.Add($buttonCheckStatus)

# Add EXIT button
$ExitButton = New-Object System.Windows.Forms.Button
$ExitButton.Location = New-Object System.Drawing.Point(700, 300)
$ExitButton.Size = New-Object System.Drawing.Size(100, 30)
$ExitButton.Text = "EXIT"
$ExitButton.BackColor = [System.Drawing.Color]::LightCoral
$ExitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ExitButton.FlatAppearance.BorderSize = 3
$ExitButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$ExitButton.Anchor = "Bottom"

$ExitButton.Add_Click({
    # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }

    $form.Close()
})

$form.Controls.Add($ExitButton)

# Add CLEAR button
$ClearButton = New-Object System.Windows.Forms.Button
$ClearButton.Location = New-Object System.Drawing.Point(500, 300)
$ClearButton.Size = New-Object System.Drawing.Size(100, 30)
$ClearButton.Text = "CLEAR"
$ClearButton.BackColor = [System.Drawing.SystemColors]::ButtonFace
$ClearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ClearButton.FlatAppearance.BorderSize = 3
$ClearButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$ClearButton.Anchor = "Bottom"

$ClearButton.Add_Click({
    # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }

    ClearTable -dataGridView $dataGridView
    $serverTextbox.Text = ""
    
})

$form.Controls.Add($ClearButton)

# Add event handler to update services dropdown when server name changes
$serverTextbox.add_TextChanged({
    PopulateServicesDropdown
})


## Create a label control for the watermark
$watermarkLabel = New-Object System.Windows.Forms.Label
$watermarkLabel.Text = "Created by Kiran Madival "
$watermarkLabel.AutoSize = $true
$watermarkLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Regular)
$watermarkLabel.BackColor = [System.Drawing.Color]::White
$watermarkLabel.Location = New-Object System.Drawing.Point(820, 495)

# Set anchor to bottom-right so it stays in place when resizing the form
$watermarkLabel.Anchor = "Bottom","Right"

# Add the watermark label to the form
$form.Controls.Add($watermarkLabel)



# Create a new MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# Set background color
$menuStrip.BackColor = [System.Drawing.SystemColors]::InactiveCaption

# Set font size
$font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9) # Adjust the font size as needed
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
    [System.Windows.Forms.MessageBox]::Show("Kiran created this application to check the serviceÂ status on the servers and control their start and stop actions.", "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Add the "About" submenu to the "Help" menu
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
$saveButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)

# Create Separator
$separator = New-Object System.Windows.Forms.ToolStripSeparator

# Create Print Button
$printButton = New-Object System.Windows.Forms.ToolStripButton
$printButton.Text = "Print"
$printButton.ToolTipText = "Print DataGridView"
$printButton.ForeColor = [System.Drawing.Color]::Black
$printButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)

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
                $font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
                $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
                $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black)

                # Validate if DataGridView is not null
                if ($null -eq $dataGridView) {
                    throw "DataGridView is null or not assigned."
                }

                $x = 50  # Initial x position
                $y = 50  # Initial y position

                # Define cell dimensions
                $cellWidth = 250
                $cellHeight = 25

                # Watermark text
                $watermarkText = "Made By Kiran Madival"
                $watermarkFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 15, [System.Drawing.FontStyle]::Regular)
                $watermarkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightGray)

                # Draw watermark
                $watermarkSize = $e.Graphics.MeasureString($watermarkText, $watermarkFont)
                $watermarkX = ($e.PageBounds.Width - $watermarkSize.Width) / 2
                $watermarkY = ($e.PageBounds.Height - $watermarkSize.Height) / 2
                $e.Graphics.DrawString($watermarkText, $watermarkFont, $watermarkBrush, $watermarkX, $watermarkY)

                # Draw table headers
                for ($col = 0; $col -lt ($dataGridView.ColumnCount - 2); $col++) {  # Exclude last two columns
                    $e.Graphics.DrawRectangle($pen, $x, $y, $cellWidth, $cellHeight)
                    $e.Graphics.DrawString($dataGridView.Columns[$col].HeaderText, $font, $brush, ($x + 5), ($y + 5))
                    $x += $cellWidth
                }

                $x = 50  # Reset x position
                $y += $cellHeight  # Move to the next row

                # Draw table data
                for ($row = 0; $row -lt $dataGridView.Rows.Count; $row++) {
                    for ($col = 0; $col -lt ($dataGridView.Columns.Count - 2); $col++) {  # Exclude last two columns
                        $cellValue = $dataGridView.Rows[$row].Cells[$col].Value
                        if ($cellValue -ne $null) {
                            $e.Graphics.DrawRectangle($pen, $x, $y, $cellWidth, $cellHeight)
                            $e.Graphics.DrawString($cellValue.ToString(), $font, $brush, ($x + 5), ($y + 5))
                        }
                        $x += $cellWidth
                    }
                    $x = 50  # Reset x position
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
$watermarkLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Regular)
$watermarkLabel.BackColor = [System.Drawing.Color]::White
$watermarkLabel.Location = New-Object System.Drawing.Point(810, 342)

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
