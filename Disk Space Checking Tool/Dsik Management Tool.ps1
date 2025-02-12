﻿Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Function to clear all columns except "ServerName" and clear rows
# Function to clear DataGridView


$credentialFilePath = "C:\Temp\credential.xml"  # Change this path as needed

function ClearTable {
    param (
        [System.Windows.Forms.DataGridView]$dataGridView
    )
    $dataGridView.Rows.Clear()
}



# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Drive Space checking Tool"
$form.Size = New-Object System.Drawing.Size(1000, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::DarkCyan
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D


$labelServerNames = New-Object System.Windows.Forms.Label
$labelServerNames.Location = New-Object System.Drawing.Point(20, 60)
$labelServerNames.Size = New-Object System.Drawing.Size(180,30)
$labelServerNames.Text = "Enter Servers Name"
$labelServerNames.Anchor = "TOP","LEFT"
$labelServerNames.AutoSize = $false
$labelServerNames.BackColor = [System.Drawing.Color]::LightSalmon
$labelServerNames.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelServerNames.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$labelServerNames.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelServerNames)

# Create a text box for entering server names
$textBoxServers = New-Object System.Windows.Forms.TextBox
$textBoxServers.Location = New-Object System.Drawing.Point(20,90)
$textBoxServers.Size = New-Object System.Drawing.Size(180, 320)
$textBoxServers.Anchor = "Left","Top","Bottom"
$textBoxServers.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$textBoxServers.Multiline = $true
$textBoxServers.ScrollBars = "Vertical"
$textBoxServers.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($textBoxServers)


# Create DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 350)
$dataGridView.Location = New-Object System.Drawing.Point(200, 60)
$dataGridView.Anchor = "Top","Bottom", "Left", "Right"
$dataGridView.BackgroundColor = [System.Drawing.SystemColors]::InactiveCaption
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.RowHeadersVisible = $false

# Configure DataGridView column headers
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightSalmon
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$dataGridView.ColumnHeadersHeight = 30;
$form.Controls.Add($dataGridView)
# Create a Font object with size 10 and bold style
$font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)

# Set the font for column headers
$dataGridView.ColumnHeadersDefaultCellStyle.Font = $font
# Center align all text in cells
$dataGridView.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

$columnServerName  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnServerName.HeaderText = "Server Name"
$columnServerName.Name = "ServerName"
$columnServerName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($columnServerName )

$columnDriveName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnDriveName.HeaderText = "Drive Name"
$columnDriveName.Name = "DriveName"
$columnDriveName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($columnDriveName)

$columnTotalSpace = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnTotalSpace.HeaderText = "Total Space (GB)"
$columnTotalSpace.Name = "TotalSpace"
$columnTotalSpace.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($columnTotalSpace)

$columnFreeSpace = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnFreeSpace.HeaderText = "Free Space (GB)"
$columnFreeSpace.Name = "FreeSpace"
$columnFreeSpace.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($columnFreeSpace)

$columnFreeSpacePercent = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnFreeSpacePercent.HeaderText = "Free Space (%)"
$columnFreeSpacePercent.Name = "FreeSpacePercent"
$columnFreeSpacePercent.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($columnFreeSpacePercent)    

# Add CHECK button
$button1 = New-Object System.Windows.Forms.Button
$button1.Location = New-Object System.Drawing.Point(280, 435)
$button1.Size = New-Object System.Drawing.Size(100, 30)
$button1.Text = "CHECK"
$button1.BackColor = [System.Drawing.SystemColors]::ActiveCaption
$button1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button1.FlatAppearance.BorderSize = 3
$button1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$button1.Anchor = "Bottom"
$button1.Add_Click({
    $credential = $null

    # Check if credentials file exists
    if (Test-Path $credentialFilePath) {
        $credential = Import-Clixml -Path $credentialFilePath
    } else {
        # Prompt for credentials and save them
        $credential = Get-Credential "Enter credentials for remote servers"
        $credential | Export-CliXml -Path $credentialFilePath
    }

    CheckDiskSpace -credential $credential
})
$form.Controls.Add($button1)





# Add CLEAR button
$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Point(480, 435)
$button2.Size = New-Object System.Drawing.Size(100, 30)
$button2.Text = "CLEAR"
$button2.BackColor = [System.Drawing.SystemColors]::ButtonFace
$button2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button2.FlatAppearance.BorderSize = 3
$button2.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$button2.Anchor = "Bottom"
$button2.Add_Click({
    ClearTable -dataGridView $dataGridView
    
    # Additional reset logic
    $textBoxServers.Text = ""
    # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }
    
})

$form.Controls.Add($button2)

# Add EXIT button
$button3 = New-Object System.Windows.Forms.Button
$button3.Location = New-Object System.Drawing.Point(680, 435)
$button3.Size = New-Object System.Drawing.Size(100, 30)
$button3.Text = "EXIT"
$button3.BackColor = [System.Drawing.Color]::LightCoral
$button3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button3.FlatAppearance.BorderSize = 3
$button3.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$button3.Anchor = "Bottom"

$button3.Add_Click({
     # Remove the credentials file
    if (Test-Path $credentialFilePath) {
        Remove-Item $credentialFilePath -Force
    }
    $form.Close()
})


$form.Controls.Add($button3)

# Function to check disk space
function CheckDiskSpace {
    param (
        [System.Management.Automation.PSCredential]$credential
    )

    $servers = $textBoxServers.Lines | Where-Object { $_ -match '\S' }  # Remove empty lines

    foreach ($server in $servers) {
        Write-Host "Checking disk space on $server..."

        try {
            $disks = Get-WmiObject -ComputerName $server -Class Win32_LogicalDisk -Filter "DriveType=3" -Credential $credential -ErrorAction Stop

            $serverNameShown = $false  # Reset flag for each server

            foreach ($disk in $disks) {
                $driveName = $disk.DeviceID
                $totalSpaceGB = [math]::Round($disk.Size / 1GB, 2)
                $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                $freeSpacePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)

                # Add row for each drive, show server name only once
                if (-not $serverNameShown) {
                    $rowIndex = $dataGridView.Rows.Add()
                    $dataGridView.Rows[$rowIndex].Cells["ServerName"].Value = $server
                    $serverNameShown = $true
                }
                else {
                    $rowIndex = $dataGridView.Rows.Add()
                    $dataGridView.Rows[$rowIndex].Cells["ServerName"].Value = ""
                }

                $dataGridView.Rows[$rowIndex].Cells["DriveName"].Value = $driveName
                $dataGridView.Rows[$rowIndex].Cells["TotalSpace"].Value = $totalSpaceGB
                $dataGridView.Rows[$rowIndex].Cells["FreeSpace"].Value = $freeSpaceGB
                $dataGridView.Rows[$rowIndex].Cells["FreeSpacePercent"].Value = $freeSpacePercent
            }
        }
        catch {
            Write-Host "Failed to retrieve disk information from $server : $_" -ForegroundColor Red
        }
    }
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
    [System.Windows.Forms.MessageBox]::Show("Kiran created this tool to Check the disk space on the servers.", "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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
                $cellWidth = 150
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


# 
# Show the form
$form.ShowDialog()
