# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Function to add server names to the DataGridView and perform checks
function AddServerNamesToGridViewAndCheck {
    param (
        [System.Windows.Forms.TextBox]$serverTextbox,
        [System.Windows.Forms.DataGridView]$dataGridView,
        [System.Windows.Forms.ToolStripProgressBar]$toolStripProgressBar,
        [System.Windows.Forms.StatusStrip]$statusStrip
    )
    
    # Get server names from the text box
    $serverNames = $serverTextbox.Lines | ForEach-Object { $_.Trim() }

    # Reset progress bar
    $toolStripProgressBar.Value = 0

    # Count the total number of servers to check
    $totalServers = ($serverNames | Where-Object { $_ -match '\S' }).Count
    $serversChecked = 0

# Iterate through each server name
foreach ($serverName in $serverNames) {
    if (-not [string]::IsNullOrWhiteSpace($serverName)) {
        # Add a new row for the server
        $row = $dataGridView.Rows[$dataGridView.Rows.Add()]
        $row.Cells["ServerName"].Value = $serverName

        # Perform server checks
        if ($serverName) {
            $ipResult = nslookup $serverName
            $ipAddresses = $ipResult -split "`n" | Where-Object { $_ -like "Address:*" } | ForEach-Object { ($_ -split ":")[1].Trim() }
            $ipAddress = if ($ipAddresses.Count -gt 1) { $ipAddresses[1] } else { "Not Found" }
            $domainName = ($ipResult | Select-String -Pattern "Name:" | Select-Object -First 1).ToString().Split(":")[1].Trim() -replace "$serverName\.", ""
            $domainName = if (-not $domainName) { "Not Found" } else { $domainName }

            $pingResult = $null
            if ($ipAddress -ne "Not Found") {
                $pingResult = Test-Connection -ComputerName $serverName -Count 1 -Quiet
            }

            if ($pingResult) {
                $row.Cells["IPDetails"].Value = $ipAddress
                $row.Cells["DomainName"].Value = $domainName
                $row.Cells["PingResult"].Value = "Pinging"
                $row.Cells["PingResult"].Style.BackColor = [System.Drawing.Color]::Green
            } else {
                $row.Cells["IPDetails"].Value = $ipAddress
                $row.Cells["DomainName"].Value = $domainName
                $row.Cells["PingResult"].Value = "Not Pinging"
                $row.Cells["PingResult"].Style.BackColor = [System.Drawing.Color]::Red
            }
        }

        # Update progress bar
        $serversChecked++
        $progress = [math]::Min(($serversChecked / $totalServers) * 100, 100)
        $toolStripProgressBar.Value = $progress
    }
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



# Function to clear DataGridView
function ClearTable {
    param (
        [System.Windows.Forms.DataGridView]$dataGridView
    )
    $dataGridView.Rows.Clear()
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ping Testing Tool"
$form.Size = New-Object System.Drawing.Size(1000, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightSlateGray
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

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
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::BurlyWood
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$dataGridView.ColumnHeadersHeight = 25;

# Create a Font object with size 9 and bold style
$font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)

# Set the font for column headers
$dataGridView.ColumnHeadersDefaultCellStyle.Font = $font

# Add columns to the DataGridView
$serverNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$serverNameColumn.HeaderText = "Server Name"
$serverNameColumn.Name = "ServerName"
$serverNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serverNameColumn)

$ipDetailsColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ipDetailsColumn.HeaderText = "IP Address"
$ipDetailsColumn.Name = "IPDetails"
$ipDetailsColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($ipDetailsColumn)

$domainNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$domainNameColumn.HeaderText = "Domain Name"
$domainNameColumn.Name = "DomainName"
$domainNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($domainNameColumn)

$pingResultColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$pingResultColumn.HeaderText = "Ping Result"
$pingResultColumn.Name = "PingResult"
$pingResultColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($pingResultColumn)

# Center align all text in cells
$dataGridView.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

# Add DataGridView to the form
$form.Controls.Add($dataGridView)

# Add server label and textbox
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(20, 60)
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
$serverTextbox.Location = New-Object System.Drawing.Point(20, 85)
$serverTextbox.Size = New-Object System.Drawing.Size(180, 325)
$serverTextbox.Anchor = "Left","Top","Bottom"
$serverTextbox.Multiline = $true
$serverTextbox.ScrollBars = "Vertical"
$serverTextbox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($serverTextbox)

## Create a label control for the watermark
$watermarkLabel = New-Object System.Windows.Forms.Label
$watermarkLabel.Text = "Created by : Kiran Madival "
$watermarkLabel.AutoSize = $true
$watermarkLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$watermarkLabel.BackColor = [System.Drawing.Color]::White
$watermarkLabel.Location = New-Object System.Drawing.Point(820, 492)

# Set anchor to bottom-right so it stays in place when resizing the form
$watermarkLabel.Anchor = "Bottom","Right"

# Add the watermark label to the form
$form.Controls.Add($watermarkLabel)



# Create a new MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# Set background color
$menuStrip.BackColor = [System.Drawing.SystemColors]::InactiveCaption

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

$fileMenu.DropDownItems.AddRange(@($fileSubMenu1, $fileSubMenu2, $fileSubMenu3))

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"

# Create the "About" submenu
$helpSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpSubMenu.Text = "About"

# Add event handler for the "About" submenu click event
$helpSubMenu.Add_Click({
    # Display a message box with the note or additional information
    [System.Windows.Forms.MessageBox]::Show("Kiran made this application to check ping status and gather information about the server's IP address and domain name.", "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
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
                $font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
                $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
                $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black)

                # Validate if DataGridView is not null
                if ($null -eq $dataGridView) {
                    throw "DataGridView is null or not assigned."
                }

                $x = 50  # Initial x position
                $y = 50  # Initial y position

                # Define cell dimensions
                $cellWidth = 180
                $cellHeight = 25

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

                $x = 50  # Reset x position
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




# Add CHECK button
$CheckButton = New-Object System.Windows.Forms.Button
$CheckButton.Location = New-Object System.Drawing.Point(280, 430)
$CheckButton.Size = New-Object System.Drawing.Size(100, 30)
$CheckButton.Text = "CHECK"
$CheckButton.BackColor = [System.Drawing.SystemColors]::ActiveCaption
$CheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CheckButton.FlatAppearance.BorderSize = 3
$CheckButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$CheckButton.Anchor = "Bottom"

$CheckButton.Add_Click({
    AddServerNamesToGridViewAndCheck -serverTextbox $serverTextbox -dataGridView $dataGridView -toolStripProgressBar $toolStripProgressBar -statusStrip $statusStrip
})

$form.Controls.Add($CheckButton)

# Add CLEAR button
$ClearButton = New-Object System.Windows.Forms.Button
$ClearButton.Location = New-Object System.Drawing.Point(410, 430)
$ClearButton.Size = New-Object System.Drawing.Size(100, 30)
$ClearButton.Text = "CLEAR"
$ClearButton.BackColor = [System.Drawing.SystemColors]::ButtonFace
$ClearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ClearButton.FlatAppearance.BorderSize = 3
$ClearButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$ClearButton.Anchor = "Bottom"

$ClearButton.Add_Click({
    ClearTable -dataGridView $dataGridView
    $serverTextbox.Text = ""
    $toolStripProgressBar.Value = 0
})

$form.Controls.Add($ClearButton)

# Add EXIT button
$ExitButton = New-Object System.Windows.Forms.Button
$ExitButton.Location = New-Object System.Drawing.Point(540, 430)
$ExitButton.Size = New-Object System.Drawing.Size(100, 30)
$ExitButton.Text = "EXIT"
$ExitButton.BackColor = [System.Drawing.Color]::LightCoral
$ExitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ExitButton.FlatAppearance.BorderSize = 3
$ExitButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$ExitButton.Anchor = "Bottom"

$ExitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($ExitButton)

# Create StatusStrip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = "Bottom"
$statusStrip.BackColor = [System.Drawing.Color]::White
$statusStrip.Height = 25

# Add ToolStripProgressBar to StatusStrip
$toolStripProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$toolStripProgressBar.Size = New-Object System.Drawing.Size(100, 16)
$toolStripProgressBar.Anchor = "Right","Bottom"
$statusStrip.Items.Add($toolStripProgressBar)
$form.Controls.Add($statusStrip)

# Show the form
$form.ShowDialog() | Out-Null

