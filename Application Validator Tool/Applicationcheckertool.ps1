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



# Function to check if specified applications are installed on servers
function Check-Applications {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [string[]]$Applications,
        [System.Management.Automation.PSCredential]$Credential
    )

    foreach ($row in $DataGridView.Rows) {
        $server = $row.Cells["ServerName"].Value

        foreach ($app in $Applications) {
            $installed = $null
            try {
                # Use Invoke-Command to run the command on the remote server with the provided credential
                $installed = Invoke-Command -ComputerName $server -Credential $Credential -ScriptBlock {
                    param ($appName)
                    Get-WmiObject -Class Win32_Product -Filter "Name LIKE '%$appName%'"
                } -ArgumentList $app

            } catch {
                Write-Host "Error occurred while checking $app on $server : $_"
            }

            if ($installed) {
                $row.Cells[$app].Value = "Installed"
            } else {
                $row.Cells[$app].Value = "Not Installed"
            }
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


# Function to clear all columns except "ServerName" and clear rows
function ClearTable {
    param (
        [System.Windows.Forms.DataGridView] $dataGridView
    )
    
    # Clear existing columns except "ServerName"
    $columnsToRemove = @()
    foreach ($column in $dataGridView.Columns) {
        if ($column.Name -ne "ServerName") {
            $columnsToRemove += $column
        }
    }
    
    foreach ($column in $columnsToRemove) {
        $dataGridView.Columns.Remove($column)
    }
    
    # Clear existing rows
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


# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Application validator Tool"
$form.Size = New-Object System.Drawing.Size(1000, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::Teal
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

# Create DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 350)
$dataGridView.Location = New-Object System.Drawing.Point(200, 60)
$dataGridView.Anchor = "Top","Bottom", "Left", "Right"

$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.RowHeadersVisible = $false
$dataGridView.AllowUserToAddRows = $False
# Configure DataGridView column headers
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::NavajoWhite
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$dataGridView.ColumnHeadersHeight = 32;

# Create a Font object with size 10 and bold style
$font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)

# Set the font for column headers
$dataGridView.ColumnHeadersDefaultCellStyle.Font = $font

# Add columns to the DataGridView
$serverNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$serverNameColumn.HeaderText = "Server Name"
$serverNameColumn.Name = "ServerName"
$serverNameColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGridView.Columns.Add($serverNameColumn)



# Center align all text in cells
$dataGridView.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

# Add DataGridView to the form
$form.Controls.Add($dataGridView)

# Add server label and textbox
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(20, 60)
$serverLabel.Size = New-Object System.Drawing.Size(180,32)
$serverLabel.Text = "Enter Servers Name"
$serverLabel.Anchor = "TOP","LEFT"
$serverLabel.AutoSize = $false
$serverLabel.BackColor = [System.Drawing.Color]::NavajoWhite
$serverLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$serverLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$serverLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($serverLabel)

$serverTextbox = New-Object System.Windows.Forms.TextBox
$serverTextbox.Location = New-Object System.Drawing.Point(20,92)
$serverTextbox.Size = New-Object System.Drawing.Size(180, 160)
$serverTextbox.Anchor = "Left","Top","Bottom"
$serverTextbox.Multiline = $true
$serverTextbox.ScrollBars = "Vertical"
$serverTextbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($serverTextbox)


$appLabel = New-Object System.Windows.Forms.Label
$appLabel.Location = New-Object System.Drawing.Point(20, 260)
$appLabel.Size = New-Object System.Drawing.Size(180,30)
$appLabel.Text = "Enter Applications Name"
$appLabel.Anchor = "Bottom","LEFT"
$appLabel.AutoSize = $false
$appLabel.BackColor = [System.Drawing.Color]::NavajoWhite
$appLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$appLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$appLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($appLabel)

$appTextbox = New-Object System.Windows.Forms.TextBox
$appTextbox.Location = New-Object System.Drawing.Point(20,290)
$appTextbox.Size = New-Object System.Drawing.Size(180, 120)
$appTextbox.Anchor = "Left","Bottom"
$appTextbox.Multiline = $true
$appTextbox.ScrollBars = "Vertical"
$appTextbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($appTextbox)
 
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

    $applications = $dataGridView.Columns | Where-Object { $_.Name -ne "ServerName" } | ForEach-Object { $_.HeaderText }
    Check-Applications -DataGridView $dataGridView -Applications $applications -Credential $credential
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
    $serverTextbox.Text = ""
    $appTextbox.Text = ""
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

# Add Enter button
$button4 = New-Object System.Windows.Forms.Button
$button4.Location = New-Object System.Drawing.Point(40, 435)
$button4.Size = New-Object System.Drawing.Size(100, 30)
$button4.Text = "Enter"
$button4.BackColor = [System.Drawing.Color]::RosyBrown
$button4.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button4.FlatAppearance.BorderSize = 3
$button4.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$button4.Anchor = "Bottom","Left"


# Assuming $dataGridView already exists and is configured elsewhere in your script

# Assuming $dataGridView already exists and is configured elsewhere in your script

# Function to handle Enter button click
$button4.Add_Click({
    # Clear existing columns except "ServerName"
    $columnsToKeep = @()
    foreach ($column in $dataGridView.Columns) {
        if ($column.Name -eq "ServerName") {
            $columnsToKeep += $column
        } else {
            $dataGridView.Columns.Remove($column)
        }
    }

    # Ensure the "Server Name" column exists or add it if not
    $serverColumn = $dataGridView.Columns["ServerName"]
    if (-not $serverColumn) {
        $serverColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $serverColumn.HeaderText = "Server Name"
        $serverColumn.Name = "ServerName"
        $serverColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
        $dataGridView.Columns.Add($serverColumn)
    }

    # Get server names from the server textbox
    $servers = $serverTextbox.Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # Clear existing rows and add new rows based on server names
    $dataGridView.Rows.Clear()
    foreach ($server in $servers) {
        $row = $dataGridView.Rows.Add()
        $dataGridView.Rows[$row].Cells["ServerName"].Value = $server.Trim()
    }

    # Get application names from the textbox
    $applications = $appTextbox.Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # Add columns for each application
    foreach ($app in $applications) {
        if (-not $dataGridView.Columns[$app]) {
            $newColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $newColumn.HeaderText = $app
            $newColumn.Name = $app
            $newColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
            $dataGridView.Columns.Add($newColumn)
        }
    }
})


$form.Controls.Add($button4)

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
    [System.Windows.Forms.MessageBox]::Show("This tool was created by Kiran to verify the installed applications on the servers.", "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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
                $watermarkFont = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
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
