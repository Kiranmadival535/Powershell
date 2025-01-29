# Drive Space Monitoring Tool

## Author: Kiran Madival  
**Version:** 1.0  
**Date:** 24/10/2024  

## Table of Contents
1. [Introduction](#introduction)
2. [Features](#features)
3. [Installation Guide](#installation-guide)
4. [How to Use](#how-to-use)
5. [Technical Overview](#technical-overview)
6. [Credential Management](#credential-management)
7. [Export and Print Capabilities](#export-and-print-capabilities)
8. [FAQ](#faq)
9. [Troubleshooting](#troubleshooting)
10. [License and Credits](#license-and-credits)

## Introduction
The **Drive Space Monitoring Tool** is designed to assist administrators in monitoring disk usage across multiple servers. This tool allows users to quickly check the available and used disk space on specified drives, providing a clear overview of each server's storage capacity. It also offers features to export and print disk usage reports, making it easier to track storage trends over time.

## Features
- **Multi-Server Disk Check**: Check disk space on multiple servers simultaneously by entering server names.
- **Detailed Drive Information**: Displays total space, free space (in GB), and free space percentage for each drive on each server.
- **User-Friendly Interface**: Simplifies navigation and operation for quick and efficient drive space monitoring.
- **Export and Print**: Ability to save or print drive space information for documentation and audit purposes.
- **Clear and Exit Options**: Easily reset the fields for a new operation or exit the tool.

## Installation Guide

### Prerequisites:
- Administrator privileges on the machine running the tool.
- Network connectivity to target servers for retrieving drive information.

### Installation Steps:
1. Download the tool’s executable file.
2. Right-click the file and select 'Run as Administrator.'

## How to Use

### Main Interface
The tool features a clean and organized layout with the following key components:

- **Enter Server Names**: Input server names in the 'Enter Servers Name' section on the left.
- **Check Disk Space**: Click the *CHECK* button to load disk information for each drive on the specified servers.
- **Save or Print**: Use the *Save* or *Print* options in the menu to document the current disk space details.
- **Clear Fields**: Press *CLEAR* to reset the tool and prepare it for a new operation.
- **Exit Tool**: Press *EXIT* to close the application.

Example:

![Main Interface](example-image.png) *(This is just a placeholder for a screenshot of the main interface.)*

## Technical Overview
The **Drive Space Monitoring Tool** leverages PowerShell scripting combined with Windows Forms to create an intuitive interface for monitoring disk usage across remote servers. It uses Windows Management Instrumentation (WMI) to retrieve disk space information and display it in a structured format.

### Core Components:
- **PowerShell Scripting**: Provides the backend logic to retrieve disk space information.
- **Windows Forms GUI**: Ensures an intuitive user experience.
- **WMI Integration**: Enables remote disk space data collection from servers.
- **Export Functionality**: Allows exporting drive space data to Excel, PDF, or CSV.

## Credential Management
The tool allows users to enter credentials securely for server access. These credentials are temporarily stored in `C:\temp` during the session and are automatically erased upon exit or when the *Clear* button is pressed. This ensures safe handling of user credentials.

## Export and Print Capabilities
The tool includes export and print options to facilitate auditing and documentation:
- **Export**: Save the drive space information in Excel, PDF, or CSV formats.
- **Print**: Generate a hard copy of the current disk usage details directly from the tool.

## FAQ

- **Q1: Can this tool check disk space on Linux servers?**  
  **A1**: No, this tool is specifically designed for checking disk space on Windows servers.

- **Q2: What if a server is unreachable?**  
  **A2**: The tool will display an error message and allow you to proceed with other available servers.

- **Q3: Do I need admin rights on the target servers?**  
  **A3**: Yes, administrator privileges are required to retrieve drive space information remotely.

## Troubleshooting

### Common Issues:
- **Issue: Unable to load disk space data.**  
  **Solution**: Ensure the server names are correct and the target servers are reachable.

- **Issue: Permission Denied.**  
  **Solution**: Run the tool with administrator privileges to access disk space information.

- **Issue: Error in data retrieval.**  
  **Solution**: Check network connectivity and ensure the WMI service is enabled on the target servers.

## License and Credits
Developed by: Kiran Madival  
Acknowledgments: Suggestions and feedback for improvement are welcome.

---
Feel free to modify or extend the sections above as needed. This template provides a solid foundation for your project's documentation.
