# Application Validator Tool

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
The **Application Validator Tool** is designed to streamline the process of verifying application installation status across multiple servers. By entering server names and application names, users can instantly check which applications are installed on each server, saving valuable time and effort in application validation tasks. This tool provides a visual interface to display the validation results, making it easy to monitor the application status in a centralized manner.

## Features
- **Quick Verification**: Check the installation status of applications on multiple servers simultaneously.
- **Simple Interface**: User-friendly GUI for adding server and application names and viewing results.
- **Real-Time Results**: Instantly display which applications are installed and which are missing.
- **Export and Print**: Save or print reports for documentation and audit purposes.
- **Customizable Options**: Easily add or remove servers and applications to validate based on specific needs.

## Installation Guide

### Prerequisites:
- Administrator privileges on the machine running the tool.
- Network connectivity to target servers for retrieving information.

### Installation Steps:
1. Download the tool’s executable file.
2. Right-click the file and select 'Run as Administrator.'

## How to Use

### Main Interface
The tool features a clean and organized layout with the following key components:

- **Enter Server Names**: In the 'Enter Servers Name' field, add the names of the servers you wish to validate.
- **Enter Application Names**: In the 'Enter Applications Name' field, list the applications you want to check.
- **Check Application Status**: Click on the *CHECK* button to display the installation status of each application on the specified servers.
- **View Results**: The main display area will show each server along with the corresponding application status (e.g., *Installed* or *Not Installed*).
- **Clear Data**: Use the *CLEAR* button to reset the fields and enter new server or application details if needed.
- **Exit**: Click on *EXIT* to close the tool when finished.

Example:

![Main Interface](example-image.png)  *(This is just a placeholder for a screenshot of the main interface.)*

## Technical Overview
The **Application Validator Tool** is developed using PowerShell and Windows Forms for the graphical user interface (GUI). The tool takes input for server names and application names and validates the installation status remotely.

### Key Components:
- **PowerShell Scripting**: Handles backend operations such as connecting to servers and checking installed applications.
- **Windows Forms**: Constructs the graphical interface that users interact with.
- **Application Validation**: Uses PowerShell commands to verify if the specified applications are installed on the target servers.

## Credential Management
The tool allows users to enter credentials securely for server access. These credentials are temporarily stored in `C:\temp` during the session and cleared immediately after clicking *Clear* or closing the tool. This ensures secure handling of user credentials.

## Export and Print Capabilities
The tool includes options for exporting and printing validation reports:
- **Export to Excel/PDF/CSV**: Save the validation results in various formats for record-keeping and audits.
- **Print**: Provides an option to print the results for physical documentation.

## FAQ

- **Q1: Can I add multiple applications at once?**  
  **A1**: Yes, you can list multiple applications in the 'Enter Applications Name' field, separated by commas.

- **Q2: What types of applications can this tool validate?**  
  **A2**: The tool can check for any installed applications, as long as it has the necessary permissions on the target servers.

- **Q3: Is this tool compatible with non-Windows servers?**  
  **A3**: Currently, the tool is designed to work only in Windows-based environments.

## Troubleshooting

### Common Issues:
- **Issue: Results are not displaying after pressing CHECK.**  
  **Solution**: Ensure the server names and application names are entered correctly and that you have network access to the servers.

- **Issue: Unable to connect to a server.**  
  **Solution**: Verify that the server is online, and check that you have the necessary permissions and credentials to access it.

## License and Credits
Developed by: Kiran Madival  
Acknowledgments: Feedback and suggestions for improvements are always welcome.

---
Feel free to modify or extend the sections above as needed. This template provides a solid foundation for your project's documentation.
