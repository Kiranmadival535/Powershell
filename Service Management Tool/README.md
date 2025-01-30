# Server Service Management Tool

## Author: Kiran Madival  
**Version:** 1.0  
**Date:** 24/10/2024  

## Table of Contents
1. [Introduction](#introduction)
2. [Features](#features)
3. [Installation Guide](#installation-guide)
4. [How to Use](#how-to-use)
5. [GUI Overview](#gui-overview)
6. [Technical Overview](#technical-overview)
7. [Credential Management](#credential-management)
8. [Export and Print Capabilities](#export-and-print-capabilities)
9. [FAQ](#faq)
10. [Troubleshooting](#troubleshooting)
11. [License and Credits](#license-and-credits)

## Introduction
The **Server Service Management Tool** is a powerful solution for system administrators to manage and monitor the status of services across multiple servers. With a user-friendly graphical interface, the tool simplifies tasks like starting, stopping, and checking the status of services on remote servers. It also provides export and print capabilities for documentation and auditing purposes.

## Features
- **Service Management**: Monitor, start, and stop services on multiple servers.
- **Multiple Servers Input**: Add several server names at once for bulk management.
- **Start/Stop Services**: Easily start or stop selected services on remote servers.
- **Real-time Status Updates**: Displays the current status of services, showing whether they are running or stopped.
- **Save & Print Options**: Export results in Excel, PDF, or CSV format.
- **User-friendly GUI**: Clean and intuitive design for efficient server service management.
- **Error Handling**: Informative error messages in case of server connection failures or invalid service names.

## Installation Guide

### Prerequisites:
- Windows PowerShell 5.0 or later.
- Administrator privileges to run PowerShell scripts.
- Necessary network access to the servers you wish to manage.

### Installation Steps:
1. Download the executable file for the Server Service Management Tool.
2. Right-click the file and choose *Run as Administrator*.
3. Follow the on-screen prompts to complete the installation process.
4. Once installed, launch the application from your desktop or start menu.

## How to Use

### Main Interface
The tool features a clean and organized layout with the following key components:

1. **Enter Server Names**: Add server names in the 'Enter Servers Name' section on the left panel. The servers should be listed one by one.
2. **Check Service Status**: Press the *CHECK* button after entering the server names. The tool will query the status of the selected services across all listed servers.
3. **Start/Stop Services**: To start or stop a service, select the desired service from the list and press the *START* or *STOP* buttons, depending on the required action.
4. **Clear Fields**: The *CLEAR* button resets the list and removes the server names and service statuses.
5. **Exit**: Use the *EXIT* button to close the application.

## GUI Overview
The tool’s GUI is divided into the following sections:
- **Servers Input Panel**: On the left side, users can input multiple server names.
- **Service Information Table**: In the central part, the tool displays the server name, the service selected, and its current status (Running/Stopped).
- **Control Buttons**:
  - *ENTER*: Loads servers into the data grid table and fetches the services list for each server.
  - *CHECK*: Queries the status of services for the entered servers.
  - *CLEAR*: Resets the tool, clearing all inputs and results.
  - *START* and *STOP*: Start or stop the selected service on the listed server.
  - *EXIT*: Closes the tool.
- **Menu Options**:
  - *File Menu*: Includes options to save the results or print the data.
  - *Help Menu*: Provides user documentation and troubleshooting help.

Example:

![Main Interface](example-image.png) *(This is just a placeholder for a screenshot of the main interface.)*

## Technical Overview
The **Server Service Management Tool** is built using PowerShell with Windows Forms for the graphical interface. It connects to remote servers using WMI (Windows Management Instrumentation) to query the status of services and execute start/stop commands. The tool is wrapped as an executable for ease of distribution and does not require users to manually execute scripts.

### Core Components:
- **PowerShell Scripting**: Handles all back-end logic for managing services.
- **Windows Forms GUI**: Provides an intuitive interface for user interaction.
- **WMI Integration**: Allows the tool to query and control services remotely.
- **Export Functionality**: Allows users to export the service information to Excel, PDF, and CSV formats.

## Credential Management
To keep things secure and hassle-free, this tool asks for your credentials just once per session. These credentials are securely stored in the `C:\temp` folder for temporary use, letting you work without interruptions. When you’re done, simply hit *Clear* or close the tool to safely lock everything up—no leftover credentials, no extra worries. It’s as easy as enter, use, and erase!

## Export and Print Capabilities
The tool offers flexible export and print options:
- **Export**: The results of service queries can be saved as Excel, PDF, or CSV files. This is ideal for archiving, auditing, or sharing with team members.
- **Print**: The service status report can also be printed directly from the tool for immediate documentation or reference.

## FAQ

- **Q1: What happens if a server is unreachable?**  
  **A1**: The tool will display an error message and move on to the next server without crashing.

- **Q2: Can I manage services on both Windows and Linux servers?**  
  **A2**: The tool is currently designed to manage Windows services only.

- **Q3: Do I need administrator privileges to run this tool?**  
  **A3**: Yes, administrator privileges are required for starting and stopping services.

- **Q4: How many servers can I manage at once?**  
  **A4**: The tool can handle multiple servers, depending on your network configuration and system performance.

## Troubleshooting

### Common Issues:
- **Issue: Unable to connect to server.**  
  **Solution**: Ensure that the target server is online and reachable via the network.

- **Issue: Services not found.**  
  **Solution**: Verify that the service names are spelled correctly and exist on the target server.

- **Issue: Permissions error.**  
  **Solution**: Make sure you are running the tool with administrator privileges.

- **Issue: Export errors.**  
  **Solution**: Ensure the tool has permission to write to the destination folder when exporting data.

## License and Credits
- **Developed by**: Kiran Madival  
- Contributions and feedback are welcome to improve future versions of the tool.

---
Feel free to modify or extend the sections above as needed. This template provides a solid foundation for your project's documentation.
