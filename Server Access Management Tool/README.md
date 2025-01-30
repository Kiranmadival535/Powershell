# Server Access Management Tool 📑
**Author**: Kiran Madival  
**Version**: 1.0  
**Date**: 24/10/2024

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

---

## 1. Introduction
The **Server Access Management Tool** is designed to streamline user role management across multiple Windows servers. It allows administrators to check, assign, and remove user roles efficiently. The tool provides real-time access information for **Remote Desktop (RDP)** and **administrative users**, simplifying the management of user access levels across multiple servers.

---

## 2. Features
- **Bulk Role Management**: Easily add or remove users to groups across multiple servers.
- **Access Level Control**: Assign access levels such as Admin or Remote Desktop User to new and existing users.
- **Real-Time User Role Display**: View the current user roles on selected servers.
- **Multi-Server Support**: Manage user roles on multiple servers simultaneously.
- **User-Friendly Interface**: Simple GUI for easy navigation and usage.
- **Export and Print**: Export data in CSV, PDF, or Excel formats and print it for auditing purposes.

---

## 3. Installation Guide
### Prerequisites
- Windows OS (compatible with PowerShell)
- Network connectivity to target servers

### Installation Steps:
1. Download the executable file: `ServerAccessManagementTool.exe`.
2. Place the executable in a folder of your choice.
3. Right-click the file and select **'Run as Administrator'**.

---

## 4. How to Use
### Main Interface
The tool has a clean layout with these key sections:

- **Load User Roles**: Enter server names and click **CHECK** to load current user roles.
- **Assign Roles**: Select servers and users, then choose an access level (Admin, Remote Desktop User) and click **ADD**.
- **Remove Roles**: Select servers and users to remove roles by clicking **REMOVE**.
- **Save or Print Configurations**: Use **Save** or **Print** in the menu to document the current user roles and access levels.
- **Clear Entries**: Press **CLEAR** to reset all fields.

### Example:
- Buttons include **CHECK**, **CLEAR**, and **EXIT**.

The progress bar shows the status of the operations.

---

## 5. Technical Overview
The tool leverages **PowerShell scripting** and **Windows Forms** to provide a user-friendly interface for remote server role management. It uses **WMI** and **Active Directory** commands to fetch and modify user roles on remote servers.

Key Components:
- **PowerShell**: Backend logic for user role retrieval and management.
- **Windows Forms**: GUI for user interaction.
- **WMI & Active Directory**: Remote interaction with user roles and permissions.

---

## 6. Credential Management
Credentials are stored securely in `C:\temp` for the duration of your session. They are cleared once you exit or hit **CLEAR**, ensuring no lingering sensitive data.

---

## 7. Export and Print Capabilities
- **Export**: Save configurations in **CSV**, **Excel**, or **PDF** formats.
- **Print**: Print the user roles and access levels for documentation.

---

## 8. FAQ
- **Q1**: Can I manage Linux server users with this tool?  
  **A1**: No, this tool is specifically designed for Windows servers.
  
- **Q2**: Do I need admin rights on the target servers?  
  **A2**: Yes, administrator privileges are required to modify user roles on remote servers.

- **Q3**: Can I add multiple users to multiple servers at once?  
  **A3**: Yes, the tool supports bulk operations.

- **Q4**: How do I save the results?  
  **A4**: Use the **Save** option in the top menu to export the results.

---

## 9. Troubleshooting
- **Unable to load roles**: Ensure server names are correct and reachable.
- **Permission Denied**: Run the tool with administrator privileges.
- **Error with Add/Remove**: Double-check selected access level and valid user names.

---

## 10. License and Credits
- **Developed by**: Kiran Madival  
- **Acknowledgments**: Feedback and suggestions for improvements are welcome.
