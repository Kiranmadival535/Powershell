Server Access Management Tool – Documentation
Author: Kiran Madival
Version: 1.0
Date: 24/10/2024

Table of Contents
1. Introduction
2. Features
3. Installation Guide
4. How to Use
5. Technical Overview
6. Credential Management
7. Export and Print Capabilities
8. FAQ
9. Troubleshooting
10. License and Credits

1. Introduction
The Server Access Management Tool is designed to streamline user role management across multiple servers. This tool allows administrators to check, assign, and remove user roles efficiently. It provides real-time access information for remote desktop (RDP) and administrative users and simplifies the process of managing user access levels across multiple servers simultaneously.

2. Features
	Bulk Role Management: Easily add or remove users to groups across multiple servers.
	Access Level Control: Select different access levels, such as Admin and Remote Desktop User, for new and existing users.
	Real-Time User Role Display: Load and display current user roles on selected servers.
	Multi-Server Support: Allows the entry of multiple servers for simultaneous role management.
	User-Friendly Interface: Simple GUI for quick and effective server access management.
	Export and Print: Ability to save or print the access configuration for documentation and audit purposes.

3. Installation Guide
1. Prerequisites:
   -  Windows OS (compatible with PowerShell)
   -  Network connectivity to target servers for access management.
2. Installation Steps:
   -  Download the executable file: `PingTestingTool.exe`.
   -  Place the executable in a folder of your choice.
   -  Right-click the file and select 'Run as Administrator.'

4. How to Use
Main Interface
The tool features a clean and organized layout (as shown below):

 
- Load User Roles: Enter the server names in the 'Enter Servers Name' section. Click CHECK to load current user roles for each server. The roles (Admin or Remote Desktop User) and users will be displayed.
- Assign Roles: In the lower section, enter the server names and user names you want to modify. Choose the domain and select the desired access level (Admin, Remote Desktop User) from the dropdown menus. Click ADD to assign the specified role to the user(s) on the selected server(s).
- Remove Roles: Select the servers and users you want to remove. Choose the access level and click REMOVE to revoke the selected user's access.
- Save or Print Configurations: Use the 'Save' or 'Print' options in the menu to document the current user roles and access levels.
- Clear Entries: Press CLEAR to reset the fields for a new operation.

Example:

 
Buttons and Functionalities
	CHECK: Starts the ping test for the listed servers and updates the progress bar as each server is tested.
	CLEAR: Clears the input, results, and resets the progress bar to 0%.
	EXIT: Closes the application.

The progress bar visually shows the status of the ping tests and fills up as the servers are tested one by one, providing a clear indication of progress.
Additionally, the top menu includes options to save and print the results.
5. Technical Overview
The Server Access Management Tool leverages PowerShell scripting combined with Windows Forms to create a user-friendly interface for remote server role management. It uses WMI and Active Directory commands to fetch and modify user roles on remote servers.

Key Components:
•	PowerShell Scripting: Provides the backend logic for user role retrieval and management.
•	Windows Forms: The GUI is constructed using Windows Forms, making it easy for users to interact with    the tool.
•	WMI & Active Directory Integration: Allows interaction with user roles and permissions remotely.

6. Credential Management
For secure yet streamlined access, this tool only asks for your credentials once per session. They’re stored safely in C:\temp while you work, and vanish as soon as you hit “Clear” or exit. No lingering data, no extra steps—just smooth, secure validation every time.
7. Export and Print Capabilities
The tool offers robust export and print options to facilitate auditing and documentation:
- Export: Save the user roles and access configurations in Excel, PDF, or CSV formats.
- Print: Generate hard copies of the current server access details directly from the tool.
8. FAQ
Q1: Can I manage Linux server users with this tool?
A1: No, this tool is specifically designed for managing user roles on Windows servers.
Q2: Do I need admin rights on the target servers?
A2: Yes, administrator privileges are required to modify user roles on remote servers.
Q3: Can I add multiple users to multiple servers at once?
A3: Yes, the tool supports bulk operations for adding/removing users across multiple servers.
Q4: How do I save the results?
A4: Use the Save option in the top menu to export the results in CSV, Excel, or PDF formats.

9. Troubleshooting
- Unable to load roles: Ensure the server names are correct and the target servers are reachable.
- Permission Denied: Run the tool with administrator privileges to manage user roles.
- Error with Add/Remove: Double-check the selected access level and ensure the user names are valid in the selected domain.
10. License and Credits
Developed by: Kiran Madival
Acknowledgments: Feedback and suggestions for improvements are always welcome.

