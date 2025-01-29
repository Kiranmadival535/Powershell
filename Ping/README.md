# Ping Testing Tool

## Author: Kiran Madival  
**Version:** 1.0  
**Date:** 24/10/2024  

## Table of Contents
1. [Introduction](#introduction)
2. [Features](#features)
3. [Installation Guide](#installation-guide)
4. [How to Use](#how-to-use)
5. [Technical Overview](#technical-overview)
6. [Export and Print Capabilities](#export-and-print-capabilities)
7. [FAQ](#faq)
8. [Troubleshooting](#troubleshooting)
9. [License and Credits](#license-and-credits)

## Introduction
The **Ping Testing Tool** is a graphical user interface (GUI) application developed using PowerShell to assist administrators or users in testing the connectivity of multiple servers. It allows you to input server details, perform ping tests, and view the results in an organized manner. The tool also provides options to save and print reports, making it a convenient solution for documentation and troubleshooting purposes.

## Features
- **Ping Testing**: Check the connectivity status of multiple servers simultaneously.
- **GUI**: A simple, intuitive graphical interface for seamless user experience.
- **Progress Bar**: Visual representation of test progress as each server is pinged.
- **Save & Print**: Export the ping results and server information in Excel, PDF, or CSV formats.
- **Error Handling**: Accurate error reporting in cases where servers are unreachable.
- **Cross-Platform**: Runs on any Windows machine where PowerShell is available.

## Installation Guide

### Prerequisites:
- **Windows OS** (compatible with PowerShell)
- **.NET Framework** (required for Windows Forms)

### Installation Steps:
1. Download the executable file: `PingTestingTool.exe`.
2. Place the executable in a folder of your choice.
3. Double-click the file to run the tool.

## How to Use

### Main Interface
The tool features a clean and organized layout with the following key components:

- **Server Name**: Displays the name of the server being tested.
- **IP Address**: Shows the corresponding IP address of the server.
- **Domain Name**: Displays the domain name of the server (if available).
- **Ping Result**: The status of the server (Green for reachable, Red for unreachable).

### Buttons and Functionalities:
- **CHECK**: Starts the ping test for the listed servers and updates the progress bar.
- **CLEAR**: Clears the input, results, and resets the progress bar to 0%.
- **EXIT**: Closes the application.

The progress bar visually shows the status of the ping tests and fills up as the servers are tested, providing a clear indication of progress.

## Technical Overview
The Ping Testing Tool is built using PowerShell scripts combined with Windows Forms for the GUI. It leverages the `Test-Connection` cmdlet in PowerShell to send ICMP echo requests to the specified servers.

### Key Components:
- **PowerShell Scripting**: Handles backend operations such as performing ping tests and gathering results.
- **Windows Forms**: The GUI is constructed using Windows Forms, making it easy for users to interact with the tool.
- **Data Handling**: Results are dynamically displayed in the DataGridView component, allowing users to view detailed ping statuses.

## Export and Print Capabilities
The tool offers robust export and print options:
- **Export**: Save the results in Excel, CSV, or PDF formats.
- **Print**: Generate hard copies of the reports.

## FAQ

- **Q1: What happens if a server is unreachable?**  
  A1: The ping result will show as *Not Pinging* in red.

- **Q2: Can I add multiple servers for testing?**  
  A2: Yes, the tool allows you to input and test multiple servers at once.

- **Q3: How do I clear the results?**  
  A3: Click on the *Clear* button to remove all data from the input fields and results table.

- **Q4: How do I save the results?**  
  A4: Use the *Save* option in the top menu to export the results in CSV, Excel, or PDF formats.

## Troubleshooting

### Common Issues:
- **Ping Fails for All Servers**: Ensure the servers are accessible and that the network connection is stable.
- **Application Not Responding**: Restart the tool and ensure your PowerShell environment is properly configured.

### Error Messages:
- **IP Address Not Found**: This indicates that the server details provided are incorrect or unreachable.
- **Domain Name Not Found**: This occurs when no domain information is available for the server.

## License and Credits
This tool was created by Kiran Madival. It is provided as-is without warranty, and users are encouraged to adapt it to their environment.

---
Feel free to adapt this README as needed and add any extra details specific to your project or user needs!
