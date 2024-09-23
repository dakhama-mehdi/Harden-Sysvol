# Hardensysvol - Ensure the security of your AD

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1.0%2B-blue.svg)
![Status: In Progress](https://img.shields.io/badge/Status-In%20Progress-orange)

## 🚧 Project Status: In Progress

⚠️ **This project is currently under construction**. Features may change, and some functionality might be incomplete. Please feel free to test it and report any issues or suggestions as we continue to improve it.


## Description
*Hardensysvol* is a PowerShell module designed to scan the Sysvol folder for files containing sensitive information, such as passwords, usernames, certificates, and configuration data. 
It helps in identifying potential security risks by detecting files that may expose sensitive content, such as documents, scripts, and configuration files.

The tool analyzes file integrity, flags files that require further scrutiny, and helps administrators improve the overall security of their **Active Directory** environment by ensuring that the Sysvol folder does not inadvertently expose sensitive information.

## Features
- Scans the Sysvol folder for files containing sensitive data.
- Detects potential security risks, including passwords, usernames, and configuration details.
- Analyzes file integrity for discrepancies and unusual data patterns.
- Supports detection of document types like `docx`, `xlsx`, `pdf`, and more.
- Flags files for further inspection if integrity checks fail or content appears suspicious.
- Generates detailed reports of findings, helping with the security hardening of your **Active Directory**.

## Requirements
- **PowerShell**: 5.1 or higher.
- **Permissions**: The tool can be run by any standard account on the domain.
- **Compatibility**: Works with Windows Server environments and Windows 10/11

## Installation

### Install via PowerShell Gallery
To install directly from PowerShell Gallery, run:

```powershell

