# Hardensysvol - Ensure the security of your AD

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1.0%2B-blue.svg)
![Status: In Progress](https://img.shields.io/badge/Status-In%20Progress-orange)

## 🚧 Project Status: In Progress

⚠️ **This project is currently under construction**. Features may change, and some functionality might be incomplete. Please feel free to test it and report any issues or suggestions as we continue to improve it.


## Description
*Hardensysvol* is a PowerShell module designed to enhance Active Directory (AD) security by analyzing and detecting threats within the Sysvol folder. It scans for sensitive keywords, identifies suspicious files, and generates a detailed HTML report for easier filtering. 

Hardensysvol can be used for AD audits or pentesting, complementing existing solutions such as PingCastle, PurpleKnight, and GPOZaurr.
## Key Features of Hardensysvol

| **Feature**                         | **Description**                                                                                                      | **Supported File Types**                                                                |
|-------------------------------------|----------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------|
| **Binary Comparison**               | Analyzes and compares well-known binaries with the ability to extend to additional signatures to detect suspicious files. | All binary types (EXE, DLL, etc.) with customizable signature extension.                  |
| **Keyword Search**                  | Searches for sensitive keywords such as passwords and usernames across a wide variety of files.                      | Excel, docx, doc, ppt, bat, reg, xml, and other scripts.                                  |
| **Certificate Verification**        | Verifies certificates protected by password or containing exportable private keys.                                    | PFX, CER, DER, PEM, P7B certificates.                                                    |
| **Steganography**                   | Analyzes images to detect hidden files by searching for file signatures like EXE, ZIP, etc.                           | Images (JPEG, PNG, BMP, GIF, etc.) and hidden files (EXE, ZIP, RAR, 7z).                 |


## Requirements
- **PowerShell**: 5.1 or higher.
- **Permissions**: The tool can be run by any standard account on the domain.
- **Compatibility**: Works with Windows Server environments and Windows 10/11

## Installation

### Install via PowerShell Gallery
To install directly from PowerShell Gallery, run:

```powershell

