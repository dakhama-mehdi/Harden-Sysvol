# Hardensysvol - Ensure the security of your AD

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1.0%2B-blue.svg)
![Status: In Progress](https://img.shields.io/badge/Status-In%20Progress-orange)
![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/Hardensysvol?color=orange&label=Download%20Powershell%20Gallery)
![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/Hardensysvol)

### Support This Project : ❤️<a href="https://www.paypal.com/paypalme/mdunca13" target="_blank">Buy me a cofee /paypalme</a>☕

## 🚧 Project Status: In Progress

⚠️ **This project is currently under construction**. Features may change, and some functionality might be incomplete. Please feel free to test it and report any issues or suggestions as we continue to improve it.


## Description
*Hardensysvol* is a PowerShell module designed to enhance Active Directory (AD) security by analyzing and detecting threats within the Sysvol folder. It scans for sensitive keywords, identifies suspicious files, and generates a detailed HTML report for easier filtering. 

Hardensysvol can be used for AD audits or pentesting, complementing existing solutions such as PingCastle, PurpleKnight, and GPOZaurr.
## Key Features of Hardensysvol

| **Feature**                         | **Description**                                                                                                      | **Supported File Types**                                                                |
|-------------------------------------|----------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------|
| **Binary Comparison**               | Analyzes and compares well-known binaries with the ability to extend to additional signatures to detect suspicious files. | All binary types (EXE, DLL, MSI, etc.) with customizable signature extension.                  |
| **Keyword Search**                  | Searches for sensitive keywords such as passwords and usernames across a wide variety of files.                      | Pdf, docx, xlsx, doc, xls, pptx, ods, odt, odp, bat, reg, ps1, vbs, py, xml, and other scripts.                                  |
| **Certificate Verification**        | Verifies certificates protected by password or containing exportable private keys.                                    | PFX, CER, DER, PEM, P7B certificates.                                                    |
| **Steganography**                   | Analyzes images to detect hidden files by searching for file signatures like EXE, ZIP, etc.                           | Images (JPEG, PNG, BMP, GIF, etc.) and hidden files (EXE, MSI, ZIP, RAR, 7z).                 |

## Requirements
- **PowerShell**: 5.1 or higher.
- **Permissions**: The tool can be run by any standard account on the domain.
- **Compatibility**: Works with Windows Server environments and Windows 10/11

## Installation

### Install via PowerShell Gallery
To install directly from PowerShell Gallery, run:

### Magic number default check : 
doc, xls, msi, ppt, vsd, docx, xlsx, pptx, odp, ods, jar, odt, zip, ott, vsdx, exe, dll, rar, zip, 7z, png, pdf, jpg, jpeg, gif, tif, ico, class, msu, cab, bmp, p7b, p7c, cer, pfx, der, pem, p7b, otf, webp, mp3, gz, tar, jp2, rtf
### Default extensions support  : 
bat, bmp, cab, class, csproj, config, csv, cer, der, doc, docx, dll, exe, gif, gz, html, ico, ini, jar, jpg, jpeg, jp2, msi, msu, mp3, odp, ods, odt, otf, ott, p7b, p7c, pdf, pfx, png, pol, pptx, ppt, py, ps1, psm1, rar, rdp, reg, rtf, tar, tif, txt, vbs, xls, xlsx, xml, vbsx, webp, zip, 7z
### Default pattern check : 
accesskey, auth, credentials, cred, identifiant, mdp, mdpass, motdepasse, private-key, pwd, secret, ssh-key, token, login, apikey, password, securestring, SHA-1, SHA-256, SHA-512, net user

