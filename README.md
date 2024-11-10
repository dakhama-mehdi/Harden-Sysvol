# HardenSysvol: Inspect AD Sysvol Vulnerabilities in Under 5 Minutes

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1.0%2B-blue.svg)
![Status: In Progress](https://img.shields.io/badge/Status-In%20Progress-orange)
![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/Hardensysvol?color=orange&label=Download%20Powershell%20Gallery)
![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/Hardensysvol)


Support this project ❤️ [PayPal](https://www.paypal.com/paypalme/mdunca13)

## 🚧 Project Status: In Progress

## Description
HardenSysvol is a free PowerShell module designed to quickly scan one of Active Directory's most critical folders, Sysvol, with just one command, no major prerequisites, and no elevated rights. It operates in read-only mode (safemode), ensuring no modifications are made to the system. 
It detects suspicious binaries from over 180 standard extensions, identifies vulnerabilities, and searches for both default and custom keywords.

This tool complements audit solutions like PurpleKnight, PingCastle, and GPOZaurr by offering deeper insights into Sysvol content, often overlooked by other tools. HardenSysvol is ideal for AD audits and pentesting.

<a href="https://dakhama-mehdi.github.io/Harden-Sysvol/Exemples_HTML/hardensysvol.html#Tab-zqtd4y6c" target="_blank">View Example HTML Page</a>

## Key Features of Hardensysvol

| **Feature**                         | **Description**                                                                                                      | **Supported File Types**                                                                |
|-------------------------------------|----------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------|
| **Binary Comparison**               | Analyzes and compares well-known binaries with the ability to extend to additional signatures to detect suspicious files. | All binary types (EXE, DLL, MSI, etc.) with customizable signature extension support 180 extensions listed below .                  |
| **Keyword Search**                  | Searches for sensitive keywords such as passwords and usernames across a wide variety of files.                      | Pdf, docx, xlsx, doc, xls, pptx, ods, odt, odp, bat, reg, ps1, vbs, py, xml, and other scripts.                                  |
| **Certificate Verification**        | Verifies certificates protected by password or containing exportable private keys.                                    | PFX, CER, DER, PEM, P7B certificates.                                                    |
| **Steganography**                   | Analyzes images to detect hidden files by searching for file signatures like EXE, ZIP, etc.                           | Images (JPEG, PNG, BMP, GIF, etc.) and hidden files (EXE, MSI, ZIP, RAR, 7z).                 |
| **File Signature Verification**     | Verifies file signatures for security compliance, including detection of password-protected ZIP files.              | MSI, EXE, DLL, JAR, MSU, CAB, and ZIP (password-protected).                                 |


## Requirements
- **PowerShell**: 5.1 or higher.
- **Permissions**: The tool can be run by any standard account on the domain.
- **Compatibility**: Works with Windows Server environments and Windows 10/11

## Installation from Powershell Gallery
Run the following command in PowerShell:
```powershell
Install-Module -Name HardenSysvol -Scope CurrentUser -Force
````
### To launch the scan
```powershell
Invoke-HardenSysvol
````
### Frequently Used Example
```powershell
Invoke-HardenSysvol -Addpattern admin -Addextension adml,admx,adm
Invoke-HardenSysvol -Allextensions
````

### Parameters

| Parameter      | Explanation                                                                                               | Example                                         |
|----------------|-----------------------------------------------------------------------------------------------------------|-------------------------------------------------|
| Addpattern     | Adds custom keywords to search for that are not present by default.                                       | `-Addpattern admins,@mydomain,hack`             |
| Removepattern  | Removes a keyword from the default search list.                                                           | `-Removepattern ipv4,sha1,password`             |
| Addextension   | Adds an additional file extension to include in the search.                                               | `-Addextension adml,admx,adm`                   |
| Ignoreextension| Excludes a default extension from the search.                                                             | `-Ignoreextension pdf,bat,ps1`                  |
| Allextensions  | Scans all file types without any exceptions.                                                              | `-Allextensions`                                |
| DnsDomain      | Targets a specific child domain or Domain Controller (DC).                                                | `-Dnsdomain dc-2` or `-Dnsdomain domain.local` |
| Custompatterns | Allows the use of a custom pattern file, as long as it follows the original .xml format.                  | `-Custompatterns C:\temp\custom.xml`            |

## How It Works
HardenSysvol first analyzes the shared folders on the Domain Controller where it is run, or on a specified target defined by parameters. For each file, it checks against a list of 180 default extensions. If a file, such as a .doc file, is renamed to .exe (or vice versa), it will trigger an error, making it difficult for suspicious files to bypass detection.

The tool also performs keyword searches within scripts, inspects certificate signatures, and identifies hidden files embedded in images. This multi-layered analysis helps uncover vulnerabilities that might otherwise be overlooked, providing administrators with comprehensive security insights.

## Default file types, extensions, and patterns

| Category                  | Details                                                                                                                           |
|---------------------------|-----------------------------------------------------------------------------------------------------------------------------------|
| **Default Extensions**    | `bat`, `bmp`, `cab`, `class`, `csproj`, `config`, `csv`, `cer`, `der`, `doc`, `docx`, `dll`, `exe`, `gif`, `gz`, `html`, `ico`, `ini`, `jar`, `jpg`, `jpeg`, `jp2`, `msi`, `msu`, `mp3`, `odp`, `ods`, `odt`, `otf`, `ott`, `p7b`, `p7c`, `pdf`, `pfx`, `png`, `pol`, `pptx`, `ppt`, `py`, `ps1`, `psm1`, `rar`, `rdp`, `reg`, `rtf`, `tar`, `tif`, `txt`, `vbs`, `xls`, `xlsx`, `xml`, `vbsx`, `webp`, `zip`, `7z`,`kdb` ,`db`  |
| **Default Pattern Check** | `accesskey`, `auth`, `credentials`, `cred`, `identifiant`, `mdp`, `mdpass`, `motdepasse`, `private-key`, `pwd`, `secret`, `ssh-key`, `token`, `login`, `apikey`, `password`, `securestring`, `md5`,`SHA-1`, `SHA-256`, `SHA-512`, `net user`,`ipv4` |
| **Magic Numbers**         | "doc", "xls", "msi", "ppt", "vsd", "db", "msg", "xla", "apr", "dot", "suo","epub", "docx", "xlsx", "pptx", "odp", "ods", "jar","odt", "zip", "ott","vsdx", "xps", "kmz", "kwd", "oxps", "sxc", "sxd", "sxi", "sxw", "xpi","msix", "exe", "bin", "dll", "IDX_DLL", "sys", "tlb", "ocx", "olb", "odf","rll", "rar", "7z", "png", "pdf", "jpg", "jpeg", "gif", "tif", "ico","class", "msu", "cab", "bmp", "p7b", "p7c", "p7s", "cer", "pfx", "der","pem", "otf", "webp", "avi", "wav", "tar", "jp2", "kdb", "kdbx", "rtf","mpg", "mpeg", "mp4", "ogg", "flac", "mkv", "webm", "vmdk", "pst", "mdb","eps", "sln", "123", "ttf", "tgz", "gz", "hqx", "mxf", "oga", "ogv", "ogx","p10", "ai", "fdf", "msf", "fm", "tpl", "wk4", "wk3", "wk1", "nsf", "ntf","org", "lwp", "sam", "mif", "asf", "wma", "wmv", "chm", "wks", "qxd", "mmf","cap", "dmp", "wpd", "xar", "spf", "dtd", "amr","au", "m4a", "koz", "mp3","fits", "tiff", "psd", "dwg", "hdr", "wmf", "eml", "vcf", "dms", "3g2",    "3gp", "m4v", "mov", "cpt", "vcd", "csh", "rpm", "swf", "sit", "xz", "mid",    "midi", "aiff", "ram", "rm", "ra", "pgm", "sqlite", "rgb" |

## Credits
This project makes use of [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice), which provides essential functionality for read/write office.

Special thanks to the contributors of [HardenAD community](https://hardenad.net/) for their work and dedication.

## Licence
This project follows the AGPLv3 license due to its use of PsWritePDF, which relies on iText 7 Community for .NET under AGPLv3. If PsWritePDF were replaced by another PDF library with a more permissive license, such as MIT, we could adopt a different licensing model. For now, as long as this project remains open-source and free, the AGPLv3 requirements should not pose any issues.

Before duplicating or using this code, please review the following resources to understand the terms of this license:

[PsWritePDF Project](https://github.com/EvotecIT/PSWritePDF)
AGPLv3 License Overview


