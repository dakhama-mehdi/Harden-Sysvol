<#
.SYNOPSIS
    Sysvol Security Audit Module

.DESCRIPTION
  This module is designed to scan the Sysvol folder for files containing sensitive information, such as passwords, usernames, certificates, and configuration data. 
  It identifies potential security risks by detecting files that may expose sensitive content, such as documents, scripts, and configuration files. 
  The tool also analyzes file integrity and flags files that require additional scrutiny, helping administrators to harden their Sysvol directory and 
  Ensure a secure Active Directory environment.
  
.VERSION
    1.7

.Contribution
    Credit : HardenAD Community HardenAD
    Credit : It-connect Community It-Connect

.AUTHOR
    DAKHAMA Mehdi

.PARAMETER dnsDomain
    Specifies the DNS domain to be scanned. Defaults to the current user's DNS domain if not provided.

.PARAMETER ignoreExtensions
    Specifies file extensions to ignore during the scan.

.EXAMPLE
    # Scan the Sysvol folder of the current domain
    Invoke-HardenSysvol

    # Scan the Sysvol folder of a specific domain, ignoring .txt and .log files
    Invoke-HardenSysvol -dnsDomain "example.com" -ignoreExtensions "txt", "log" -Addpattern admin -AddExtensions adml,admx,adm

.NOTES
    This script not requires administrative privileges to access and scan the Sysvol directory.

.LINK
    https://github.com/dakhama-mehdi/Harden-Sysvol
#>

function Invoke-HardenSysvol {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$dnsDomain = $env:USERDNSDOMAIN,
        
        [Parameter(Mandatory = $false)]
        [String[]]$ignoreExtensions,

        [Parameter(Mandatory = $false)]
        [String[]]$AddExtensions,

        [Parameter(Mandatory = $false)]
        [String[]]$Addpattern,
        
        [Parameter(Mandatory = $false)]
        [String[]]$removepattern,        

        #Scann all extensions
        [Parameter(ValueFromPipeline = $true, HelpMessage = "Scann all extension")]
        [switch]$Allextensions,

        #Location the report will be saved
        [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: %temp%")]
        [String]$SavePath = $env:TEMP,

        #Location the report will be saved
        [Parameter(ValueFromPipeline = $true, HelpMessage = "maximum size of the file to be considered large in MB")]
        [int]$Maxfilesize = '10',

        #Location the report will be saved
        [Parameter(ValueFromPipeline = $true, HelpMessage = "maximum size of the binary to be considered large in MB")]
        [String]$MaxBinarysize = '50',

        #Location Custom pattern file
        [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter Custome XML file path; Default: %module%\patterns.xml")]
        [String]
        [ValidateScript({ 
        if ($_ -match '\.xml$') { 
        return $true 
        } else { 
        throw "The file must have a .xml extension."
         }
        })]$Custompatterns
    )

#region script
#region code

#region load prerequist

# Test access to the share
$testpath = Test-Path "\\$dnsDomain\sysvol\"
if ($testpath -eq $false) {
throw "Cannot access domain or share, pls check with GCI $dnsDomain"
}

# Test Modules
$modulesToCheck = @("PSWriteOffice", "PSWritePDF", "PSWriteHTML")

foreach ($module in $modulesToCheck) {
    try {
        # Check if module installed
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Output "Installation du module :  $module."
            install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Ignore
            Write-Output "The module $module has been successfully installed"
        } else {
            Write-Output "Module $module is installed"
        }
    } catch {
        Write-Error "Erreur lors d'installation du module $module : $_"
        throw "Script stopped due to an error during module installation $module."
    }
}

# Check about Word office if installed or not to read old version doc,xls
function Is-WordInstalled {
    try {
        $wordApp = New-Object -ComObject Word.Application
        # If the COM object is created successfully, Word is installed
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Checkif 7zip is installed to read rar,7z protected file by password or encrypted
function Is-7zip {
$sevenZipRegPath = "HKLM:\SOFTWARE\7-Zip"
$sevenZipPath = Get-ItemProperty -Path $sevenZipRegPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path

if (-not $sevenZipPath) {
    $sevenZipRegPath = "HKCU:\Software\7-Zip"  #Check in HKCU
    $sevenZipPath = Get-ItemProperty -Path $sevenZipRegPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path
}

if ($sevenZipPath) {
    return $sevenZipPath
} else {
    return $false
}
}

    $wordinstalled = Is-WordInstalled
    $zipinstalled = Is-7zip

    # Removing ghost runspaces
    $PreviousRS = get-runspace | where-object {($_.id -ne 1)} 
    if ($PreviousRS) { $PreviousRS.dispose() }

    $Results =  $null
    # Script start: obtain the current date and time
    $startDate = Get-Date
    
    #region getxmlcontenant
    # Retrieve file extensions from the XML file
    $xmlFileExtensions =  Join-Path -Path $PSScriptRoot -ChildPath "file_extensions.xml"
        
    $extensionsXML = [xml](Get-Content $xmlFileExtensions -Encoding UTF8)

    if (!$Allextensions) {
    $fileExtensions = $extensionsXML.root.FileExtensions.Extension
    } else {
    $fileExtensions = "*.*"
    }

    if ($AddExtensions) {

    $fileExtensions = [System.Collections.ArrayList]@($fileExtensions)

    $AddExtensions | ForEach-Object {
    $fileExtensions.Add("*." + $_)
    }

    }

    # Retrieve password patterns from the XML file
    if ($Custompatterns) {
    try {
    $CustomextensionsXML = [xml](Get-Content $Custompatterns -Encoding UTF8)
    $passwordPatterns = $CustomextensionsXML.root.PasswordPatterns.Pattern
    } catch {
    Write-Error "Error while reading the custom patterns file: $_"
    throw "The custom patterns file could not be found or is invalid."
    }
    } 
    else {
    $passwordPatterns = $extensionsXML.root.PasswordPatterns.Pattern
    }

    if ($passwordPatterns.Count -eq 0) {
    throw "The custom patterns file could not be found or is invalid."
    }
    if ($Addpattern) {
    $passwordPatterns = [System.Collections.ArrayList]@($passwordPatterns)

    $Addpattern | ForEach-Object {
    $passwordPatterns.Add($_)
    }
    }
    if ($removepattern) {
    $passwordPatterns = [System.Collections.ArrayList]@($passwordPatterns)
    $removepattern | ForEach-Object {
        $patternToRemove = $_
        $patternsToRemove = $passwordPatterns | Where-Object { $_ -like "*$patternToRemove*" }
        foreach ($pattern in $patternsToRemove) {
            $passwordPatterns.Remove($pattern)
        }
    }
}
    if ($ignoreExtensions) {
        $ignoreExtensions = $ignoreExtensions | ForEach-Object {
        "*." + $_
    }
    }

    #get binary sign from json and load module path
    $module = Join-Path -Path $PSScriptRoot -ChildPath "FileHandlers.psm1"
    $jsonfile = Join-Path -Path $PSScriptRoot -ChildPath "extensions.json"

    try {
    $jsonContent = Get-Content -Path $jsonfile -Raw | ConvertFrom-Json
    } catch {
    throw "Script stopped due to an error during json file $_ "
    }

    # Initialize a list to store the results
    $Results = @()

    #endregion getxmlcontenant

#Initialize the variables
$notAccessibleFiles = $fichiertraite = $Results = $null
$pool = $runspaces = $null

# Pool definition (creation of slots)
$pool = [RunspaceFactory]::CreateRunspacePool(1,10)
$pool.ApartmentState = "MTA"
$pool.Open()
$runspaces = @()

#endregion load prerequist

# Region Scriptfunction
$scriptblock = {
    Param (
        [string]$sysfiles,
        [string[]]$passwordPatterns,
        [string]$wordinstalled,
        [string]$zipinstalled,
        [string]$module,
        [int]$Maxfilesize,
        [int]$MaxBinarysize,
        [object]$jsonContent
    )

    # Import modul FileHandlers.psm1
    Import-Module $module -Verbose    

    [String]$detectedType =  Get-FileType -filePath $sysfiles -jsonContent $jsonContent -maxfilesize $Maxfilesize -maxbinarysize $MaxBinarysize    
       
    # Function to search pattern by extensions
    switch ($detectedType)  {
    'docx' {
        $results = Get-DocxContent -filePath $sysfiles -patterns $passwordPatterns
    }
    'xlsx' {
        $results = Get-XlsxContent -filepath $sysfiles -patterns $passwordPatterns
    }
    {$_ -in "xlsm","xlam"} {
        $results = Get-Xlsmcontent -filepath $sysfiles
    }
    'pptx' {
        $results = Get-PPTContent -filepath $sysfiles -patterns $passwordPatterns -wordinstalled $wordinstalled
    }
    'doc' {
        $results = Get-DocContent -filepath $sysfiles -patterns $passwordPatterns -wordinstalled $wordinstalled
    }
    'xls' {
        $results = Get-XlsContent -filepath $sysfiles -patterns $passwordPatterns -wordinstalled $wordinstalled
    }
    {$_ -in "odp","ods","odt"} {
        $results = Get-OdsContent -filepath $sysfiles -patterns $passwordPatterns # Même traitement que ods
    }
    'pdf' {
        $results = Get-PdfContent -filepath $sysfiles -patterns $passwordPatterns
    }
    'xml' {
        $results = Get-XmlContent -filepath $sysfiles -patterns $passwordPatterns
    }
    {$_ -in "exe","dll","msi","msu","cab"} {
        $results = Get-ExecutablesContent -filepath $sysfiles   
    }
    {$_ -in "pfx","cer","der"} {
        $results = Get-CertifsContent -filepath $sysfiles
    }
    {$_ -in "p7b","p7c","p7s"} {
        $results = Get-P7bCertContent -filepath $sysfiles
    }
    {$_ -in "bmp","webp","ico","bmp","tif"} {
        $results =  Get-HiddenFilesInImage -filepath $sysfiles
    }
    {$_ -in "jpg","jpeg","png","gif"} {
        $results =  Get-HiddenFilesSpecificInImage -filepath $sysfiles
    }    
    {$_ -in "zip","7z","rar"} {
        $results =  Get-Zipprotectedbypass -filepath $sysfiles -zipinstalled $zipinstalled
    }
    'requires_check' {
        $results = Get-RequiredCheckContent -filepath $sysfiles
    }
    'others' {
        $results = Get-OthersContent -filepath $sysfiles -patterns $passwordPatterns
    }
    'bigsize' {
        $results = Get-checkfilesize -filepath $sysfiles 
    }
}
    
    # Execute the appropriate command based on the detected file type
    return $results
}

if (Is-WordInstalled) {
# Terminate the running Word and Excel processes to prevent double opening
Get-process *winword* -erroraction SilentlyContinue | Stop-Process
Get-Process excel -erroraction SilentlyContinue | Stop-Process
}

$fichiertraite = 0
# Define the array to store inaccessible files
[System.Collections.Generic.List[Object]]$notAccessibleFiles = @()
# Create Jobs 

Get-ChildItem -Path \\$dnsDomain\sysvol -Recurse -File -Include $fileExtensions -Exclude $ignoreExtensions -Force -ErrorAction SilentlyContinue -ErrorVariable notacess | ForEach-Object {

if ($notacess) { 
    Write-Host $notacess -ForegroundColor Red
    $notacess.GetEnumerator() | ForEach-Object {
    $errorDetails = [PSCustomObject]@{
            Error    = $_
        }
        }
        $notAccessibleFiles.Add($errorDetails)
        $notacess.Clear()
 } else {
 
$fichiertraite++
$sysfiles = $_.FullName

#clear
Write-Host scans : $sysfiles -ForegroundColor Cyan
$keepscrenn += "scans :" + $sysfiles 


$runspace = [PowerShell]::Create()
$null = $runspace.AddScript($scriptblock)
$null = $runspace.AddArgument($sysfiles)
$null = $runspace.AddArgument($passwordPatterns)
$null = $runspace.AddArgument($wordinstalled)
$null = $runspace.AddArgument($zipinstalled)
$null = $runspace.AddArgument($module)
$null = $runspace.AddArgument($Maxfilesize)
$null = $runspace.AddArgument($MaxBinarysize)
$null = $runspace.AddArgument($jsonContent)
$runspace.RunspacePool = $pool
$runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
}
}

Write-Host "scan finished prepare analyse ..." -ForegroundColor Green

# Display the slots and current statistic
    while ($runspaces.Status -ne $null) {

    Start-Sleep 2
    Clear-Host

    $slt_tot = Get-Runspace | Where-Object { $_.Id -ne 1 -and $_.RunspaceIsRemote -eq $false }
    $slt_encours = Get-Runspace | Where-Object { $_.Id -ne 1 -and $_.RunspaceAvailability -like "InUse" }
    
    Write-Host "All Objects = " $runspaces.Count
    Write-Host "Total Slots = " $slt_tot.Count
    Write-Host "Used SLots= " $slt_encours.Count
    Write-Host "Remaining objects =" ($runspaces | Where-Object { $_.Status.IsCompleted -eq $false }).Count

    $completed = $runspaces | Where-Object { $_.Status.IsCompleted -eq $true }

    foreach ($runspace in $completed) {
        $Results += $runspace.Pipe.EndInvoke($runspace.Status)
        $runspace.Status = $null
    }
    }

#endregion code

#region summary
Write-Host "Scan completed, calculating statistics" -ForegroundColor Green

# Sort results in unique mode
$Results = $Results | Select-Object -Unique FilePath, pattern, Reason

$sortedGroups = $Results.filepath | Group-Object | Sort-Object -Property Count -Descending 

# Select the first 5 Path
$top5path = $sortedGroups | Select-Object Count,Name -First 5

# Remove commun domain path 
$commonPath = "\\$dnsDomain\sysvol\$dnsDomain"
$top5path = $top5path | ForEach-Object {
    $_.Name = $_.Name -replace [regex]::Escape($commonPath), '' -replace ("\\"), '\\'
    $_
}

# Top 5 words
$top5Words= $Results.pattern | Group-Object | Sort-Object -Property Count -Descending | Select-Object Count,name -First 5

# Skip default GPO policy password settings
$FilteredResults = $Results | Where-Object {
    -not ($_."FilePath" -match '\\Windows NT\\SecEdit\\GptTmpl\.inf' -and $_.Reason -match 'MinimumPasswordAge|MaximumPasswordAge|MinimumPasswordLength|PasswordComplexity|PasswordHistorySize|RequireLogonToChangePassword|ClearTextPassword')
}
$Allwords = $FilteredResults | Group-Object -Property pattern | Sort-Object -Property Count -Descending | Select-Object Count, Name

#Number found objects
$Objectfound = 0
$Objectfound = $sortedGroups.Count

# Group file paths by file extension
$groupedFiles = $sortedGroups.name | Group-Object -Property { ($_ -split "\.")[-1] } | select Count,name

# End of the script: obtain the current date and time
$endDate = Get-Date

# Calculate the time difference
$elapsedTime = New-TimeSpan -Start $startDate -End $endDate
$elapsedTime = $($elapsedTime.ToString("hh\:mm\:ss"))

#region Calcul potentiel risk
# Assume that $top5Groups, $Allwords, and $Objectfound are already defined"
$totalRisk = 0

# Assess the risk based on the number of files containing passwords
if ($Objectfound -gt 20) {
    $totalRisk += 10
} else {
    $totalRisk += ($Objectfound/20) * 10
}

# Iterate through the keywords in $Allwords and adjust the risk score
foreach ($word in $Allwords) {
    switch -Regex ($word.Name) {
        "AutoLogon|cpassword" {
        $totalRisk += 7 * $word.Count
            break
        }
        "Password|Pass|\bpass\b|\bpwd\b" {
            $totalRisk += 3 * $word.Count
            break
        }
        "net use|net user|NotSigned|check required" {
            $totalRisk += 3 * $word.Count
            break
        }
        "sha1|md5" {
            $totalRisk += 5 * $word.Count
            break
        }
        "credentials|\bsecret\b|IPv4" {
            $totalRisk += 2 * $word.Count
            break
        }
    }
}

# Limited score to 100%
if ($totalRisk -gt 100) {
    $totalRisk = 100
}

#endregion Calcul potentiel risk

#endregion Summary

# Close all Slots and pool
$pool.Close()
$pool.Dispose()

#endregion Script

#region HTML

$logo = "https://github.com/dakhama-mehdi/Harden-Sysvol/blob/main/Pictures/HardenSysvol.png?raw=true"
$rightlogo = "https://github.com/dakhama-mehdi/Harden-Sysvol/blob/main/Pictures/Rightlogo.png?raw=true"

# Generate HTML report
Write-Host "Generate HTML" -ForegroundColor Green

[String]$SavePath = $SavePath + '\hardensysvol.html'

New-HTML -TitleText 'HardenSysvol' -FilePath $SavePath -ShowHTML:$true {
    New-HTMLHeader {
    New-HTMLSection -Invisible  {            
            New-HTMLPanel -Invisible {
            New-HTMLText -Text "Domain : $($dnsDomain)" -Alignment left -FontSize 30 -FontWeight 100 -Color Blue
            New-HTMLText -Text "Report date: $startDate" -Alignment left -FontSize 15
            New-HTMLText -Text "Elapsed : $elapsedTime" -Alignment left -FontSize 15 
            } -AlignContentText left
            New-HTMLPanel -Invisible -AlignContentText right {
                New-HTMLImage -Source $logo -Class 'otehr' -Width 30%  -Height 20%
            }

        }
    }
    New-HTMLTab -Name 'Tab 1 : Dashboard' -IconRegular chart-bar  {      
    New-HTMLTabStyle  -BackgroundColorActive Teal  
    New-HTMLSection -Width "80%" -Invisible   {
    New-HTMLSection -Width "40%" -Invisible {
    New-HTMLGage -Label 'Indicator Risk' -MinValue 0 -MaxValue 100 -Value $totalRisk -ValueColor Black -LabelColor Black -Pointer -StrokeColor Akaroa -SectorColors AirForceBlue 
    }
    New-HTMLSection -Width "60%" -name 'Total processed' -HeaderBackGroundColor Teal {
    New-HTMLChart -Gradient -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette4 
                     New-ChartPie -Name 'Total objects' -Value $runspaces.count
                     New-ChartPie -Name 'Found objects' -Value $Objectfound                                   
                }
            }
    New-HTMLSection -Width "60%" -name 'Extensions by type' -HeaderBackGroundColor Teal {
        if ($groupedFiles) {
        New-HTMLChart -Gradient  -TitleAlignment center -Height 200   { 
        New-ChartTheme  -Mode light
        foreach ($grpfiles in $groupedFiles) {
                    New-ChartPie -Name $grpfiles.name -Value $grpfiles.count 
                    }                    
                }
                }       
 } 
    }
    New-HTMLSection -name 'Resume' -HeaderBackGroundColor Teal  {  
    New-HTMLPanel  {
    if ($top5path) {
            New-HTMLChart -Title 'Top 5 Files' -TitleAlignment center -Height "140%" {
            $legendNames = @()
            $chartValues = @()
            foreach ($word in $top5path) {
                $legendNames += $word.name
                $chartValues += $word.count
            }
            New-ChartToolbar -Download pan                   
            New-ChartLegend -Name $legendNames -HideLegend
            New-ChartBarOptions -Type barStacked 
            New-ChartBar -Name 'Path' -Value $chartValues
            }
            }
    else {  New-HTMLText -FontSize 20px -Text "<br><br>Top files<br>No data" -Alignment center -Color Grey }
      }          
    New-HTMLPanel  {
    if ($top5Words) {               
    New-HTMLChart -Title 'Top 5 Reason' -TitleAlignment Center -Height 100% {
            $legendNames = @()
            $chartValues = @()
            foreach ($word in $top5Words) {
                $legendNames += $word.name
                $chartValues += $word.count
            }               
            New-ChartToolbar -Download -Pan
            New-ChartBarOptions -Gradient         
            New-ChartLegend -Name $legendNames -HideLegend 
            New-ChartBar -Name 'Pattern' -Value $chartValues          
            } 
            } 
    else {  New-HTMLText -FontSize 20px -Text "<br><br>Top reason<br>No data" -Alignment center -Color Grey }
            }
    New-HTMLPanel -Width "70%" {
         New-HTMLList {
              New-HTMLListItem -Text 'Harden-Sysvol _ Version : 1.7 _ Release : 10/2024' 
              New-HTMLListItem -Text 'Author : Dakhama Mehdi<br>
              <br> Credit : HardenAD Community [HardenAD](https://www.hardenad.net/)
              <br> Credit : It-connect Community [It-Connect](https://www.it-connect.fr/)
              <br> Thanks : Przemyslaw Klys  [Evotec](https://evotec.xyz) for Module PSWriteHTML/PswriteOffice'
              } -FontSize 14
            }    

    } 
    New-HTMLSection -Width "60%"  -HeaderBackGroundColor Teal -name 'Tips & Best pratices'  {
        New-HTMLPanel -Width "60%" {
            New-HTMLImage -Source $rightlogo 
        }       
    New-HTMLPanel  {         
    New-HTMLTabPanel -Orientation vertical -Theme 'pills' -AutoProgress -TransitionSpeed 1   {
                    New-HTMLTab -Name 'Why check Sysvol' -IconBrands 500px  {
                        New-HTMLText -FontSize 18px -Text "The Sysvol folder is crucial for distributing scripts and Group Policy Objects (GPOs) to all domain computers. 
                        It may contain sensitive information, such as plain-text passwords, making it a prime target for attackers. 
                        <br>A vulnerability in Sysvol can compromise the entire domain. Therefore, it is essential to restrict permissions, monitor changes, 
                        and regularly audit its contents to ensure network security and compliance."
                    }
                    New-HTMLTab -Name 'Audit GPO' -IconBrands 500px {
                        New-HTMLText -FontSize 18px -Text "Regularly audit GPOs to verify their contents, such as plain-text passwords in configuration files or auto-logon scripts, 
                        and the presence of unsigned sources. <br>Frequently run the GPOZaurr tool, which provides a comprehensive report to help identify and mitigate these risks.
                        <br>[GPOZaurr](https://github.com/EvotecIT/GPOZaurr/)<br>"
                    }
                    New-HTMLTab -Name 'Best Pratic' -IconBrands 500px {
                        New-HTMLText -FontSize 18px -Text "Enable audits on the Sysvol folder and monitor logs for multiple search attempts, as this may indicate enumeration attempts. 
                        Some elements in the Sysvol folder are not meant to be accessed by everyone. If possible, place a honeypot script in the Netlogon folder to trigger alerts for suspicious activity.
                        <br>[Autologon](https://learn.microsoft.com/fr-fr/sysinternals/downloads/autologon/)<br>"
                    }
                    New-HTMLTab -Name 'Tips ' -IconBrands 500px {
                        New-HTMLText -FontSize 18px -Text "Do not store large files, such as ISO or .zip files, in the Sysvol folder. This can lead to replication issues and unnecessary consumption of storage resources, impacting the performance and reliability of your network<br> Move your scripts to a shared folder and grant access only to the relevant groups, not authenticated users. This will reduce vulnerabilities, especially if the scripts contain credentials or deploy critical applications."
                    }
                    New-HTMLTab -Name 'Hardening AD' -IconBrands 500px {
                        New-HTMLText -FontSize 18px -Text "Use AD hardening to ensure security and reduce risks. <br>Disable old protocols like SMB1 and anonymous enumeration on DC shares. 
                        <br>Implement an N-tier architecture model, a PAW, and Silos. 
                        To facilitate this, refer to the HardenAD project.
                        <br>[HardenAD](https://github.com/LoicVeirman/HardenAD/)<br>"
                    }
                    New-HTMLTab -Name 'Help' -IconBrands 500px {
                        New-HTMLText -FontSize 18px -Text "Use this command to improve the research : invoke-hardensysvol -Allextensions -addpattern admin -Maxfilesize 1
                        <br>Link to doc. 
                        You can support the project
                        <br>[Documentation](https://github.com/dakhama-mehdi/Harden-Sysvol/tree/main/Docs)<br>"
                    }
                }
    } 
   }    
    }
    New-HTMLTab -Name 'Tab 2 : Details' -IconSolid user-alt   {     
    New-HTMLSection  -Invisible  {
    New-HTMLTableOption -DataStore JavaScript 
    New-htmlTable -HideFooter -DataTable $Results -TextWhenNoData 'Information: No sentivity data found'
        }
    New-HTMLSection -HeaderBackGroundColor CarrotOrange -HeaderText 'Errors log' {
    New-HTMLTableOption -DataStore JavaScript 
    New-htmlTable -HideFooter -DataTable $notAccessibleFiles -TextWhenNoData 'No errors during scanning'
        }
             }  
}

#endregion HTML

}

# SIG # Begin signature block
# MIImVgYJKoZIhvcNAQcCoIImRzCCJkMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB9mJX0kaixC1tR
# YhGBsm2Q6OgezCpMU2dvymgtWgG7d6CCH+wwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggYoMIIEEKADAgECAhBrxlWg9go45bxtH9Zi+WCgMA0GCSqG
# SIb3DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0YSBT
# eXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBD
# QTAeFw0yNDExMDYwOTQ0MjlaFw0yNTExMDYwOTQ0MjhaME4xCzAJBgNVBAYTAkZS
# MQ8wDQYDVQQHDAZUb3Vsb24xFjAUBgNVBAoMDU1laGRpIERha2hhbWExFjAUBgNV
# BAMMDU1laGRpIERha2hhbWEwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAwggGKAoIB
# gQCsFc3e5PwEJuycVRR54Qp8hFEckVwj7u1hMc7fejXKC/oR+uixlujLAHA9NcGX
# jcQIXNP3GmezLF3Tj6Jvcs/kNT/a5zqjI5HEfIap7EHwf03f5060+Rc21v1UDjzj
# DZzi9xFFum8eeGLc4pTzUB3wP3+M+mY7d5QlTjIxZSNnMBisJE8ASqG9JtRcQmIz
# HACI70xRCQVV8ZjJ8J+Shr6wkNdDy/IjR+Y9VkMRIJozWR+pqbKuQOIDBSxQYVHg
# bT+gsLOfvHkBPJN0ZQe7eJdG7J78Z1nzNH9yXhZ0HHdPB80tUwM0HC1n4LO3kki/
# IBmg4Qq/UyMMQd826fJk3ylbAlf8w7N80INQcLLBGVECmWI21d9f3l5usvWDa+mJ
# ma57c6GUDY05Jg5owLgNREZsyRt5rOlg68NLmz9tuEkJA1D4ntpKq0KZc3HJv04x
# XTcfTEqbKYr7vZ//ENsell5UdUQxL6rGJzazhsK02ZcmasICiHNLfG/tBaolCbeM
# 8ekCAwEAAaOCAXgwggF0MAwGA1UdEwEB/wQCMAAwPQYDVR0fBDYwNDAyoDCgLoYs
# aHR0cDovL2Njc2NhMjAyMS5jcmwuY2VydHVtLnBsL2Njc2NhMjAyMS5jcmwwcwYI
# KwYBBQUHAQEEZzBlMCwGCCsGAQUFBzABhiBodHRwOi8vY2NzY2EyMDIxLm9jc3At
# Y2VydHVtLmNvbTA1BggrBgEFBQcwAoYpaHR0cDovL3JlcG9zaXRvcnkuY2VydHVt
# LnBsL2Njc2NhMjAyMS5jZXIwHwYDVR0jBBgwFoAU3XRdTADbe5+gdMqxbvc8wDLA
# cM0wHQYDVR0OBBYEFAG3sIcT8bRm7QyFu8699Gpkr5NmMEsGA1UdIAREMEIwCAYG
# Z4EMAQQBMDYGCyqEaAGG9ncCBQEEMCcwJQYIKwYBBQUHAgEWGWh0dHBzOi8vd3d3
# LmNlcnR1bS5wbC9DUFMwEwYDVR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCJ58BnchFNGzLksJ9oHFEWTs643G+PKOHr
# 9RmrKSB/4MtPriG5iez+MFsGqYwkYd5QzqOIYg24ctfbWXJWG8Xj+YMfp1r+hkYq
# O0Abpv26sZ1ZjNGgGUbb3z7KqhY+IdVpZf2aG/Rycl5dE2LbhWqp9h24WfQCIS/e
# XxH7HmM9SEBHYbfOqlEA+RF/gRGYCQOg0ui2j0ZzIOrQGj3Njn/5rzP9OmPmLt4h
# DsixjFWgu598XmRKj5KW1MShFIjUuUzSmOWDgKA16lJi6LggdFAB/MImiDH48v8N
# /9R9En24pUGGj2XOgBX5SZ4kj+VN1YaY1vYPFp3wLu85zpgRZgZQC+WurX8s1tRn
# iCIj/+ajUB4G4TcbTz6k16X1Yz9ba1y7p/hJB92uDW7esMGgqzEv+Osd11bVoNmv
# CE8Twsz0cuFJqBtVZIycCkgw/AVyJIsNS6RADi94PvbOf8rty8HV3bHmm6O4wJVc
# 5ch50bL7JVyYTPN5OTzXSDx62wKi5ePZvEF7RX3cQlTQMYticde91khs2n2FZ06K
# Uin5DtQgxy0Q1ufFIDZthsk5AaSWiZzFgAgJt8JaQGPyGAYl2Sr8a/gMLpcBsPwI
# zdlDUOJwyHPxlR9ZiraUzF/1SSN7CgjqFSDAAZ+i4i8gZsPpU38GtBSLrw/CrnUB
# /KGcFNMvszCCBq4wggSWoAMCAQICEAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcN
# AQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3Rl
# ZCBSb290IEc0MB4XDTIyMDMyMzAwMDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkG
# A1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdp
# Q2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8Ty
# kTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsm
# c5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTn
# KC3r07G1decfBmWNlCnT2exp39mQh0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2
# R/dhgxndX7RUCyFobjchu0CsX7LeSn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0
# QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/
# oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1ps
# lPJSlRErWHRAKKtzQ87fSqEcazjFKfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhI
# fxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8
# I41Y99xh3pP+OcD5sjClTNfpmEpYPtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkU
# EBIDfV8ju2TjY+Cm4T72wnSyPx4JduyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1G
# nrXTdrnSDmuZDNIztM2xAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB0GA1UdDgQWBBS6FtltTYUvcyl2mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC
# 0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQu
# Y3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0B
# AQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7
# cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2p
# Vs8Vc40BIiXOlWk/R3f7cnQU1/+rT4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxk
# Jodskr2dfNBwCnzvqLx1T7pa96kQsl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpkn
# G6skHibBt94q6/aesXmZgaNWhqsKRcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2
# n82HhyS7T6NJuXdmkfFynOlLAlKnN36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fm
# w0HNT7ZAmyEhQNC3EyTN3B14OuSereU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvt
# Cl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU
# 5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8K
# vYHZE/6/pNHzV9m8BPqC3jLfBInwAM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/
# GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9EFUrnEw4d2zc4GqEr9u3WfPwwgga5MIIE
# oaADAgECAhEAmaOACiZVO2Wr3G6EprPqOTANBgkqhkiG9w0BAQwFADCBgDELMAkG
# A1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAl
# BgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMb
# Q2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIxOFoXDTM2
# MDUxODA1MzIxOFowVjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRh
# IFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIx
# IENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAnSPPBDAjO8FGLOcz
# cz5jXXp1ur5cTbq96y34vuTmflN4mSAfgLKTvggv24/rWiVGzGxT9YEASVMw1Aj8
# ewTS4IndU8s7VS5+djSoMcbvIKck6+hI1shsylP4JyLvmxwLHtSworV9wmjhNd62
# 7h27a8RdrT1PH9ud0IF+njvMk2xqbNTIPsnWtw3E7DmDoUmDQiYi/ucJ42fcHqBk
# bbxYDB7SYOouu9Tj1yHIohzuC8KNqfcYf7Z4/iZgkBJ+UFNDcc6zokZ2uJIxWgPW
# XMEmhu1gMXgv8aGUsRdaCtVD2bSlbfsq7BiqljjaCun+RJgTgFRCtsuAEw0pG9+F
# A+yQN9n/kZtMLK+Wo837Q4QOZgYqVWQ4x6cM7/G0yswg1ElLlJj6NYKLw9EcBXE7
# TF3HybZtYvj9lDV2nT8mFSkcSkAExzd4prHwYjUXTeZIlVXqj+eaYqoMTpMrfh5M
# CAOIG5knN4Q/JHuurfTI5XDYO962WZayx7ACFf5ydJpoEowSP07YaBiQ8nXpDkNr
# UA9g7qf/rCkKbWpQ5boufUnq1UiYPIAHlezf4muJqxqIns/kqld6JVX8cixbd6Pz
# kDpwZo4SlADaCi2JSplKShBSND36E/ENVv8urPS0yOnpG4tIoBGxVCARPCg1BnyM
# J4rBJAcOSnAWd18Jx5n858JSqPECAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMB
# Af8wHQYDVR0OBBYEFN10XUwA23ufoHTKsW73PMAywHDNMB8GA1UdIwQYMBaAFLah
# VDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggr
# BgEFBQcDAzAwBgNVHR8EKTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9j
# dG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1
# YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRv
# cnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYI
# KwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwF
# AAOCAgEAdYhYD+WPUCiaU58Q7EP89DttyZqGYn2XRDhJkL6P+/T0IPZyxfxiXumY
# lARMgwRzLRUStJl490L94C9LGF3vjzzH8Jq3iR74BRlkO18J3zIdmCKQa5LyZ48I
# fICJTZVJeChDUyuQy6rGDxLUUAsO0eqeLNhLVsgw6/zOfImNlARKn1FP7o0fTbj8
# ipNGxHBIutiRsWrhWM2f8pXdd3x2mbJCKKtl2s42g9KUJHEIiLni9ByoqIUul4Gb
# lLQigO0ugh7bWRLDm0CdY9rNLqyA3ahe8WlxVWkxyrQLjH8ItI17RdySaYayX3Ph
# RSC4Am1/7mATwZWwSD+B7eMcZNhpn8zJ+6MTyE6YoEBSRVrs0zFFIHUR08Wk0ikS
# f+lIe5Iv6RY3/bFAEloMU+vUBfSouCReZwSLo8WdrDlPXtR0gicDnytO7eZ5827N
# S2x7gCBibESYkOh1/w1tVxTpV2Na3PR7nxYVlPu1JPoRZCbH86gc96UTvuWiOruW
# myOEMLOGGniR+x+zPF/2DaGgK2W1eEJfo2qyrBNPvF7wuAyQfiFXLwvWHamoYtPZ
# o0LHuH8X3n9C+xN4YaNjt2ywzOr+tKyEVAotnyU9vyEVOaIYMk3IeBrmFnn0gbKe
# TTyYeEEUz/Qwt4HOUBCrW602NCmvO1nm+/80nLy5r0AZvCQxaQ4wgga8MIIEpKAD
# AgECAhALrma8Wrp/lYfG+ekE4zMEMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYT
# AlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQg
# VHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQw
# OTI2MDAwMDAwWhcNMzUxMTI1MjM1OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UE
# ChMIRGlnaUNlcnQxIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo
# /ZEfGMSIO2qZ46XB/QowIEMSvgjEdEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1cz
# SzvUQ5xF7z4IQmn7dHY7yijvoQ7ujm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJ
# oJqAsP8YuhRvflJ9YeHjes4fduksTHulntq9WelRWY++TFPxzZrbILRYynyEy7rS
# 1lHQKFpXvo2GePfsMRhNf1F41nyEg5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66Y
# X2LZPxS4oaf33rp9HlfqSBePejlYeEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXP
# lKdE4fBIn5BBFnV+KwPxRNUNK6lYk2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxK
# aXN12HgR+8WulU2d6zhzXomJ2PleI9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt
# 92nm7Mheng/TBeSA2z4I78JpwGpTRHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4R
# rcnKJ3FbjyPAGogmoiZ33c1HG93Vp6lJ415ERcC7bFQMRbxqrMVANiav1k425zYy
# FMyLNyE1QulQSgDpW9rtvVcIH7WvG9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDL
# nBHXgYly/p1DhoQo5fkCAwEAAaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNV
# HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGog
# j57IbzAdBgNVHQ4EFgQUn1csA3cOKBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBP
# oE2gS4ZJaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0
# UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMw
# gYAwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEF
# BQcwAoZMaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3Rl
# ZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsF
# AAOCAgEAPa0eH3aZW+M4hBJH2UOR9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7
# xCKhVireKCnCs+8GZl2uVYFvQe+pPTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpW
# O9QGFwfMEy60HofN6V51sMLMXNTLfhVqs+e8haupWiArSozyAmGH/6oMQAh078qR
# h6wvJNU6gnh5OruCP1QUAvVSu4kqVOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJ
# Lsxtvge/mzA75oBfFZSbdakHJe2BVDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehS
# R7vM+C13v9+9ZOUKzfRUAYSyyEmYtsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1Hv
# IiulqJ1Elesj5TMHq8CWT/xrW7twipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vE
# bFkEiF2abhuFixUDobZaA0VhqAsMHOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75
# KiNbh0c+hatSF+02kULkftARjsyEpHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3B
# pIiowOIIuDgP5M9WArHYSAR16gc0dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/
# Z5lGLvNwQ7XHBx1yomzLP8lx4Q1zZKDyHcp4VQJLu2kWTsKsOqQxggXAMIIFvAIB
# ATBqMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0ZW1z
# IFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQQIQa8ZV
# oPYKOOW8bR/WYvlgoDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDWEJy5JOxiu1z8Lzek
# WH85GeOAlad06kcVMneRmXPGqDANBgkqhkiG9w0BAQEFAASCAYCK+dIAWcxZ+1rm
# EL67XDeuTHBO6WBGKI6SlwVJJq2QXYAsLF4daBoYt6iD9OwoimE1TwLtWCmdqAU9
# VrEvD4komSeKICscjNDBaDjhGFBPBLrZuwobuCztZNHoFRS82TxWJ7isPFZZaoGM
# MKS8LQr3auNiy+i0w5T5286NkslsC97OZA3bzw7o4jERl29HaxyHcWkGb+iISkwx
# Qd1DXAEEDsrCSHYlyDWUIqm1mGdjZyMrLzzDCiBv+Xl1hhfl93jNeaTg/lM4znTi
# jObGpPjOjrZYZ2lKMyZAKBHPgqmxSS7Pbq0CSJ65DrrZXy/p3crCWyN+Aip8S0Q4
# 4t8OU1KfZufdymof12xZfEr+awVDoRR8kervByxQ0UvVKLLTlYHiVNYuYQhNrI6x
# Iz0qT2U/qWe7GhetatSLon3mbFYHS45wUItKiPTYYieizQExu1sumttXFt213+lg
# 0IMfCLbz7igtbe8/6zxiMK7RI7qKL67mst30s3epNxhcS/BO/XOhggMgMIIDHAYJ
# KoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNB
# NDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhALrma8Wrp/lYfG+ekE4zMEMA0G
# CWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjUwMTIwMjA0MDU2WjAvBgkqhkiG9w0BCQQxIgQgQOzjZFPjssA2
# c2LnNhzwjlM7pgEwpM93mxx4szkzILwwDQYJKoZIhvcNAQEBBQAEggIAoz27pRpq
# MNzIP9OQqP0FyU2eKT9M4mHVixT+7RbLs77dEdDQg1dWrOAoSuXbN6HW7LnmgjN5
# a70xw6nW4Omkeu3Iu1zNi2n+M3QmsgFeXAUK9W8E3gUFfW/jlIAp8oSWtGI/P9Zb
# jkpSeJMutRBocM+K/297oHDyUqw9jRRI4UA2Qo1hMMIHU8DKov9u0G0GAUKEDK5o
# tEx0hQEqLIkMKY0nCzwiHBmRywpozNUefaC58lUwbLtfVUcXe4O5zsG6aFH91n3V
# 83xOuK5U+5zYF75bWJt5UuCwJAarnxdZsfjbGFWnfGw7C2vkNvTMH2obth3y/CoH
# Zys80oTnLp+rRy76muE544HI1rljLTM5WsVR+t7GFjb4NIGCcJ4JIQ5JKBOFIh+u
# umK4B2p2jTp+OHFhELih/bXbc5rWGkTPOb250hQAqKBF8pco8nTso1mA7/bmermJ
# ZvbV2FErbTiEvBmQTWzPEDpR5rIuqDlwL9SKf3gPJuuy7iYBLdngHP9K/f2/5piz
# +TeRfo4jM4HrYE3RmoewFADYWz4yJx9med6fSlw0ppBx7LfgOgYWpMb3bDwL1X9q
# L3EgjrpD5jC2pgm1iEl0rK3S4RfJ+ChxN1lpHTpGlHs5JO5EE1dIE7AdifqBpe2N
# vdQTmwKEs+BvaQYSFqRhnrhihCObQThlM9s=
# SIG # End signature block
