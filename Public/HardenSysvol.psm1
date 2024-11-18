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
            install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Stop
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
        [object]$jsonContent
    )

    # Import modul FileHandlers.psm1
    Import-Module $module -Verbose    

    [String]$detectedType =  Get-FileType -filePath $sysfiles -jsonContent $jsonContent       
       
    # Function to search pattern by extensions
    switch ($detectedType)  {
    'docx' {
        $results = Get-DocxContent -filePath $sysfiles -patterns $passwordPatterns
    }
    'xlsx' {
        $results = Get-XlsxContent -filepath $sysfiles -patterns $passwordPatterns
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
    {$_ -in "jpg","jpeg","bmp","webp","png","ico","gif","bmp","tif"} {
        $results =  Get-HiddenFilesInImage -filepath $sysfiles
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
$notAccessibleFiles = @()
# Create Jobs 

Get-ChildItem -Path \\$dnsDomain\sysvol -Recurse -File -Include $fileExtensions -Exclude $ignoreExtensions -Force -ErrorAction SilentlyContinue -ErrorVariable notacess | ForEach-Object {

if ($notacess) { 
    Write-Output $notacess -ForegroundColor Red
    $errorDetails = [PSCustomObject]@{
            FilePath = $_.FullName
            Error    = $notacess
        }
        $notAccessibleFiles += $errorDetails
        $notacess = $null
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
    New-HTMLTabPanel -Orientation vertical -Theme 'pills'  {
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

