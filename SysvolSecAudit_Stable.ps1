<#
.SYNOPSIS
    Sysvol Security Audit Module

.DESCRIPTION
    This module is designed to scan the Sysvol folder for files containing sensitive information 
    such as passwords and usernames. It helps in identifying potential security risks within the Sysvol directory.

.VERSION
    1.3.0

.Contribution
    Loic, Bastien, Florian, Thirrey

.AUTHOR
    Dakhama

.PARAMETER dnsDomain
    Specifies the DNS domain to be scanned. Defaults to the current user's DNS domain if not provided.

.PARAMETER ignoreExtensions
    Specifies file extensions to ignore during the scan.

.EXAMPLE
    # Scan the Sysvol folder of the current domain
    Invoke-SysvolSecurityAudit

    # Scan the Sysvol folder of a specific domain, ignoring .txt and .log files
    Invoke-SysvolSecurityAudit -dnsDomain "example.com" -ignoreExtensions "*.txt", "*.log"

.NOTES
    This script requires administrative privileges to access and scan the Sysvol directory.

.LINK
    https://github.com/dakhama-mehdi/CheckSysvolsecurity
#>

function Invoke-SysvolAudit {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$dnsDomain = $env:USERDNSDOMAIN,
        
        [Parameter(Mandatory = $false)]
        [String[]]$ignoreExtensions
    )

#region code

# Tester l'acces au partage
$testpath = Test-Path "\\$dnsDomain\sysvol\"
if ($testpath -eq $false) {
throw "Cannot access domain or share, pls check with GCI $dnsDomain"
}

# Tester la présence de module
# Tester la présence des modules
$modulesToCheck = "PSWriteOffice", "PSWritePDF", "PSWriteHTML"
foreach ($module in $modulesToCheck) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Output "installation du module $module"
        Install-Module -Name $module -Force -Scope CurrentUser
    } else {
    Import-Module PSWriteHTML -Force
}
}

# Suppression de runspace "fantomes"
$PreviousRS = get-runspace | where-object {($_.id -ne 1)} 
if ($PreviousRS) { $PreviousRS.dispose() }

$Results =  $null
# Début du script : obtenir la date et l'heure actuelles
$startDate = Get-Date

 #region getxmlcontenant
    # Récupérer les extensions de fichiers à partir du fichier XML
    $xmlFileExtensions = "C:\temp\file_extensions.xml"
    $extensionsXML = [xml](Get-Content $xmlFileExtensions -Encoding UTF8)
    $fileExtensions = $extensionsXML.root.FileExtensions.Extension

    # Récupérer les motifs de mots de passe à partir du fichier XML
    $passwordPatterns = $extensionsXML.root.PasswordPatterns.Pattern

    # Initialiser une liste pour stocker les résultats
    $Results = @()

    #endregion getxmlcontenant

# définition du Pool (creation des slots)
$pool = [RunspaceFactory]::CreateRunspacePool(1,10)
$pool.ApartmentState = "MTA"
$pool.Open()
$runspaces = @()


$scriptblock = {
    Param (
        [string]$sysfiles,
        [string[]]$passwordPatterns
    )

    # Définir les commandes pour chaque type de fichier
    $fileCommands = @{
        '.docx' = {
            $results = @()  # Initialisation de la variable $results dans le scriptblock
            Import-Module PSWriteOffice 
            $document = Get-OfficeWord -FilePath $sysfiles -ReadOnly
            foreach ($pattern in $passwordPatterns) {
                $var = $document.Find($pattern) | Select-Object -First 1
                if ($var) {
                    $result = [PSCustomObject]@{
                        FilePath = $sysfiles
                        pattern  = $pattern
                        Word     = ''
                    }
                    $results += $result
                }
            }
            Close-OfficeWord -Document $document
            $results  # Retourne tous les résultats à la fin du scriptblock
        }
        '.xlsx' = {
            $results = @()  # Initialisation de la variable $results dans le scriptblock
            Import-Module PSWriteOffice -Force -Global -Scope Local
            $excel = Get-OfficeExcel -FilePath $sysfiles

            foreach ($pattern in $passwordPatterns) {
                $var = $excel.Search($pattern) | Select-Object -First 1

                if ($var) {
                    $result = [PSCustomObject]@{
                        FilePath = $sysfiles
                        pattern  = $pattern
                        Word     = ''
                    }
                    $results += $result
                }
            }
       $results  # Retourne tous les résultats à la fin du scriptblock
        }
        '.pdf' = {
            $results = @()  # Initialisation de la variable $results dans le scriptblock
            Import-Module PSWritePDF
            $text = (Convert-PDFToText -FilePath $sysfiles) -join "`n"
            
            foreach ($pattern in $passwordPatterns) {
                if ($text -match $pattern) {
                    $result = [PSCustomObject]@{
                        FilePath = $sysfiles
                        pattern  = $pattern
                        Word     = ''
                    }
                    $results += $result
                }
            }
      $results  # Retourne tous les résultats à la fin du scriptblock
        }
        '.xml' = {
    $results = @()  # Initialisation de la variable $results dans le scriptblock
    $xmlContent = Get-Content -Path $sysfiles -Raw

    # Vérifier si le fichier contient "cpassword"
    if ($xmlContent -match 'cpassword') {
        $xml = [xml]$xmlContent

        # Vérifier le contenu du fichier pour déterminer comment le traiter
        if ($xml.Groups.User) {
            $cpasswords = $xml | Select-Xml "/Groups/User/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        } elseif ($xml.NTServices.NTService) {
            $cpasswords = $xml | Select-Xml "/NTServices/NTService/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        } elseif ($xml.ScheduledTasks.Task) {
            $cpasswords = $xml | Select-Xml "/ScheduledTasks/Task/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        } elseif ($xml.DataSources.DataSource) {
            $cpasswords = $xml | Select-Xml "/DataSources/DataSource/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        } elseif ($xml.Printers.SharedPrinter) {
            $cpasswords = $xml | Select-Xml "/Printers/SharedPrinter/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        } elseif ($xml.Drives.Drive) {
            $cpasswords = $xml | Select-Xml "/Drives/Drive/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        }

        # Ajouter les résultats à la liste des résultats
            $result = [PSCustomObject]@{
                FilePath  = $sysfiles
                Pattern  = 'cpassword'
                Word     = $cpasswords
        }
        $results += $result
    } 
    elseif ($xmlContent -match 'DefaultUserName') {
        # Autre traitement pour les fichiers XML sans cpassword

        $xml = [xml]$xmlContent

        # Initialiser les variables
        $userName = ""
        $password = ""

        # Vérifier et extraire les valeurs
        foreach ($registry in $xml.RegistrySettings.Registry) {
        if ($registry.Properties.name -eq "DefaultUserName") {
            $userName = $registry.Properties.value
        }
        if ($registry.Properties.name -eq "DefaultPassword") {
            $password = $registry.Properties.value
        }
    }

        # Retourner les résultats
        $result  = [PSCustomObject]@{
        FilePath = $sysfiles
        Pattern  = 'AutoLogon'
        Word     = "$userName : $password"
        }

        $results += $result
    }
    else {
        $xml = [xml]$xmlContent
        foreach ($pattern in $passwordPatterns) {
            $matches = $xmlContent | Select-String -Pattern $pattern
            foreach ($match in $matches) {
                $result = [PSCustomObject]@{
                    FilePath = $sysfiles
                    Pattern  = $pattern
                    Word     = $match.Matches.Value
                }
                $results += $result
            }
        }
    }

    return $results  # Retourne tous les résultats à la fin du scriptblock
}
        '.exe' = {         
        $results = @()  # Initialisation de la variable $results dans le scriptblock

        if ( (Get-AuthenticodeSignature -FilePath $sysfiles).status -ne "Valid") {
         $result = [PSCustomObject]@{
                FilePath = $sysfiles
                pattern  = "NotSigned"
                Word     = 'NotSigned'
            }    
            $results += $result         
        }   
        $results            
         }
        default = {
            $results = @()  # Initialisation de la variable $results dans le scriptblock
            $matches = Select-String -Path $sysfiles -Pattern $passwordPatterns -AllMatches
            $n= 0

            foreach ($match in $matches) {
            $patternMatch = $match.Matches.Value
            [string]$word= $match.line.Replace($patternMatch,'')

           #if ($word -replace '^[=:"]', '' -ne '') {
           if (-not([string]::IsNUllOrEmpty($word -replace '^[=:"]'))) { 
            $n++
            $result = [PSCustomObject]@{
                FilePath = $sysfiles
                pattern  = $patternMatch
                Word     = $match.Line
            }

            $results += $result
            }
            
            if ($n -eq 4) {
            break;
            }
        }

        $results  # Retourne tous les résultats à la fin du scriptblock
        }
    }

   
    # Obtenir le nom du fichier
    $fileName = Split-Path -Leaf $sysfiles

    # Utiliser un switch pour exécuter la commande appropriée en fonction du nom du fichier
    switch -Wildcard ($fileName) {
    '*.doc*'  { & $fileCommands['.docx'] }
    '*.exe'  { & $fileCommands['.exe'] }
    '*.msi'  { & $fileCommands['.exe'] }
    '*.xlsx*' { & $fileCommands['.xlsx'] }
    '*.pdf*'  { & $fileCommands['.pdf'] }
    '*.xml*'  { & $fileCommands['.xml'] }
    default   { & $fileCommands['default'] }
    }

}

$fichiertraite = 0
# creation des jobs par machine et lancement des jobs

Get-ChildItem -Path  \\$dnsDomain\sysvol -Recurse -File -Include $fileExtensions -Exclude $ignoreExtensions -Force -ErrorAction SilentlyContinue -ErrorVariable notacess  | ForEach-Object {

if ($notacess) { Write-Host $notacess -ForegroundColor Red; $notacess = $null }

$fichiertraite++
$sysfiles = $_.FullName

#clear
Write-Host Scanne : $sysfiles -ForegroundColor Cyan

$runspace = [PowerShell]::Create()
$null = $runspace.AddScript($scriptblock)
$null = $runspace.AddArgument($sysfiles)
$null = $runspace.AddArgument($passwordPatterns)
$runspace.RunspacePool = $pool
$runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
}


# Afficher les slots + statistiques en cours
while ($runspaces.Status -ne $null)
{
start-sleep 1
cls
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)  -and ($_.runspaceAvailability -like "Available")}
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $true)}
$slt_encours = get-runspace | where-object {($_.id -ne 1) -and ( $_.runspaceisremote -eq $true)}
$slt_tot = get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)}

write-host  "Nbre Objets total= " $runspaces.count
write-host  "Nbre slots totaux= " $slt_tot.count
write-host  "Nbre slots utilisés= " $slt_encours.count
write-host  "Nbre objets restants =" $runspaces.Status.IsCompleted.count

$completed = $runspaces | Where-Object { $_.Status.IsCompleted -eq $true }

foreach ($runspace in $completed)
{
    $Results += $runspace.Pipe.EndInvoke($runspace.Status)
    $runspace.Status = $null
}

}

#region Summary
$sortedGroups = $Results.filepath | Group-Object | Sort-Object -Property Count -Descending

# Sélectionner les 5 premiers groupes
$top5Groups = $sortedGroups | Select-Object Count,Name -First 5

# Afficher les 5 premiers groupes
$commonPath = "\\$dnsDomain\sysvol\$dnsDomain"
$top5Groups = $top5Groups | ForEach-Object {
    $_.Name = $_.Name -replace [regex]::Escape($commonPath), ''
    $_
}

# Top 5 word
$top5Words= $Results.pattern | Group-Object | Sort-Object -Property Count -Descending | Select-Object Count,name -First 5
$Allwords = $Results.pattern | Group-Object | Sort-Object -Property Count -Descending | Select-Object Count,name

#nombre d'objet max 
$objettrouve = $sortedGroups.Count

# Regrouper les chemins de fichier par extension de fichier
$groupedFiles = $sortedGroups.name | Group-Object -Property { ($_ -split "\.")[-1] } | select Count,name

# Fin du script : obtenir la date et l'heure actuelles
$endDate = Get-Date

# Calculer la différence de temps
$elapsedTime = New-TimeSpan -Start $startDate -End $endDate
$elapsedTime = $($elapsedTime.ToString("hh\:mm\:ss"))

# Calcul potentiel risk
# Supposons que $top5Groups, $Allwords et $objettrouve soient déjà définis
$totalRisk = 0

# Évaluer le risque basé sur le nombre de fichiers contenant des mots de passe
if ($objettrouve -gt 10) {
    $totalRisk += 30
} else {
    $totalRisk += ($objettrouve/10) * 30
}

# Parcourir les mots-clés dans $Allwords et ajuster le score de risque
foreach ($word in $Allwords) {
    switch -Regex ($word.Name) {
        "Password|Pass|motdepasse|\bpass\b|\bpwd\b" {
            $totalRisk += 5 * $word.Count
            break
        }
        "cpassword" {
            $totalRisk += 20 * $word.Count
            break
        }
        "net use|NotSigned|\bidentifiant\b" {
            $totalRisk += 5 * $word.Count
            break
        }
        "AutoLogon" {
            $totalRisk += 20 * $word.Count
            break
        }
        "credentials|\bsecret\b" {
            $totalRisk += 5 * $word.Count
            break
        }
    }
}

# Limiter le score de risque à 100%
if ($totalRisk -gt 100) {
    $totalRisk = 100
}

#endregion Summary

# Fermeture les connexions et suppression des slots du Pool
$pool.Close()
$pool.Dispose()

#endregion code

# Generation du rapport HTML
New-HTML -TitleText 'AD_ModernReport' -ShowHTML {
    New-HTMLHeader {
        New-HTMLText -LineBreak
        New-HTMLSection -Invisible  {
            
            New-HTMLPanel -Invisible {
            New-HTMLText -LineBreak 
            New-HTMLText -Text "Sysvol ClearPass Check : $($dnsDomain)" -Alignment left -FontSize 30 -FontWeight bold -Color Blue
            New-HTMLText -Text "Report date: $startDate" -Alignment left -FontSize 15
            New-HTMLText -Text "Elapsed : $elapsedTime" -Alignment left -FontSize 15 
            } -AlignContentText left

            New-HTMLPanel -Invisible -AlignContentText right {
                New-HTMLImage -Source 'C:\temp\logo.png' -Class 'otehr' -Width '30%'
            }

        }
        New-HTMLText -LineBreak
    }   
    New-HTMLTab -Name 'Dashboard' -IconRegular chart-bar  {
    New-HTMLSection {
    New-HTMLTableOption -DataStore JavaScript 
    New-htmlTable -HideFooter -DataTable $Results -TextWhenNoData 'Information: No Groups were found'
        }
    New-HTMLSection {    
    New-HTMLPanel  {
    New-HTMLPanel -Width "60%" {
                New-HTMLChart -Gradient -Title 'Total traitement' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette4
                     New-ChartPie -Name 'Objets totale' -Value $runspaces.count
                     New-ChartPie -Name 'Objets trouvés' -Value $objettrouve                                   
                }
            }
    New-HTMLPanel -Width "60%" {
        New-HTMLChart -Gradient -Title 'Types extension' -TitleAlignment center -Height 200   { 
        New-ChartTheme  -Mode light
        $groupedFiles.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }                    
                }
            


    }
    }
    New-HTMLPanel  {
    New-HTMLChart -Title 'Top 5 Files' -TitleAlignment center {
                    New-ChartToolbar -Download pan 
                    New-ChartBarOptions -Gradient 
                    New-ChartLegend -Name $top5Groups.name[0],$top5Groups.name[1],$top5Groups.name[2],$top5Groups.name[2],$top5Groups.name[2] -HideLegend 
                    New-ChartBar -Name 'Patch' -Value $top5Groups.GetEnumerator().count[0],$top5Groups.GetEnumerator().count[1],$top5Groups.GetEnumerator().count[2],$top5Groups.GetEnumerator().count[2],$top5Groups.GetEnumerator().count[2]
                }
            }  
    New-HTMLPanel  {
                    New-HTMLChart -Title 'Top 5 Word' -TitleAlignment center  {
                    New-ChartToolbar -Download pan 
                    New-ChartBarOptions -Gradient  
                    New-ChartLegend -Name $top5Words.name[0],$top5Words.name[1],$top5Words.name[2],$top5Words.name[3],$top5Words.name[4] -HideLegend 
                    New-ChartBar -Name 'Word' -Value $top5Words.GetEnumerator().count[0],$top5Words.GetEnumerator().count[1],$top5Words.GetEnumerator().count[2],$top5Words.GetEnumerator().count[3],$top5Words.GetEnumerator().count[4]
                }
            }
    }
    }
    New-HTMLTab -Name 'Resume' -IconSolid user-alt   {     
    New-HTMLSection -Width "60%" -HeaderBackGroundColor Teal -name 'Groups Without members'  {          
    New-HTMLPanel -Width "40%" {
    New-HTMLGage -Label 'Indicator Risk' -MinValue 0 -MaxValue 100 -Value $totalRisk -ValueColor Black -LabelColor Black -Pointer -StrokeColor Akaroa -SectorColors AirForceBlue 
    }
    New-HTMLPanel  {         
          New-HTMLTabPanel -Orientation vertical -Theme 'pills' {
                    New-HTMLTab -Name 'Why check Sysvol 2.1' -IconBrands 500px {
                        New-HTMLText -FontSize 20px -Text "The Sysvol folder is crucial for distributing scripts and Group Policy Objects (GPOs) to all domain computers. 
                        It may contain sensitive information, such as plain-text passwords, making it a prime target for attackers. <br>A vulnerability in Sysvol can compromise the entire domain. Therefore, it is essential to restrict permissions, monitor changes, and regularly audit its contents to ensure network security and compliance."
                    }
                    New-HTMLTab -Name 'Audit GPO 2.2' -IconBrands 500px {
                        New-HTMLText -FontSize 20px -Text "Regularly audit GPOs to verify their contents, such as plain-text passwords in configuration files or auto-logon scripts, and the presence of unsigned sources. <br>Frequently run the GPOZaurr tool, which provides a comprehensive report to help identify and mitigate these risks.
                        <br>[GPOZaurr](https://github.com/EvotecIT/GPOZaurr/)<br>"
                    }
                    New-HTMLTab -Name 'Best Pratic 2.3' -IconBrands 500px {
                        New-HTMLText -FontSize 20px -Text "Enable audits on the Sysvol folder and monitor logs for multiple search attempts, as this may indicate enumeration attempts. Some elements in the Sysvol folder are not meant to be accessed by everyone. If possible, place a honeypot script in the Netlogon folder to trigger alerts for suspicious activity.
                        <br>[Autologon](https://learn.microsoft.com/fr-fr/sysinternals/downloads/autologon/)<br>"
                    }
                    New-HTMLTab -Name 'Tips 2.4' -IconBrands 500px {
                        New-HTMLText -FontSize 20px -Text "Do not store large files, such as ISO or .zip files, in the Sysvol folder. This can lead to replication issues and unnecessary consumption of storage resources, impacting the performance and reliability of your network<br> Move your scripts to a shared folder and grant access only to the relevant groups, not authenticated users. This will reduce vulnerabilities, especially if the scripts contain credentials or deploy critical applications."
                    }
                    New-HTMLTab -Name 'Hardening AD 2.4' -IconBrands 500px {
                        New-HTMLText -Color Green -FontSize 10px -Text "Use AD hardening to ensure security and reduce risks. Disable old protocols like SMB1 and anonymous enumeration on DC shares. Implement an N-tier architecture model, a PAW, and Silos. 
                        To facilitate this, refer to the HardenAD project.
                        <br>[HardenAD](https://github.com/LoicVeirman/HardenAD*/)<br>"
                    }
                }
    }
   }
             }  
     
}

#return $Results
}

# Appeler la fonction pour trouver les mots de passe en texte clair dans les fichiers Sysvol
#Invoke-SysvolAudit
