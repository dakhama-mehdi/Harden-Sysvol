
function Find-SMSClearTextPassword {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$dnsDomain = $env:USERDNSDOMAIN,
        
        [Parameter(Mandatory = $false)]
        [String[]]$ignoreExtensions  # Tableau pour stocker les extensions à ignorer
    )

# Tester l'acces au partage
$testpath = Test-Path "\\$dnsDomain\sysvol\"
if ($testpath -eq $false) {
throw "Cannot access domain or share, pls check with GCI $dnsDomain"
}

# Tester la présence de module
if (!(Get-Module -ListAvailable -Name "PSWriteHTML"))
{    
    Write-Host "ReportHTML Module is not present, attempting to install it" -ForegroundColor Red    
    Install-Module -Name PSWriteHTML,PSWriteOffice,PSWritePDF -Force -Scope CurrentUser
    Import-Module PSWriteHTML -ErrorAction SilentlyContinue
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
    $extensionsXML = [xml](Get-Content $xmlFileExtensions)
    $fileExtensions = $extensionsXML.FileExtensions.Extension

    # Récupérer les motifs de mots de passe à partir du fichier XML
    $xmlPasswordPatterns = "C:\temp\password_patterns.xml"
    $passwordPatternsXML = [xml](Get-Content $xmlPasswordPatterns -Encoding UTF8)
    $passwordPatterns = $passwordPatternsXML.PasswordPatterns.Pattern

    # Récupérer les fichiers dans les dossiers Sysvol du domaine DNS spécifié
    #$sysvolFolders = Get-ChildItem -Path "\\$dnsDomain\sysvol" -Recurse -File -Include $fileExtensions -ErrorAction SilentlyContinue

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
                        pattern  = ''
                        Word     = $pattern
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
                        pattern  = ''
                        Word     = $pattern
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
                        pattern  = ''
                        Word     = $pattern
                    }
                    $results += $result
                }
            }
      $results  # Retourne tous les résultats à la fin du scriptblock
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
            
            if ($n -eq 2) {

                      # Retourne tous les résultats à la fin du scriptblock
                    break;
            }
        }

        $results  # Retourne tous les résultats à la fin du scriptblock
        }
    }

    # Obtenir l'extension du fichier
    $extension = [System.IO.Path]::GetExtension($sysfiles).ToLower()

    # Exécuter la commande appropriée en fonction de l'extension du fichier
    if ($extension -like '*.doc*' -and $fileCommands.ContainsKey('.docx')) {
        & $fileCommands['.docx']
    }
    if ($fileCommands.ContainsKey($extension)) {
        & $fileCommands[$extension]
    } else {
        & $fileCommands['default']
    }
}


$fichiertraite = 0
# creation des jobs par machine et lancement des jobs

Get-ChildItem -Path \\$dnsDomain\sysvol -Recurse -File -Include $fileExtensions -Exclude $ignoreExtensions -Force -ErrorAction Continue | ForEach-Object {

$fichiertraite++
$sysfiles = $_.FullName

#clear
Write-Host nous traitons $sysfiles

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
#cls
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

# Regroupement : 

#region Top5
$sortedGroups = $Results.filepath | Group-Object | Sort-Object -Property Count -Descending

# Sélectionner les 5 premiers groupes
$top5Groups = $sortedGroups | Select-Object Count,Name -First 5

# Afficher les 5 premiers groupes
$commonPath = "\\ENI.LOCAL\sysvol\eni.local"
$top5Groups = $top5Groups | ForEach-Object {
    $_.Name = $_.Name -replace [regex]::Escape($commonPath), ''
    $_
}

# Top 5 word
$top5Words= $Results.pattern | Group-Object | Sort-Object -Property Count -Descending | Select-Object Count,name -First 5

#nombre d'objet max 
$objettrouve = $sortedGroups.Count

# Regrouper les chemins de fichier par extension de fichier
$groupedFiles = $sortedGroups.name | Group-Object -Property { ($_ -split "\.")[-1] } | select Count,name

# Fin du script : obtenir la date et l'heure actuelles
$endDate = Get-Date

# Calculer la différence de temps
$elapsedTime = New-TimeSpan -Start $startDate -End $endDate
$elapsedTime = $($elapsedTime.ToString("hh\:mm\:ss"))

#endregion Top5

# Fermeture les connexions et suppression des slots du Pool
$pool.Close()
$pool.Dispose()

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
    
    New-HTMLSection {
    New-HTMLTableOption -DataStore JavaScript 
    New-htmlTable -HideFooter -DataTable $Results -TextWhenNoData 'Information: No Groups were found'
        }

    New-HTMLSection {
    
    New-HTMLPanel {
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
                New-HTMLChart -Title 'Top 5 Files' -TitleAlignment center  {
                    New-ChartToolbar -Download pan
                    New-ChartBarOptions -Gradient  
                    New-ChartLegend -Name $top5Groups.name[0],$top5Groups.name[1],$top5Groups.name[2],$top5Groups.name[2],$top5Groups.name[2] -HideLegend 
                    New-ChartBar -Name 'Patch extensions' -Value $top5Groups.GetEnumerator().count[0],$top5Groups.GetEnumerator().count[1],$top5Groups.GetEnumerator().count[2],$top5Groups.GetEnumerator().count[2],$top5Groups.GetEnumerator().count[2]
                }
            }  


    New-HTMLPanel  {
                New-HTMLChart -Title 'Top 5 Word' -TitleAlignment center  {
                    New-ChartToolbar -Download pan
                    New-ChartBarOptions -Gradient  
                    New-ChartLegend -Name $top5Words.name[0],$top5Words.name[1],$top5Words.name[2],$top5Words.name[3],$top5Words.name[4] -HideLegend 
                    New-ChartBar -Name 'Patch extensions' -Value $top5Words.GetEnumerator().count[0],$top5Words.GetEnumerator().count[1],$top5Words.GetEnumerator().count[2],$top5Words.GetEnumerator().count[3],$top5Words.GetEnumerator().count[4]
                }
            }  


    }
     
}

return $Results

}

# Appeler la fonction pour trouver les mots de passe en texte clair dans les fichiers Sysvol
$var= Find-SMSClearTextPassword