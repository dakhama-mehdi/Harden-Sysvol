
function Find-SMSClearTextPassword {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$dnsDomain = $env:USERDNSDOMAIN
    )

# Suppression de runspace "fantomes"
$PreviousRS = get-runspace | where-object {($_.id -ne 1)} 
if ($PreviousRS) { $PreviousRS.dispose() }

$Results =  $null


 #region getxmlcontenant
    # Récupérer les extensions de fichiers à partir du fichier XML
    $xmlFileExtensions = "C:\temp\file_extensions.xml"
    $extensionsXML = [xml](Get-Content $xmlFileExtensions)
    $fileExtensions = $extensionsXML.FileExtensions.Extension

    # Récupérer les motifs de mots de passe à partir du fichier XML
    $xmlPasswordPatterns = "C:\temp\password_patterns.xml"
    $passwordPatternsXML = [xml](Get-Content $xmlPasswordPatterns)
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

# définition des commandes à executer dans les jobs
$scriptblock = {
    Param (
        [string]$sysfiles,
        [string[]]$passwordPatterns
    )

    $content = Get-Content -Path $sysfiles -Raw;
    $results = @()  # Initialisation de la variable $results dans le scriptblock

    foreach ($pattern in $passwordPatterns) {
        if ($content -match $pattern) {
            $result = [PSCustomObject]@{
                FilePath = $sysfiles
                #Pattern = $pattern
                Word    = $Matches[0]
            }
            $results += $result  # Ajoute chaque résultat à la variable $results
        }
    }

    $results  # Retourne tous les résultats à la fin du scriptblock
}


# creation des jobs par machine et lancement des jobs
#foreach ($sysfiles in $sysvolFolders)
Get-ChildItem -Path "\\$dnsDomain\sysvol" -Recurse -File -Include $fileExtensions -Force -ErrorAction SilentlyContinue  | ForEach-Object { 

$sysfiles = $_.fullname

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
cls
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)  -and ($_.runspaceAvailability -like "Available")}
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $true)}
$slt_encours = get-runspace | where-object {($_.id -ne 1) -and ( $_.runspaceisremote -eq $true)}
$slt_tot = get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)}

write-host "Nbre Objets total= " $runspaces.count
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

$Results | Out-GridHtml

# Fermeture les connexions et suppression des slots du Pool
$pool.Close()
$pool.Dispose()


}

# Appeler la fonction pour trouver les mots de passe en texte clair dans les fichiers Sysvol
Find-SMSClearTextPassword