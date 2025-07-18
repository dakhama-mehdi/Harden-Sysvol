# FileHandlers.psm1

# supplementary function
function Test-ExcelMacros {
    param (
        [string]$FilePath
    )

    $result = $false # Macros check
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false

    try {
        $workbook = $excel.Workbooks.Open($FilePath, $false, $true)

        if ($workbook.HasVBProject) {
            $result = $true
        }
    } catch {
        $result = $_
    } finally {
        if ($workbook) { $workbook.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    return $result
}
function Check-SysvolAcl {
    param (
        [string]$FilePath,
        [string[]]$TrustGroups,
        [string[]]$PermissionTrust
    )

    $results = @()
    try {
        $acl = Get-Acl -Path $FilePath #-ErrorAction Stop
    } catch {
        #Write-Warning "Impossible de lire les ACL de : $fichier"
        return
    }

    $acl.Access | Where-Object {
    $_.IsInherited -eq $false -and
    $Trustgroups -notcontains $_.IdentityReference.Value -and
    $_.FileSystemRights -notin $PermissionTrust
} | ForEach-Object {
        $msg = "$($_.IdentityReference.Value) has '$($_.FileSystemRights)'"
        $results += [pscustomobject]@{
            FilePath = $FilePath
            Pattern  = 'WrongACL'
            Reason   = $msg
        }
    
    }
    return $results
}


#region FonctionDetect
function Get-OthersContent {
    param (
        [string]$filepath,
        [string[]]$patterns
    )  
    $results = @()    
    try {
        $findmatch = Select-String -Path $filepath -Pattern $patterns -AllMatches
        foreach ($match in $findmatch) {

            $patternMatch = $match.pattern -replace '\\b',''
            [string]$word= $match.line.Replace($patternMatch,'')

            if (-not([string]::IsNullOrEmpty(($word -replace '^[\s=:"]+').Trim()))){

            switch ($patternMatch) {
            '(?:[0-9]{1,3}\.){3}[0-9]{1,3}' { $patternMatch = 'IPv4' }
            '[a-fA-F0-9]{32}' { $patternMatch = 'MD5' }
            '[a-fA-F0-9]{40}' { $patternMatch = 'SHA-1' }
            '[a-fA-F0-9]{64}' { $patternMatch = 'SHA-256' }
            '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}' { $patternMatch = 'UPN' }
            default {
            switch -Wildcard ($patternMatch) {
            'net *' { $patternMatch = 'Commande Net User' }
            default { $patternMatch = $patternMatch }
             }
             }             
             }

            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = $patternMatch
                Reason   = $match.Line
            }
            $results += $result
            }            
     }
    } catch {
        $result = [PSCustomObject]@{
            FilePath = $filepath
            pattern  = 'error'
            Reason     = $_.Exception.Message
        }
        $results = $result
    }    
    return $results
}
function Get-DocxContent {
    param (
        [string]$filepath,
        [string[]]$patterns
    )
    $results = @()
    Import-Module PSWriteOffice
    try {
    $document = Get-OfficeWord -FilePath $filepath -ReadOnly
       $n = 0
    foreach ($pattern in $patterns) {
        $pattern = $pattern -replace '\\b', ''
        [string]$findtext = $document.Find($pattern) | Select-Object text -First 2
        if ($findtext) {
        if ($n -le '1') {
        foreach ($text in $findtext) {
        $n++
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = $pattern
                Reason   = $text
            }
            $results += $result
            }
        } elseif ($n -gt '1') { 
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = $pattern
                Reason    = "at least " + $n + " characters found"
            }
            $results += $result
         break }
    }
    }
    Close-OfficeWord -Document $document
    } 
    catch {   $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = 'error'
                Reason     = $_
            }
            $results = $result
    }
    $results
    }
function Get-XlsxContent {
    param (
        [string]$filepath,
        [string[]]$patterns
    )
    $results = @()
    Import-Module PSWriteOffice
    try {
    $excel = Get-OfficeExcel -FilePath $filepath 
    $n = 0
    foreach ($pattern in $patterns) {
        $pattern = $pattern -replace '\\b', ''
        [string]$findtext = $excel.Search($pattern, [System.Globalization.CompareOptions]::IgnoreCase, $false) | Select-Object value -First 2
        if ($findtext) {                    
            if ($n -le '2') {
            foreach ($text in $findtext) {
            $n++
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = $pattern
                Reason    = $text
            }
            $results += $result
        }
        } else { 
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = $pattern
                Reason    = "at least " + $n + " characters found"
            }
            $results += $result
         break }
            }
    }
    $excel.Dispose()
    } 
    catch {
    $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = 'error'
                Reason     = $_
            }
            $results = $result
    }
    $results
}
function Get-DocContent {
    param (
        [string]$filepath,
        [string[]]$patterns,
        [string]$wordinstalled
    )
    $results = @()
    if ($wordinstalled -eq $false) {
        $result = [PSCustomObject]@{
            FilePath = $filepath
            pattern  = 'requires_check'
            Reason     = 'Word is not installed'
        }
        $results += $result
    } else {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        try {
            $document = $wordApp.Documents.Open($filepath, [ref]$null, [ref]$true)
            $n = 0
            foreach ($pattern in $patterns) {
                $pattern = $pattern -replace '\\b', ''
                $find = $document.Content.Find | Select-Object -First 1
                $find.Text = $pattern
                $find.Forward = $true
                $find.Wrap = 1  # wdFindContinue
                $find.Execute() | Out-Null
                if ($find.Found) {
                    if ($n -le '2') {
                        $n++
                        $result = [PSCustomObject]@{
                            FilePath = $filepath
                            pattern  = $pattern
                            Reason     = "at least " + $n + " characters found"
                        }
                        $results = $result
                    } else { break }
                }
            }
            $document.Close([ref]$false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
        } catch {
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = 'error'
                Reason     = $_.Exception.Message
            }
            $results = $result
        } finally {
            $wordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    $results
}
function Get-XlsContent {
    param (
        [string]$filePath,
        [string[]]$patterns,
        [string]$excelInstalled
    )
    $results = @()
    
    if ($excelInstalled -eq $false) {
        $result = [PSCustomObject]@{
            FilePath = $filePath
            pattern  = 'requires_check'
            Reason   = 'Excel is not installed'
        }
    return $result
    } else {
        
        $value = Test-ExcelMacros -FilePath $filePath 

        if ($value -eq $true ) {

        $result = [PSCustomObject]@{
                                FilePath = $filePath
                                pattern  = "Macros detected"
                                Reason   = "Files contenant macros, need manualy check"
                            }
        return $result
        }

        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        try {
            $workbook = $excelApp.Workbooks.Open($filePath, [ref]$null, [ref]$true)
            $n = 0
            foreach ($pattern in $patterns) {
                $pattern = $pattern -replace '\\b', ''
                foreach ($sheet in $workbook.Sheets) {
                    $cells = $sheet.Cells.Find($pattern)
                    if ($null -ne $cells) {
                        if ($n -le '2') {
                            $n++
                            $result = [PSCustomObject]@{
                                FilePath = $filePath
                                pattern  = $pattern
                                Reason   = "at least " + $n + " characters found"
                            }
                            $results = $result
                        } else { break }
                    }
                }
            }
            $workbook.Close([ref]$false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null

        } catch {
            $result = [PSCustomObject]@{
                FilePath = $filePath
                pattern  = 'error'
                Reason     = $_.Exception.Message
            }
            $results = $result
        } finally {
            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    $results
}
function Get-PPTContent {
    param (
        [string]$filepath,
        [string[]]$patterns,
        [string]$wordinstalled
    )
    $results = @()
    if ($wordinstalled -eq $false) {
        $result = [PSCustomObject]@{
            FilePath = $filepath
            pattern  = 'requires_check'
            Reason     = 'Office is not installed'
        }
        $results += $result
    } 
    else {      
        try {
            $MSPPT = New-Object -ComObject powerpoint.application
            $PRES = $MSPPT.Presentations.Open($filepath, $true, $true, $false)

            $n = 0

            foreach($Slide in $PRES.Slides) {
            foreach ($Shape in $Slide.Shapes) {

            if ($Shape.HasTextFrame -eq "-1") {
             $text = $Shape.TextFrame.TextRange.Text            

             foreach ($pattern in $patterns) {
             
             $pattern = $pattern -replace '\\b', ''

             if ($text -match $pattern) {
             
             if ($n -le '2') {        
                 
             $n++
             
             $result = [PSCustomObject]@{
                            FilePath = $filepath
                            pattern  = $pattern
                            Reason     = "at least " + $n + " characters found"
                        }
                        $results = $result
                        }  else { break }
            }                       

            }
            }
            }
            }

            $MSPPT.PresentationClose
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MSPPT) | Out-Null
        } catch {
            $result = [PSCustomObject]@{
                FilePath = $filepath
                pattern  = 'error'
                Reason     = $_.Exception.Message
            }
            $results = $result
        } finally {
            $MSPPT.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MSPPT) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    $results
}
function Get-OdsContent {
    param (
        [string]$filePath,
        [string[]]$patterns
    )
    $results = @()


    # Open ODS file
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)

    # Extract contenant 
    $contentXmlEntry = $zip.Entries | Where-Object { $_.FullName -eq "content.xml" }
    $reader = [System.IO.StreamReader]::new($contentXmlEntry.Open())
    $contentXml = $reader.ReadToEnd()
    $reader.Close()
    $zip.Dispose()

    # Analys file XML
    [xml]$xmlContent = $contentXml

    # Read text XML
    $textContent = $xmlContent.'document-content'.InnerText

    $n = 0
    foreach ($pattern in $patterns) {
        $pattern = $pattern -replace '\\b', ''
        if ($textContent -match $pattern) {
            if ($n -le '2') {
                $n++
                $result = [PSCustomObject]@{
                    FilePath = $filePath
                    Pattern  = $pattern
                    Reason     = "at least " + $n + " characters found"
                }
                $results = $result
            }
        }
    }
    $results
}
function Get-Pdfcontent {
        param (
        [string]$filepath,
        [string[]]$patterns
        )

            $results = @()
            Import-Module PSWritePDF
            try {
            $text = (Convert-PDFToText -FilePath $filepath) -join "`n"
            $n = 0
            foreach ($pattern in $passwordPatterns) {
                if ($text -match $pattern) {
                if ($n -le '2') {
                    $n++
                    $result = [PSCustomObject]@{
                        FilePath = $filepath
                        pattern  = $pattern
                        Reason     = "at least " + $n + " characters found"
                    }
                    $results = $result
            }   else { break }
                }
            }
            } 
            catch {
            $result = [PSCustomObject]@{
                        FilePath = $filepath
                        pattern  = 'error'
                        Reason     = $_.Exception.Message
                    }
                    $results = $result
            }
            $results
}
function Get-OthersxmlContent {
    param (
        [string]$filepath,
        [string[]]$patterns,
        [int]$MaxContextLength = 40  # Number of characters to display after the word
    )  

    $results = @()    
    $maxMatchesPerLine = 2  # Limit to 2 occurrences per pattern per line

    try {
        # Read the file line by line
        $lines = Get-Content -Path $filepath

        foreach ($line in $lines) {
            foreach ($pattern in $patterns) {
                # Find all matches for the current pattern (limited to 2)
                $valueMatches = [regex]::Matches($line, $pattern) | Select-Object -First $maxMatchesPerLine | ForEach-Object { $_.Value }

                # Identify the type of detected pattern
                $patternMatch = switch ($pattern) {
                    '\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b' { 'IPv4' }
                    '\b[a-fA-F0-9]{32}\b' { 'MD5' }
                    '\b[a-fA-F0-9]{40}\b' { 'SHA-1' }
                    '\b[a-fA-F0-9]{64}\b' { 'SHA-256' }
                    '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' { 'UPN' }
                    '\b[Pp]assword\s*[:=]\s*\S+' { 'Password' }
                    default { $pattern }
                }

                # Store each found match
                foreach ($valueMatch in $valueMatches) {
                    if (-not [string]::IsNullOrEmpty($valueMatch)) {
                        
                        # Find the exact position of the word in the line
                        $startIndex = $line.IndexOf($valueMatch)
                        $endIndex = [Math]::Min($startIndex + $MaxContextLength, $line.Length)

                        # Extract the found word + 40 characters after
                        $reason = $line.Substring($startIndex, $endIndex - $startIndex).Trim()

                        # Add the result
                        $results += [PSCustomObject]@{
                            FilePath = $filepath
                            Pattern  = $patternMatch
                            Reason   = $reason  # Found word + 40 characters after
                        }
                    }
                }
            }
        }
    }
    catch {
        $results += [PSCustomObject]@{
            FilePath = $filepath
            Pattern  = 'error'
            Reason   = $_.Exception.Message
        }
    }

    return $results
}
function Get-Xmlcontent {
        param (
        [string]$filepath,
        [string[]]$patterns
        )
            $results = @()
            $xmlContent = Get-Content -Path $filepath -Raw
            if ($xmlContent -match 'cpassword') {
                $xml = [xml]$xmlContent
                # Check for cpassword
                $cpasswords = @()
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
                if ($cpasswords) {
                $result = [PSCustomObject]@{
                    FilePath  = $filepath
                    Pattern  = 'cpassword'
                    Reason     = $cpasswords
                }
                $results += $result
                }
            } 
            elseif ($xmlContent -match 'DefaultUserName') {
               
                $xml = [xml]$xmlContent
                # Check for AutoLogon
                $userName = ""
                $password = ""
                foreach ($registry in $xml.RegistrySettings.Registry) {
                    if ($registry.Properties.name -eq "DefaultUserName") {
                        $userName = $registry.Properties.value
                    }
                    if ($registry.Properties.name -eq "DefaultPassword") {
                        $password = $registry.Properties.value
                    }
                }
                $result  = [PSCustomObject]@{
                    FilePath = $filepath
                    Pattern  = 'AutoLogon'
                    Reason     = "$userName : $password"
                }
                $results += $result
            } 
            else { $results= Get-OthersxmlContent -filepath $filepath -patterns $patterns }
            $results
}
function Get-Xlsmcontent {
   param ( 
            [string]$filepath
        )             
             $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Macros detected '
                    Reason   = 'Files contenant macros, need check'
                }
             
        return $result
        }
function Get-Executablescontent {
        
        param (
        [string]$filepath
        )

        $results = @()
            if ($filepath -notlike '*.jar') { 
            $signature = Get-AuthenticodeSignature -FilePath $filepath
            if ($signature.Status -ne 'Valid') {
                $result = [PSCustomObject]@{
                    FilePath = $filepath
                    pattern  = "NotSigned"
                    Reason     = 'File is Not Signed'
                }
                $results = $result
            }
            } else {
        $isSigned = $false
        try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($filepath)
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -like "META-INF/*.SF") {
                $isSigned = $true
                break
            }
        }
        $zip.Dispose()
    } catch {
        $result = [PSCustomObject]@{
                        FilePath = $filepath
                        pattern  = 'error'
                        Reason     = $_.Exception.Message
                    }
                    $results = $result
    }

    if (!$isSigned) {
         $result = [PSCustomObject]@{
                    FilePath = $filepath
                    pattern  = "NotSigned"
                    Reason     = "Jar file not signed"
                }
                $results = $result
    }
            }
        $results
}
function Get-Zipprotectedbypass {
        param (
        [string]$filepath,
        [string]$zipinstalled
        )
    $extension = [System.IO.Path]::GetExtension($filePath).TrimStart('.')
    if ($extension -eq "zip" ) {
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($filepath)

        foreach ($entry in $zip.Entries) {
            $stream = $entry.Open()
            [byte[]]$buffer = New-Object byte[] 10
            $stream.Read($buffer, 0, $buffer.Length) | Out-Null
            $stream.Close()
            break  
        }

        $zip.Dispose()
        return $null
    }
    catch {
         if ($_.Exception.Message -match "(Read|Block|Password|Encrypted)") {
                    
          $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Zip protected'
                    Reason   = 'File protected by password'
                }
                    
            return $result
        } else {
            Write-Host "Erreur inattendue : $_"
            $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Zip check error'
                    Reason   = 'Error to read zip : $_'
                }
           return $result
        }
    }
    } elseif ($zipinstalled) {
    
    $sevenZipPath = $zipinstalled + "7z.exe"
    $output = & "$sevenZipPath" t "$filepath" -pBadPasswordConf 2>&1
    if ($output -match "Wrong password") {
            $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Zip protected'
                    Reason   = 'File protected by password'
                }                    
    return $result
    } 
    }

}
function Get-Requiredcheckcontent {
        param ( 
            [string]$filepath
        )             
             $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'check required'
                    Reason   = 'Binary does not match'
                }
             
        return $result
        }
function Get-CertifContent {
    param (
        [string]$filePath
    )

    $results = @()
    try {
        # Try to load certificat with class .NET X509Certificate2
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($filePath)        
        if (!$cert.Thumbprint) { $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Certificate Empty'
                    Reason   = 'Certificate without Thumbprint'
                }
      $results = $result }
        if ($cert.PrivateKey) { $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Certificate private key'
                    Reason     = 'certificate with exportable private key'
                }
      $results += $result }
        $cert.Dispose()
    }
    catch {
      $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Protected Certificate'
                    Reason     = ($_.Exception.Message).split(":").split("`n")[1].Trim()
                }
      $results = $result       
    }
    return $results
}
function Get-P7bCertContent {
    param (
        [string]$filePath
    )
    $results = @()
    try {
        # Create certificate collection contenant P7B ou P7C
        $certCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
        $certCollection.Import($filePath)

        # browse all certificate present in P7B/P7C
    foreach ($cert in $certCollection) {         
       if (!$cert.Thumbprint) { 
                    $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Certificate Empty'
                    Reason     = 'Certificate without Thumbprint'
                }
       $results += $result
       }
       if ($cert.PrivateKey) { 
       $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Certificate private key'
                    Reason     = 'certificate with exportable private key'
                }
       $results += $result
      }
      $cert.Dispose()
      }
      }
    catch {
      $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Protected Certificate'
                    Reason     = $_.Exception.Message
      }
      $results = $result
    }
    return $results
}
function Get-HiddenFilesInImage {
    param (
        [string]$filePath
    )
        
        $fileInfo = Get-Item $filePath
        $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)  

        # Skip file more than 4 Mo
        if ($fileSizeMB -gt 4) {
        $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Large size'
                    Reason   = "File ignored: (size: $fileSizeMB MB)"
                }
      return $result
        }

    try {
        # read binary and convert file to Hexa
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $fileHex   = [BitConverter]::ToString($fileBytes) -replace '-'

        $magicNumbers = [ordered]@{
            "MSI" = "D0CF11E0A1B11AE1" # MSI or office
            "RAR" = "526172211A0700";  # RAR file
            "ZIP" = "504B0304";   # ZIP file
            "7z"  = "377ABCAF271C"; # 7z file
            "mimikatz" = "6D696D696B61747A"
            "printf"   = "7072696E7466"
            "invoke-" = "696E766F6B652D"
            ".Net.WebClient" = "2E4E65742E576562436C69656E74"
            "EXE" = "4D5A";       # EXE file (MZ header)
        }

        foreach ($key in $magicNumbers.Keys) {
            $magicNumber = $magicNumbers[$key]
            $currentIndex = 0

            if ($fileHex -match $magicNumber) {

            if ($key -eq "EXE") {
                while ($fileHex.IndexOf($magicNumber, $currentIndex) -ne -1) {
                    $startIndex = $fileHex.IndexOf($magicNumber, $currentIndex)
                    # Extract the part after the Magic Number (limited to 400 bytes after)
                    $remainingHex = $fileHex.Substring($startIndex + $magicNumber.Length, [Math]::Min(200 * 2, $fileHex.Length - ($startIndex + $magicNumber.Length)))

                    # Check for the presence of the special string "0000004000" in this range
                    if ($remainingHex -match "0000004000") {
                    $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Suspicious Image'
                    Reason   = "EXE file with '0000004000' string detected"
                }
                return $result
                    }

                    # Continue the search from the next index
                    $currentIndex = $startIndex + $magicNumber.Length
                }
            }
            elseif ($key -eq "ZIP") {

    # Find the index where the ZIP file starts
    $startIndex = $fileHex.IndexOf($magicNumber, $currentIndex) / 2  # Divide by 2 because each hexadecimal represents 2 characters

    # Extract the ZIP bytes starting from this index
    $zipBytes = $fileBytes[$startIndex..($fileBytes.Length - 1)]

    $tempDir = [System.IO.Path]::GetTempPath()
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $tempZipPath = [System.IO.Path]::Combine($tempDir, "$fileNameWithoutExtension.zip")
    [System.IO.File]::WriteAllBytes($tempZipPath, $zipBytes)

    # Read the ZIP file and list the files
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempZipPath)

    # List the files inside the ZIP
    $fileNames = $zip.Entries | Select-Object -ExpandProperty FullName
    $fileNames | ForEach-Object { Write-Host "File in ZIP: $_" }

    $zip.Dispose()

    # Delete the temporary file
    Remove-Item $tempZipPath

    # Return the result
    $result = [PSCustomObject]@{
        FilePath =  $filePath
        Pattern  = 'Suspicious Image'
        Reason     = "ZIP detected in pictures. Containing: $($fileNames -join ', ')"
    }
    return $result
}
            else {
                $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Suspicious Image'
                    Reason     = "File $key detected in the image"
                }
                return $result
                }
            }
        }
    }
    catch {
        $result = [PSCustomObject]@{
                 FilePath =  $filepath
                 pattern  = 'Error'
                 Reason     = "Details : $_"
                }
       return $result
    }
}
function Get-HiddenFilesSpecificInImage {
    param (
        [string]$filePath
    )

    $fileInfo = Get-Item $filePath
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)  

    # Skip file more than 4 Mo
    if ($fileSizeMB -gt 4) {
        $result = [PSCustomObject]@{
                    FilePath =  $filePath
                    pattern  = 'Large size'
                    Reason   = "File ignored: (size: $fileSizeMB MB)"
                }
        return $result
    }

    function Extractzip-fromfile {
    # Extract ZIP and list files inside it
                    $startIndex = $fileHex.IndexOf($magicNumber, $currentIndex) / 2  # Divide by 2 for hex chars
                    $zipBytes = $fileBytes[$startIndex..($fileBytes.Length - 1)]

                    $tempDir = [System.IO.Path]::GetTempPath()
                    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
                    $tempZipPath = [System.IO.Path]::Combine($tempDir, "$fileNameWithoutExtension.zip")
                    [System.IO.File]::WriteAllBytes($tempZipPath, $zipBytes)

                    # Read the ZIP file and list the files inside it
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempZipPath)
                    $fileNames = $zip.Entries | Select-Object -ExpandProperty FullName
                    $fileNames | ForEach-Object { Write-Host "File in ZIP: $_" }

                    $zip.Dispose()
                    Remove-Item $tempZipPath

                    $result = [PSCustomObject]@{
                        FilePath =  $filePath
                        Pattern  = 'Suspicious Image'
                        Reason   = "ZIP detected in pictures. Containing: $($fileNames -join ', ')"
                    }
                    return $result
    }

    try {
        # Read binary and convert file to Hex
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $fileHex = [BitConverter]::ToString($fileBytes) -replace '-'
        $extension = [System.IO.Path]::GetExtension($filePath).TrimStart('.').ToLower()

        # Magic numbers for different file types
        $magicNumbers = [ordered]@{
            "MSI" = "D0CF11E0A1B11AE1" 
            "RAR" = "526172211A0700"  
            "ZIP" = "504B0304"
            "7z"  = "377ABCAF271C"
            "EXE" = "4D5A"
        }
       
        switch ($extension) {
        {$_ -in "jpg","jpeg"}  { $endoffile = "FFD9"}
        'png'   { $endoffile = "0000000049454E44AE426082"}
        'gif'   { $endoffile = "3B"}
        }
        foreach ($key in $magicNumbers.Keys) {
            $magicNumber = $magicNumbers[$key]
            $currentIndex = 0

            $truecheck = $endoffile + $magicNumber

            if (($fileHex -match $truecheck) -and ($fileHex.Substring($fileHex.Length -$endoffile.Length, $endoffile.Length) -ne $endoffile)) {

            if ($key -eq 'zip') { 
            return Extractzip-fromfile
            } else {
            $result = [PSCustomObject]@{
                                FilePath =  $filePath
                                pattern  = 'Suspicious Image'
                                Reason   = "$key file found in image with unexpected binary ending"
                            }
            return $result
            }
            }
            }        
        foreach ($key in $magicNumbers.Keys) {
            $magicNumber = $magicNumbers[$key]
            $currentIndex = 0

            if ($fileHex -match $magicNumber) {
                if ($key -eq "EXE") {
                    while ($fileHex.IndexOf($magicNumber, $currentIndex) -ne -1) {
                        $startIndex = $fileHex.IndexOf($magicNumber, $currentIndex)
                        # Extract the part after the Magic Number (up to 200 bytes)
                        $remainingHex = $fileHex.Substring($startIndex + $magicNumber.Length, [Math]::Min(200 * 2, $fileHex.Length - ($startIndex + $magicNumber.Length)))

                        # Check if the string "0000004000" appears in the remaining data
             
                        if ($remainingHex -match "0000004000") {
                            $result = [PSCustomObject]@{
                                FilePath =  $filePath
                                pattern  = 'Suspicious Image'
                                Reason   = "EXE file with '0000004000' string detected"
                            }
                            return $result
                        }

                        # Continue searching from the next index
                        $currentIndex = $startIndex + $magicNumber.Length
                    }

                }
                elseif ($key -eq "ZIP") {
                return Extractzip-fromfile
                }
                else {
                    $result = [PSCustomObject]@{
                        FilePath =  $filePath
                        pattern  = 'Suspicious Image'
                        Reason   = "File $key detected in the image"
                    }
                    return $result
                }
            }
        }
        if ($fileHex.Substring($fileHex.Length -$endoffile.Length, $endoffile.Length) -ne $endoffile) {

       $Keywords = [ordered]@{
            "EXE file"  = "4558452066696C65" 
            "printf"    = "7072696E7466"
            "fopen"     = "666F70656E"
            "malloc"    = "6D616C6C6F63"
            "strcpy"    = "737472637079"
            "system"    = "73797374656D"
            "socket"    = "736F636B6574"
            "class"     = "636C617373"
        }

        # Search other motifs in the last 2000 characters
        $searchHex = $fileHex.Substring([Math]::Max(0, $fileHex.Length - 2000))
        foreach ($word in $Keywords.Keys) {
        if ($searchHex -match $Keywords[$word]) {
        
                  $result = [PSCustomObject]@{
                        FilePath =  $filePath
                        pattern  = 'Suspicious Image'
                        Reason   = "Image probably modified, found argument : $word"
                    }
        return $result

        }
        }
        $result = [PSCustomObject]@{
                        FilePath =  $filePath
                        pattern  = 'Suspicious Image'
                        Reason   = "Image probably modified, not correctyl end binary match"
                    }
                    return $result
        }        
        }
    catch {
        $result = [PSCustomObject]@{
                 FilePath =  $filePath
                 pattern  = 'Error'
                 Reason   = "Details: $_"
                }
        return $result
    }
}
function Get-checkfilesize {
    param (
        [string]$filePath
    )

    $fileInfo = Get-Item $filePath
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2) 

    $result = [PSCustomObject]@{
                    FilePath =  $filepath
                    pattern  = 'Large size'
                    Reason   = "Size is so much, file ignored: (size: $fileSizeMB MB)"
                }             
    return $result 
}
#endregion FonctionDetect

# Add other functions for different file types...
function Get-CompressedFileType {    
    param (
        [string]$filePath,
        [object[]]$detectedType
    )


    # More check fot DOCX, XLSX, ODT, ODS ou JAR PPTX zip
    if ($detectedType -contains "docx" -or $detectedType -contains "jar") {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)

        $fileNames = $zip.Entries | Select-Object -ExpandProperty FullName

        if ($fileNames -contains "word/document.xml") {
            $detectedType = "docx"
        }
        elseif ($fileNames -contains "xl/vbaProject.bin") {
          $extension = [System.IO.Path]::GetExtension($filePath).TrimStart('.') 
            if ($extension -in "xlsm","xlam") {
            $detectedType = $extension 
            }
        }
        elseif ($fileNames -contains "ppt/presentation.xml") {
            $detectedType = "pptx"
        }
        elseif ($fileNames -contains "visio/document.xml") {
            $detectedType = "vsdx"
        }
        elseif ($fileNames -contains "xl/workbook.xml") {
            $detectedType = "xlsx"
        } 
        elseif ($fileNames -contains "content.xml") {
            $mimetypeEntry = $zip.Entries | Where-Object { $_.FullName -eq "mimetype" }
            if ($mimetypeEntry -ne $null) {
                $reader = [System.IO.StreamReader]::new($mimetypeEntry.Open())
                $mimetype = $reader.ReadToEnd()
                $reader.Close()

                switch ($mimetype) {
                    "application/vnd.oasis.opendocument.spreadsheet" {
                        $detectedType = "ods"
                    }
                    "application/vnd.oasis.opendocument.text" {
                        $detectedType = "odt"
                    }
                    "application/vnd.oasis.opendocument.presentation" {
                        $detectedType = "odp"
                    }
                    "application/vnd.oasis.opendocument.text-template" {
                        $detectedType = "ott"
                    }
                }
            }
        } 
        elseif ($fileNames -contains "META-INF/MANIFEST.MF") {
            $detectedType = "jar"
        } 
        else   { $detectedType = "others" }
        $zip.Dispose()
     } 
    catch {
        $detectedType = "requires_check"
         }
    }
    elseif ($detectedType -contains "doc") {
        
        $a = [System.IO.File]::ReadAllBytes($filePath)
        $content = [System.Text.Encoding]::ASCII.GetString($a)

        if ($content.Contains("Word.Document")) {
            $detectedType = "doc"
        } elseif ($content.Contains("MSI") -or $content.Contains("Installer")) {
            $detectedType = "msi"
        } elseif ($content.Contains("Excel") ) {
            $detectedType = "xls"
        } elseif ($content.Contains("PowerPoint")) {
            $detectedType = "ppt"
        } elseif ($content.Contains("Microsoft Visio")) {
            $detectedType = "vsd"
        } else {
            #check for db files
            $offsetBytes = [System.IO.File]::ReadAllBytes($filePath)[1024..1050]
            $offsetAscii = [System.Text.Encoding]::ASCII.GetString($offsetBytes).Trim()
            $filteredAscii = ($offsetAscii -split '').Where{ $_ -match '[A-Za-z ]' } -join ''

            if ($filteredAscii -replace '\s+', ' ' -eq "Root Entry") {
            $detectedType = "db"
            } 
            else {
            $detectedType = "others"
            }
        }
     }
    return $detectedType
}
function Get-FileType {
    param (
        [string]$filePath,
        [int]$Maxfilesize,
        [int]$MaxBinarysize,
        [object]$jsonContent
    )

    $detectedType = "others"
    $extension = [System.IO.Path]::GetExtension($filePath).TrimStart('.') 
    $fileHeaderHex = [System.IO.File]::ReadAllBytes($filePath)[0..20] | ForEach-Object { "{0:X2}" -f $_ }
    $fileHeaderHex = ($fileHeaderHex -join '').Trim()
    
    if ($fileHeaderHex.Length -eq 0) {
        return "empty"
    }

    $matchFound = $false
    foreach ($entry in $jsonContent.magic_numbers) {
        if ($matchFound) { break }

        foreach ($expectedMagic in $entry.magic) {
            if ($matchFound) { break }

            if ($expectedMagic.Length -le $fileHeaderHex.Length) {
                $difference = Compare-Object -ReferenceObject $expectedMagic -DifferenceObject ($fileHeaderHex.Substring(0, $expectedMagic.Length)) -SyncWindow 0

                if ($difference.Count -eq 0) {
                    if ($entry.offset) {
                    foreach ($offsetItem in $entry.offset) {
                        $offsetPosition = $expectedOffsetValue = $null              
                        $offsetPosition = $offsetItem[0]
                        $expectedOffsetValue = $offsetItem[1]                                                
                        $lastpostition = $offsetPosition + ($expectedOffsetValue.Length/2)-1

                        $CustomfileHeaderHex = [System.IO.File]::ReadAllBytes($filePath)[$offsetPosition..$lastpostition] | ForEach-Object { "{0:X2}" -f $_ }
                        $specificBytes = ($CustomfileHeaderHex -join '').Trim()               

                        if ($specificBytes -eq $expectedOffsetValue) {
                            if ($extension -notin $entry.extensions) {
                                $detectedType = 'requires_check'
                            } else {
                                $detectedType = $extension
                            }
                            $matchFound = $true
                            break
                        }
                       }                       
                    }
                    else {
                        if ($extension -notin $entry.extensions) {
                            $detectedType = 'requires_check'
                        } 
                        else {
                        if ($entry.extensions -contains "docx" -or $entry.extensions -contains "doc") {
                            $detectedType = Get-CompressedFileType -filePath $filePath -detectedType $entry.extensions
                        }
                        if ($detectedType -eq "others" -and $extension -in $entry.extensions) {
                         $detectedType = $extension
                         } elseif ($detectedType -ne $extension -and $detectedType -ne "others") {
                         $detectedType = 'requires_check'
                         }
                    }
                        $matchFound = $true
                        break
                    }
                }
            }
        }
    }

    if ($detectedType -eq 'others' -and $extension -in $jsonContent.magic_numbers.extensions) {
        $detectedType = 'requires_check'
    }

    if ($detectedType -eq "others") {    
        $firstLine = Get-Content -Path $filePath -TotalCount 1
        if ($firstLine -match '^<\?xml') {
            $detectedType = "xml"
        }
    }

    # Check the file size to avoid slowing down analyze files larger than 10 MB, along with installation file extensions over 50 MB
    if ($detectedType -ne 'requires_check' ) {
    $fileInfo = Get-Item $filePath
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    switch ($extension)  {
        {$_ -in "exe","msi","dll","msu","cab","zip","rar"} {
        if ($fileSizeMB -gt $MaxBinarysize) { return "bigsize" }
        }
        default {  
        if ($fileSizeMB -gt $Maxfilesize) { return "bigsize" }
        }
        }
    }

    return $detectedType
}
# SIG # Begin signature block
# MIImVgYJKoZIhvcNAQcCoIImRzCCJkMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0BUD7Ms2+2HSV
# czwDqzY9TwbA8imbLQ+7nWtFNu54N6CCH+wwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAEelmrgbNiaNbFGy5l
# xSEowwnpceE099QEP1rRds4QAzANBgkqhkiG9w0BAQEFAASCAYADTfVGpixjezLP
# JUPdmPLP50z0W1tMviidB/9nd+W69b71XpvkadUHrHpbf9f9jwTI02Eh3Oq6oGE6
# FdojCHgijl+Yovvr85duquk42yshjS06tMyv7SZjaLllKFZnzJJU0IlJLd/z8zWg
# Duda1cz0k217UjHqXprxVpaoBQpvS41VJ0tnsK74GK0hUe8/7hZsT2W8TQ7x2jrx
# +6ZmojMJss9Q0OBRfocS1DSlUfKbm6mSFMt0R+/IAXcNeaYv5NhSfeCJHsfcefRH
# vnd641OlFyYboy42ZDeTYVU+QWXMpko9RfW8FPJivzql8gdx4pzqqlZUPhrgyqI7
# 6Q6WxlD2vctsyAsnnt1/aG73TxYqfVh11rLz6qBr5L61AUG7ikSBJuc4YRz5yeDe
# D8N0Pe9N0YkWEST3CExMcibEQsafyfwZmEVl6tcv8oGyDw2zFDqqG+GfswlgVoDY
# I9EJ/ttRlcMPx8CvoDjwB0sTaM/KdouQ++vapJbROR0ppHU2SUihggMgMIIDHAYJ
# KoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNB
# NDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhALrma8Wrp/lYfG+ekE4zMEMA0G
# CWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjUwNjMwMTczMzA2WjAvBgkqhkiG9w0BCQQxIgQgnlJWOpf6yI7/
# HIUbBSNq9EFd/xUQVeS9VajrwBKQlKwwDQYJKoZIhvcNAQEBBQAEggIAdPOa+oNl
# QNFbkL/WnBHX4jt6DuvUl79SIbDFpZ4l+B9xTm2aHYOFLxu5RB2yn6AaBgUAEkIo
# tKy6Qv13bXL7eoIq0wLi2WAmj4ndcMh4t7TVN14yVVI8t/0dm9UN3qsVHYocdI/3
# peOwA8GfGi+7Ki+AqQ6q6FG/nQ4qbdc7VNr2oeRMS/i6Z4ADQBEj8V/I4Mirn6YM
# VIf9LP/Ttg8XrhZUJAOi2qy0lhB/68NTWkiLFA/dW3EpMe0bjYB9XfnzaFxjFZ1Z
# yETVIDDhyLaakJiLqxFOAS3lcLLM08nZzQPsJP7i7kG7ClqOUckRJ0jE3EWd2fQ2
# j+8+G78dOLuPiSTYFS02xn3L/1el8SdoGntDy0xKAnTQQYzWx1+IyhQ7dSq0qGlg
# Zad/0NhMN51sEhPUWQrkRE8HTW1iF5Xgp3gj4I902iJHYD3zcUKZGDH1H4BXzUbF
# JHxmAtDmrUUhz9NnemX5BjK33RT2oY/02ie+9ZFK14DQuuzCpoSCJwFiWJ3W6dEC
# n3fzoTMU2YK7k/456WHJPpzG2X/baSGTC8oXl0Xf3N2vxavWgnlPBKxH1C7UJYi5
# UBKZbIy40aIQz1KgK5tVoBhbB5hoAtnN4kbINXBd0574JNEAfekAdUoG17j7Z6QT
# diTh040er7R+K9KxvsPmsIhWQ1Cwbo5IOhc=
# SIG # End signature block
