# FileHandlers.psm1

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
                Reason     = $match.Line
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
        $results += $result
    } else {
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
                                Reason     = "at least " + $n + " characters found"
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
            else {
                $xml = [xml]$xmlContent
                foreach ($pattern in $patterns) {
                    $findmatches = $xmlContent | Select-String -Pattern $pattern
                    foreach ($match in $findmatches) {
             
             switch ($pattern) {
            '\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b' { $pattern = 'IPv4' }
            '\b[a-fA-F0-9]{32}\b' { $pattern = 'MD5' }
            '\b[a-fA-F0-9]{40}\b' { $pattern = 'SHA-1' }
            '\b[a-fA-F0-9]{64}\b' { $pattern = 'SHA-256' }
            default {
            switch -Wildcard ($pattern) {
            'net *' { $pattern = 'Commande Net User' }
            default { $pattern = $pattern }
             }
             }             
             }
             
             $result = [PSCustomObject]@{
                            FilePath = $filepath
                            Pattern  = $pattern
                            Reason     = $match.Matches.Value
                        }
                        $results += $result
                    }
                }
            }
            $results
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
                    Reason     = 'Binary does not match'
                }
             
        return $result
        }

function Get-CertifsContent {
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
                    Reason     = 'Certificate without Thumbprint'
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
                    Reason     = "File ignored: (size: $fileSizeMB MB)"
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
                    Reason     = "EXE file with '0000004000' string detected"
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
        elseif ($fileNames -contains "xl/workbook.xml") {
            $detectedType = "xlsx"
        } 
        elseif ($fileNames -contains "ppt/presentation.xml") {
            $detectedType = "pptx"
        }
        elseif ($fileNames -contains "visio/document.xml") {
            $detectedType = "vsdx"
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
        [object]$jsonContent
    )

    $detectedType = "others"
    $extension = [System.IO.Path]::GetExtension($filePath).TrimStart('.') 
    $fileHeaderHex = [System.IO.File]::ReadAllBytes($filePath)[0..16] | ForEach-Object { "{0:X2}" -f $_ }
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

    return $detectedType
}

