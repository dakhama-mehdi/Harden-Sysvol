#
# Manifest module for 'Harden Sysvol'
#
# Generated by�: Mehdi Dakhama
#
# Generated�: 11/2024

@{
    # ID du module
    RootModule        = 'HardenSysvol.psm1'

    # Infos du module
    ModuleVersion     = '1.7.0'
    GUID              = '31c02f99-d8aa-4996-8c55-647041bd7bbf'
    Author            = 'DAKHAMA Mehdi'
    CompanyName       = 'HardenAD'
    Copyright         = '(c) 2024 DAKHAMA Mehdi. All rights reserved.'
    Description       = 'Harden Sysvol is a Powershell Module to scan sysvol folder to search the sensitivity data, and vulnerability.'
    PowerShellVersion = '5.1'
    
    # D�pendances
    RequiredModules = @(@{
            ModuleVersion = '0.0.180'
            ModuleName    = 'PSWriteHTML'
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
        },  @{
            ModuleVersion = '0.2.0'
            ModuleName    = 'PSWriteOffice'
            Guid          = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
        })

    FunctionsToExport = @('Invoke-Hardensysvol')

    CmdletsToExport = @()

    VariablesToExport = '*'

    AliasesToExport = @()

    PrivateData = @{

    PSData = @{

        Tags = @('HardenSysvol','AuditAD','AuditNetlogon','AuditSysvol','ActiveDirectory','Security','Vulnerability')

        LicenseUri = 'https://github.com/dakhama-mehdi/Harden-Sysvol/blob/main/LICENSE'

        ProjectUri = 'https://github.com/dakhama-mehdi/Harden-Sysvol'

        IconUri = "https://github.com/dakhama-mehdi/Harden-Sysvol/blob/main/Pictures/HardenSysvol.png?raw=true"
			
		ReleaseNotes  = @'
        1.7.0 : Add check 7Z,Rar protected by password if 7zip is installed
        1.6.8 : Revise all magic number and add more total about 180 magic number
        1.6.7 : Add check db files, add check signing msu,cab, add check keepass files
                Add check zip protected by password
        1.6.6 : Fix score
        1.5.0 : Support certificat cer,der,pam,p7b 
                Add Json file - Extensions binary
		1.4.0 : Add new options : Add/Remove Pattern - Add/Remove Extension
			    Check ZIP/RAR
        1.3.0 : Support PPTX, ODP, Detect MD5,SHA1,SHa256, IPv4
        1.2.0 : Version Beta (Fix double result)
        1.1.0 : Fix dependence model error
        1.0.0 : Version Test
'@
    } 
} 
    HelpInfoURI = 'https://github.com/dakhama-mehdi/Harden-Sysvol/tree/main/Docs'

}

# SIG # Begin signature block
# MIImVgYJKoZIhvcNAQcCoIImRzCCJkMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBV3b5VJi6b9l74
# uY5TzhcgU0GHYF2vRuxN+hS8hEuJc6CCH+wwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCBRfXkBpqsgJAJPOJd
# IZ0H3jtcH3a43iue7OAGG1dnYTANBgkqhkiG9w0BAQEFAASCAYAR4NvK+saZhr+3
# TfhSdWMlnx+ISifrNoJYq0kPPOR+dCm4lD1B4Jkvq/iUWDM1R3Oh34J7nbJSXvEJ
# Lm/NGyeOm8T/PYOP9Vwe8k4o8EOkuH1IF8r5Q/QeuGq3kEfDl7nEYaHngTNKic7b
# fheiokn6oAp3OP6zS0HlkmO0xo8KmKP4W1qOYqX8sVXjIZ6WF6m5TdLI/vOsh0kt
# JE1OSuVskkK/sR2FaOtpFokrlR0YFPBjg0voMk9/E/81hAil5HCfRpi6xNAxglNO
# pCXLUcrjxgda9L3a9fmAbB5DkjEKvPB4RP/nIPoTpg/ZH8ILGH7OdpaRQfFOrOBD
# xH6Inr7psQ68HNqCtNKElzhHC5+0jE9CWvV64ur6Ucd+ghyjajGY/13iCwNbJFVY
# SgNFOcL/fMp4VSzTaGheNKM+FR5O+oFXT7ef1DQJAttF19HgCt/TZdbL9sosUfuh
# ycVqvtIHib52u1af5djLcEpBm966W4Ti5jJUqZdijrD2vOoUFUmhggMgMIIDHAYJ
# KoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNB
# NDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhALrma8Wrp/lYfG+ekE4zMEMA0G
# CWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjQxMTEzMjMzOTE3WjAvBgkqhkiG9w0BCQQxIgQgScTpc5k+Ym5l
# 6FFCMDXYaKba1M9TlNQ39tEFakX4eXQwDQYJKoZIhvcNAQEBBQAEggIAscfssbsg
# wLoAWmpwYpDUm4Kxyek76QGHC5lrWg2Rz/kiI7tZs8dOpiDWupjm1ZqLPbL8RcF8
# 0EvO6M1LRL/v2ooiuxnYRk42SNitoJbdPOpepnSZJd2Pe0OKMsX55UQ0maIKSXPU
# biJuBBaXZj3F3EF8Wqw38ANoQlglskNmL8C6QMunE1W2KilPA2o6Q+h3qxJKPDPy
# cvNBstClrbHfZPPN7d5mdynSaJV/oPs3IjQuDUTA2yvqZWDij5ohtKzVue9uY2G8
# ypdDno4tte5dXPX/0MJDwKdty+sh3neLv0QrZIHgM4SViChLRWjJP426P8rLKitv
# 88xY9dTP05zvulJ82eSj4pS5wdqPJyX6NkQcvWpFGhSsHD6wCfrNkK5n4jtJJsRN
# 4f+V9GUD5Ck7KbBREGC8ddQd1+ecPwlqOwbXWqL4WwwBKALdmEev4a0S3DmMHrUh
# 85QwYRW6tsVGuiYPDkY7fFQg1aObj0DMKb3s+XtOGINkqB8i111444a+rbMyBCKv
# eIDLVHtAAYR9z3waCmG0Kf4E22M29QCNvNpV24keIAMXbFe9o/OAam7UKikUKZq1
# TchvuNsYAsoMua4+ppuSRHtPJ6vfQNv4lm4xPNuy3ETkXY/kCx7d2zuwSytgX6gy
# 1z/cvrHIe1zEiMXF9HfZgu9bqhvtUUA2ios=
# SIG # End signature block