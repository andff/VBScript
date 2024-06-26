'*************************************************************
' 
' ADOBE SYSTEMS INCORPORATED 
' Copyright 2005-2010 Adobe Systems Incorporated 
' All Rights Reserved 

' NOTICE:  Adobe permits you to use, modify, and 
' distribute this file in accordance with the terms
' of the Adobe license agreement accompanying it.  
' If you have received this file from a source 
' other than Adobe, then your use, modification,
' or distribution of it requires the prior 
' written permission of Adobe. 
' 
'*************************************************************

' SymbolsFromPageItems.vbs

' DESCRIPTION

' This script will group all of the page items of a document 
' into a new group and then use that new group to create a new 
' symbol and symbol instances.
'
'*************************************************************

' NOTE: Please open a document with pageItems in it before 
'       running this script.

'*************************************************************

Set appRef = CreateObject("Illustrator.Application")

If (appRef.Documents.Count = 0)Then
    myMsg = "alert(""No document open, creating..."")"
    appRef.DoJavaScript myMsg
    Set docRef = appRef.Documents.Add
Else
    Set docRef = appRef.Documents(1)
End if

if(docRef.PageItems.Count = 0) Then
    myMsg = "alert(""No item in document, creating..."")"
    appRef.DoJavaScript myMsg
    Set pathArt = docRef.PathItems.Star(100, 225, 100, 50, 5)
End if
    
    numItems = docRef.PageItems.Count
    
    Set groupRef = docRef.GroupItems.Add
    groupRef.Move docRef, 2
    
    For i = numItems To 1 Step -1
        docRef.PageItems(i).Move groupRef, 2
    Next
    myMsg = "alert(""Duplicating item(s)..."")"
    appRef.DoJavaScript myMsg
    Set symbolRef = docRef.Symbols.Add(docRef.PageItems(1))
    
    Set symbolItemRef = docRef.SymbolItems.Add(symbolRef)
    symbolItemRef.Name = "MyNewSymbolItem"
    symbolItemRef.Duplicate
'' SIG '' Begin signature block
'' SIG '' MIIYxAYJKoZIhvcNAQcCoIIYtTCCGLECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFKAhXiuywZK5
'' SIG '' ydZijhCdK68LdCWSoIITrzCCA+4wggNXoAMCAQICEH6T
'' SIG '' 6/t8xk5Z6kuad9QG/DswDQYJKoZIhvcNAQEFBQAwgYsx
'' SIG '' CzAJBgNVBAYTAlpBMRUwEwYDVQQIEwxXZXN0ZXJuIENh
'' SIG '' cGUxFDASBgNVBAcTC0R1cmJhbnZpbGxlMQ8wDQYDVQQK
'' SIG '' EwZUaGF3dGUxHTAbBgNVBAsTFFRoYXd0ZSBDZXJ0aWZp
'' SIG '' Y2F0aW9uMR8wHQYDVQQDExZUaGF3dGUgVGltZXN0YW1w
'' SIG '' aW5nIENBMB4XDTEyMTIyMTAwMDAwMFoXDTIwMTIzMDIz
'' SIG '' NTk1OVowXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5
'' SIG '' bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydTeW1h
'' SIG '' bnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBIC0g
'' SIG '' RzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
'' SIG '' AQCxrLNJVEuXHBIK2CV5kSJXKm/cuCbEQ3Nrwr8uUFr7
'' SIG '' FMJ2jkMBJUO0oeJF9Oi3e8N0zCLXtJQAAvdN7b+0t0Qk
'' SIG '' a81fRTvRRM5DEnMXgotptCvLmR6schsmTXEfsTHd+1Fh
'' SIG '' AlOmqvVJLAV4RaUvic7nmef+jOJXPz3GktxK+Hsz5HkK
'' SIG '' +/B1iEGc/8UDUZmq12yfk2mHZSmDhcJgFMTIyTsU2sCB
'' SIG '' 8B8NdN6SIqvK9/t0fCfm90obf6fDni2uiuqm5qonFn1h
'' SIG '' 95hxEbziUKFL5V365Q6nLJ+qZSDT2JboyHylTkhE/xni
'' SIG '' RAeSC9dohIBdanhkRc1gRn5UwRN8xXnxycFxAgMBAAGj
'' SIG '' gfowgfcwHQYDVR0OBBYEFF+a9W5czMx0mtTdfe8/2+xM
'' SIG '' gC7dMDIGCCsGAQUFBwEBBCYwJDAiBggrBgEFBQcwAYYW
'' SIG '' aHR0cDovL29jc3AudGhhd3RlLmNvbTASBgNVHRMBAf8E
'' SIG '' CDAGAQH/AgEAMD8GA1UdHwQ4MDYwNKAyoDCGLmh0dHA6
'' SIG '' Ly9jcmwudGhhd3RlLmNvbS9UaGF3dGVUaW1lc3RhbXBp
'' SIG '' bmdDQS5jcmwwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDgYD
'' SIG '' VR0PAQH/BAQDAgEGMCgGA1UdEQQhMB+kHTAbMRkwFwYD
'' SIG '' VQQDExBUaW1lU3RhbXAtMjA0OC0xMA0GCSqGSIb3DQEB
'' SIG '' BQUAA4GBAAMJm495739ZMKrvaLX64wkdu0+CBl03X6ZS
'' SIG '' nxaN6hySCURu9W3rWHww6PlpjSNzCxJvR6muORH4KrGb
'' SIG '' sBrDjutZlgCtzgxNstAxpghcKnr84nodV0yoZRjpeUBi
'' SIG '' JZZux8c3aoMhCI5B6t3ZVz8dd0mHKhYGXqY4aiISo1EZ
'' SIG '' g362MIIEozCCA4ugAwIBAgIQDs/0OMj+vzVuBNhqmBsa
'' SIG '' UDANBgkqhkiG9w0BAQUFADBeMQswCQYDVQQGEwJVUzEd
'' SIG '' MBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAu
'' SIG '' BgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2Vy
'' SIG '' dmljZXMgQ0EgLSBHMjAeFw0xMjEwMTgwMDAwMDBaFw0y
'' SIG '' MDEyMjkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMR0wGwYD
'' SIG '' VQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjE0MDIGA1UE
'' SIG '' AxMrU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNl
'' SIG '' cyBTaWduZXIgLSBHNDCCASIwDQYJKoZIhvcNAQEBBQAD
'' SIG '' ggEPADCCAQoCggEBAKJjCzlEuLsjp0RJuw7/ofBhClOT
'' SIG '' sJjbrSwPSsVu/4Y8U1UPFc4EPyv9qZaW2b5heQtbyUyG
'' SIG '' duXgQ0sile7CK0PBn9hotI5AT+6FOLkRxSPyZFjwFTJv
'' SIG '' TlehroikAtcqHs1L4d1j1ReJMluwXplaqJ0oUA4X7pbb
'' SIG '' YTtFUR3PElYLkkf8q672Zj1HrHBy55LnX80QucSDZJQZ
'' SIG '' vSWA4ejSIqXQugJ6oXeTW2XD7hd0vEGGKtwITIySjJEt
'' SIG '' nndEH2jWqHR32w5bMotWizO92WPISZ06xcXqMwvS8aMb
'' SIG '' 9Iu+2bNXizveBKd6IrIkri7HcMW+ToMmCPsLvalPmQjh
'' SIG '' EChyqs0CAwEAAaOCAVcwggFTMAwGA1UdEwEB/wQCMAAw
'' SIG '' FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/
'' SIG '' BAQDAgeAMHMGCCsGAQUFBwEBBGcwZTAqBggrBgEFBQcw
'' SIG '' AYYeaHR0cDovL3RzLW9jc3Aud3Muc3ltYW50ZWMuY29t
'' SIG '' MDcGCCsGAQUFBzAChitodHRwOi8vdHMtYWlhLndzLnN5
'' SIG '' bWFudGVjLmNvbS90c3MtY2EtZzIuY2VyMDwGA1UdHwQ1
'' SIG '' MDMwMaAvoC2GK2h0dHA6Ly90cy1jcmwud3Muc3ltYW50
'' SIG '' ZWMuY29tL3Rzcy1jYS1nMi5jcmwwKAYDVR0RBCEwH6Qd
'' SIG '' MBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0yMDQ4LTIwHQYD
'' SIG '' VR0OBBYEFEbGaaMOShQe1UzaUmMXP142vA3mMB8GA1Ud
'' SIG '' IwQYMBaAFF+a9W5czMx0mtTdfe8/2+xMgC7dMA0GCSqG
'' SIG '' SIb3DQEBBQUAA4IBAQB4O7SRKgBM8I9iMDd4o4QnB28Y
'' SIG '' st4l3KDUlAOqhk4ln5pAAxzdzuN5yyFoBtq2MrRtv/Qs
'' SIG '' JmMz5ElkbQ3mw2cO9wWkNWx8iRbG6bLfsundIMZxD82V
'' SIG '' dNy2XN69Nx9DeOZ4tc0oBCCjqvFLxIgpkQ6A0RH83Vx2
'' SIG '' bk9eDkVGQW4NsOo4mrE62glxEPwcebSAe6xp9P2ctgwW
'' SIG '' K/F/Wwk9m1viFsoTgW0ALjgNqCmPLOGy9FqpAa8VnCwv
'' SIG '' SRvbIrvD/niUUcOGsYKIXfA9tFGheTMrLnu53CAJE3Hr
'' SIG '' ahlbz+ilMFcsiUk/uc9/yb8+ImhjU5q9aXSsxR08f5Lg
'' SIG '' w7wc2AR1MIIFajCCBFKgAwIBAgIQbFnvqeEA4Q7jBrqP
'' SIG '' 4CklWTANBgkqhkiG9w0BAQUFADCByjELMAkGA1UEBhMC
'' SIG '' VVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMR8wHQYD
'' SIG '' VQQLExZWZXJpU2lnbiBUcnVzdCBOZXR3b3JrMTowOAYD
'' SIG '' VQQLEzEoYykgMjAwNiBWZXJpU2lnbiwgSW5jLiAtIEZv
'' SIG '' ciBhdXRob3JpemVkIHVzZSBvbmx5MUUwQwYDVQQDEzxW
'' SIG '' ZXJpU2lnbiBDbGFzcyAzIFB1YmxpYyBQcmltYXJ5IENl
'' SIG '' cnRpZmljYXRpb24gQXV0aG9yaXR5IC0gRzUwHhcNMTIw
'' SIG '' NjA3MDAwMDAwWhcNMjIwNjA2MjM1OTU5WjCBjDELMAkG
'' SIG '' A1UEBhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBv
'' SIG '' cmF0aW9uMR8wHQYDVQQLExZTeW1hbnRlYyBUcnVzdCBO
'' SIG '' ZXR3b3JrMT0wOwYDVQQDEzRTeW1hbnRlYyBDbGFzcyAz
'' SIG '' IEV4dGVuZGVkIFZhbGlkYXRpb24gQ29kZSBTaWduaW5n
'' SIG '' IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
'' SIG '' AQEAi0OvocSoTc3Q7sc2wKCKeXMo2HflwFTHNfe77Ruf
'' SIG '' 6ldw01UbJxmpy+ABBa4F8q3nBh3RBqith7kYhC8eHQli
'' SIG '' 090N9Re0MG9eUnYWjFZ7xZA6gn21rVjmAOcYBTbtMCCh
'' SIG '' 8OzDYvSZEBqU9vBXaMlyNr18kKgWFiClSQFRMqCW84ow
'' SIG '' OKuGoRWj8hwgV1BLuGTSsWzm5Dy2CCHES0CWF7PLZ9uG
'' SIG '' Qdlb/pgdRCQ66GmhGiRrs0gU8/QOg8VNMfu9r64hPGLr
'' SIG '' 6trYndfskR6zw0QeVB2Cm+1ZE+4wcONslOEsB9OPjOph
'' SIG '' yVyrS5gqh7naPjeDCjC6tUSY/e+9qoA1sVyt9wIDAQAB
'' SIG '' o4IBhjCCAYIwNAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUF
'' SIG '' BzABhhhodHRwOi8vb2NzcC52ZXJpc2lnbi5jb20wEgYD
'' SIG '' VR0TAQH/BAgwBgEB/wIBADBlBgNVHSAEXjBcMFoGBFUd
'' SIG '' IAAwUjAmBggrBgEFBQcCARYaaHR0cDovL3d3dy5zeW1h
'' SIG '' dXRoLmNvbS9jcHMwKAYIKwYBBQUHAgIwHBoaaHR0cDov
'' SIG '' L3d3dy5zeW1hdXRoLmNvbS9ycGEwNAYDVR0fBC0wKzAp
'' SIG '' oCegJYYjaHR0cDovL2NybC52ZXJpc2lnbi5jb20vcGNh
'' SIG '' My1nNS5jcmwwHQYDVR0lBBYwFAYIKwYBBQUHAwIGCCsG
'' SIG '' AQUFBwMDMA4GA1UdDwEB/wQEAwIBBjAqBgNVHREEIzAh
'' SIG '' pB8wHTEbMBkGA1UEAxMSVmVyaVNpZ25NUEtJLTItMjE0
'' SIG '' MB0GA1UdDgQWBBSjjs8ZQj0x4ashiYRty9l5orKyWjAf
'' SIG '' BgNVHSMEGDAWgBR/02Wnwt3su/AwCfNDOfoCrzMxMzAN
'' SIG '' BgkqhkiG9w0BAQUFAAOCAQEAavMdvF9N3gP5SUkdrT12
'' SIG '' HJa6G0Pm9IYCQnV4xwzC5Z3ENE8OqelKtL5BhIfq9Ie0
'' SIG '' TNsQSTv33xWQuoT4t0frW2VQ86NKcRAWexzh9dbtv1BW
'' SIG '' b/iZs6lRtkauxpfg55sMFT67KHsxowDzLouHSBKJgu8J
'' SIG '' X0kMkJ7I9paje5p1E8hH8D4/bwtQKWwreEww/ORgDBNA
'' SIG '' 1jh1qQd5ZP3KPOTvSJML4ApI/wdrOwKD0WbVueGY9A6f
'' SIG '' acQuVS4Bln1+hAyAdnU2y/1GYfRpzBqdZCu6BG7pEVLa
'' SIG '' EpmhWrCDxLxHgKYnTQB6NgM8vmGYY8ufBe6Ahe7dlZL3
'' SIG '' 7lDUY9yPpCR5vzCCBaQwggSMoAMCAQICEFmHbU0xEV7o
'' SIG '' X4yGqE/PbqowDQYJKoZIhvcNAQELBQAwgYwxCzAJBgNV
'' SIG '' BAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
'' SIG '' dGlvbjEfMB0GA1UECxMWU3ltYW50ZWMgVHJ1c3QgTmV0
'' SIG '' d29yazE9MDsGA1UEAxM0U3ltYW50ZWMgQ2xhc3MgMyBF
'' SIG '' eHRlbmRlZCBWYWxpZGF0aW9uIENvZGUgU2lnbmluZyBD
'' SIG '' QTAeFw0xNDAxMTQwMDAwMDBaFw0xNjAxMDcyMzU5NTla
'' SIG '' MIIBDjETMBEGCysGAQQBgjc8AgEDEwJVUzEZMBcGCysG
'' SIG '' AQQBgjc8AgECFAhEZWxhd2FyZTEdMBsGA1UEDxMUUHJp
'' SIG '' dmF0ZSBPcmdhbml6YXRpb24xEDAOBgNVBAUTBzI3NDgx
'' SIG '' MjkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIDApDYWxpZm9y
'' SIG '' bmlhMREwDwYDVQQHDAhTYW4gSm9zZTEjMCEGA1UECgwa
'' SIG '' QWRvYmUgU3lzdGVtcyBJbmNvcnBvcmF0ZWQxLDAqBgNV
'' SIG '' BAsMI0lsbHVzdHJhdG9yLCBJbkRlc2lnbiwgSW5Db3B5
'' SIG '' LCBNdXNlMSMwIQYDVQQDFBpBZG9iZSBTeXN0ZW1zIElu
'' SIG '' Y29ycG9yYXRlZDCCASIwDQYJKoZIhvcNAQEBBQADggEP
'' SIG '' ADCCAQoCggEBAMrOf/qkHn0tgwVmo5Bfpa5AjiSU7OF/
'' SIG '' 8E1m/Z8BfggsICRciV8j+JzZVZFWOUa5B8BGY6wI2WXt
'' SIG '' foOJAhzk2knRWNzchnwXDh5q8cUQf27+DHjoaWliQYOm
'' SIG '' T8aW0NXOfOR3twRaOCh16WnCl1bNJc/AQZoE2ARkNLag
'' SIG '' 1zWCx93emp1TpY+R4vVbvdJWYfK28AtUJWJn4q4TCbJM
'' SIG '' fNv4R30oCW5XLFiPBwAjGrzxlTOW+r+iyUawzHjA8qZG
'' SIG '' tBr6YFGTVtVIpNM+5gJXx7z5v/punHQQEf6gFs2xCi8m
'' SIG '' v77INHAtPyqRALfXLI63YfTh61DuuKsUocE7N1tTS+ne
'' SIG '' f60CAwEAAaOCAXswggF3MC4GA1UdEQQnMCWgIwYIKwYB
'' SIG '' BQUHCAOgFzAVDBNVUy1EZWxhd2FyZS0yNzQ4MTI5MAkG
'' SIG '' A1UdEwQCMAAwQgYDVR0gBDswOTA3BgtghkgBhvhFAQcX
'' SIG '' BjAoMCYGCCsGAQUFBwIBFhpodHRwOi8vd3d3LnN5bWF1
'' SIG '' dGguY29tL2NwczA5BgNVHR8EMjAwMC6gLKAqhihodHRw
'' SIG '' Oi8vZXZjcy1jcmwud3Muc3ltYW50ZWMuY29tL2V2Y3Mu
'' SIG '' Y3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMDMA4GA1Ud
'' SIG '' DwEB/wQEAwIHgDByBggrBgEFBQcBAQRmMGQwLAYIKwYB
'' SIG '' BQUHMAGGIGh0dHA6Ly9ldmNzLW9jc3Aud3Muc3ltYW50
'' SIG '' ZWMuY29tMDQGCCsGAQUFBzAChihodHRwOi8vZXZjcy1h
'' SIG '' aWEud3Muc3ltYW50ZWMuY29tL2V2Y3MuY2VyMB8GA1Ud
'' SIG '' IwQYMBaAFKOOzxlCPTHhqyGJhG3L2XmisrJaMA0GCSqG
'' SIG '' SIb3DQEBCwUAA4IBAQBRnDlLrWB8NUohBmiH6VRvWoBG
'' SIG '' WsOHwVtAPJaql6taSEX05msF9hQM/m1lHYsN6xU5I1do
'' SIG '' 17lLKtJM6yK2CyFvSAD2lLTm22mn/npPTCSVxgaVJHzy
'' SIG '' 8RDZbnfmyMsdK0MziCSBi8S1h43eSh+PbV9490iLEsoH
'' SIG '' zenmdp4BlMSyKSyCcZ9Jh8L4bLCKftr1RWsmo4rcoKbI
'' SIG '' GQXPZQaYcJ1iQdQMtD+XHH9EVpFWCPMt2E9PjQO2yP2z
'' SIG '' FeaHpyL1Bwh++W1DCYgQV07on50qwyPUPHDz4SGD8YLg
'' SIG '' vbI9RZVpj/Fn27yb6XNmePQdFTt1U3PD3pnBHEQUIKaZ
'' SIG '' V3s/npq3MYIEgTCCBH0CAQEwgaEwgYwxCzAJBgNVBAYT
'' SIG '' AlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlv
'' SIG '' bjEfMB0GA1UECxMWU3ltYW50ZWMgVHJ1c3QgTmV0d29y
'' SIG '' azE9MDsGA1UEAxM0U3ltYW50ZWMgQ2xhc3MgMyBFeHRl
'' SIG '' bmRlZCBWYWxpZGF0aW9uIENvZGUgU2lnbmluZyBDQQIQ
'' SIG '' WYdtTTERXuhfjIaoT89uqjAJBgUrDgMCGgUAoIGmMBkG
'' SIG '' CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQB
'' SIG '' gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJ
'' SIG '' BDEWBBTAZ7Mdi6Ru1OJ+M7V778Ee0E5I0zBGBgorBgEE
'' SIG '' AYI3AgEMMTgwNqA0gDIAQQBkAG8AYgBlACAASQBsAGwA
'' SIG '' dQBzAHQAcgBhAHQAbwByACAAQwBDACAAMgAwADEANTAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQCuNv66KO58VTF2/IaCN3Ks
'' SIG '' D94j6HkFbCKpcunCq0B4JZb+qJtQqey7lQjqyTsvB6ti
'' SIG '' kDy2+TfvvIJRbqJn6+8sxyB01eitBh3Prt7sjSv5Tmt+
'' SIG '' P6WzKLygZNSdfcf6AzZnENY8NcX7fRALb7vuh1Ko0Q7L
'' SIG '' +uWotQOLw5ipenZsbQRGbn8MMrNCewK2kuKDZNTstUnt
'' SIG '' NaCsLXXtFVj3z8AKbUzDJJTOMmqw13XOVA5TWVY9PO2z
'' SIG '' GmWuZNdMD7ajm4lwAa6t0AN/vvqPwWvOVZAPuFY2pmd0
'' SIG '' SW+on0kSxxce/F49NnLN5ykwnaxOyEZimameWpMLpyzT
'' SIG '' RBPwnjvAjJuFoYICCzCCAgcGCSqGSIb3DQEJBjGCAfgw
'' SIG '' ggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoT
'' SIG '' FFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
'' SIG '' eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENB
'' SIG '' IC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYFKw4DAhoF
'' SIG '' AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
'' SIG '' KoZIhvcNAQkFMQ8XDTE1MDUyODIxNTMyM1owIwYJKoZI
'' SIG '' hvcNAQkEMRYEFEdg0nW9IAbRhD/84+0N66fuzqsjMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBADGLKVK1o7EBop2arPW+C8YE
'' SIG '' kC3Sm+NXceFM0qt862IXBJA2XuzW1ykt1qy7z7hYvoI9
'' SIG '' lTQIVB0l/IgQN5Sv5wHRuny0fs7XYYa83uaiN3f/GCZI
'' SIG '' Arv27VZZDiflWL2Npmzf7jg9UBKtrhPp6gMhVN7DRqXG
'' SIG '' FhcveikIKQsNcCMajCqNxQ1AFoCsrbMD1Hi02ElZ8hPJ
'' SIG '' qXpmk3/gJSkVqkxZvb1w8H3qIFt65YasAFRpggERkIGz
'' SIG '' LeRmxUWBknVWu6viEESyyz4cX8xWWL2PWYU5UaB8uSMJ
'' SIG '' EQPmZl9TXOPaWUiV/1dDP9ZPaz17EaedHMig2fRRQnbl
'' SIG '' 4FGdLc+puX8=
'' SIG '' End signature block
