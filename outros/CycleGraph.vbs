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

' CycleGraph.vbs

' DESCRIPTION

' This sample accepts radius and step values from the developer
' and draws a cyclic graph figure from it. 
' To draw different objects try the values given in the 
' Examples1.ai file.
' 
'*************************************************************

' Create the general AI objects
Set appObj = CreateObject("Illustrator.Application")
Set docObj = appObj.Documents.Add

' Set the document properties
docObj.DefaultFilled = False
docObj.DefaultStroked = True
pathWidth = 1
docObj.DefaultStrokeWidth = pathWidth
pathOpacity = 100

Set groupObj = docObj.ActiveLayer.GroupItems.Add

myMsg = "alert(""Values used to draw the CycleGraph: Big Radius - 100, Small Radius - 35, Pen Distance - 30, Points/Cycle - 50, Cycles - 7. Change values in the script to draw different objects."")"

appObj.DoJavaScript myMsg

' Change values here to create different objects from Examples1.ai
bigRadius = 100
smallRadius = 35
penDistance = 30
stepsPerCycle = 50
cycles = 7

' OR
' Get the values from the user
' bigRadius = InputBox("Enter Big Radius", "CycleGraph", 100)
' smallRadius = InputBox("Enter Small Radius", "CycleGraph", 35)
' penDistance = InputBox("Enter Pen Distance", "CycleGraph", 30)
' stepsPerCycle = InputBox("Enter Points/Cycle", "CycleGraph", 50)
' cycles = InputBox("Enter Cycles", "CycleGraph", 7)

' set center of large circle at middle of page
docWidth = docObj.Width
docHeight = docObj.Height
BigX = docWidth / 2
BigY = docHeight / 2

' set small circle at top of large circle
SmallX = BigX
SmallY = BigY + bigRadius - smallRadius
radiusDifference = bigRadius - smallRadius

' initial angles
alphaAngle = 0
betaAngle = 0

' initialize pi value, step sizes
pi = 4 * Atn(1)
bigStepSize = (2 * pi) / stepsPerCycle
smallStepSize = bigStepSize * bigRadius / smallRadius

' initialize trailing point
prevX = SmallX
prevY = SmallY + penDistance
' for each step
For i = 0 To stepsPerCycle * cycles
    alphaAngle = alphaAngle + bigStepSize
    betaAngle = betaAngle - smallStepSize
    SmallX = BigX + (radiusDifference * Sin(alphaAngle))
    SmallY = BigY + (radiusDifference * Cos(alphaAngle))
    X = SmallX + (penDistance * Sin(betaAngle))
    Y = SmallY + (penDistance * Cos(betaAngle))
    CreateLine groupObj, prevX, prevY, X, Y, pathOpacity
    prevX = X
    prevY = Y
Next

' ****************************************************************

' Factor inToFactor into the array outArray

Private Function Factor(inToFactor, outArray())
	
	If (inToFactor > 1) Then
	    ReDim outArray(Sqr(inToFactor))
	    factorCount = 0
	    toFactor = inToFactor
	    ' Factor out all powers of 2
	    Do While ((toFactor Mod 2) = 0)
		outArray(factorCount) = 2
		factorCount = factorCount + 1
		toFactor = toFactor / 2
	    Loop
	    limit = toFactor
	    ' Walk through the odd numbers, testing each one
	    For testFactor = 3 To limit Step 2
		Do While ((toFactor Mod testFactor) = 0)
		    outArray(factorCount) = testFactor
		    factorCount = factorCount + 1
		    toFactor = toFactor / testFactor
		Loop
	    Next
	    ReDim Preserve outArray(factorCount - 1)
	    Factor = factorCount - 1
	Else
	    ' if number to factor is less than 1, return single factor 1
	    ReDim outArray(1)
	    factorCount = 0
	    outArray(0) = 1
	    Factor = 0
	End If
	
End Function

' ****************************************************************

' Calculate the cycles using the radii entered

Private Sub CalculateCycles
	
	' The required number of cycles (times around the Big circle) is
	' determined by the largest common multiple of the radii of the
	' Big and Small circles
	If (StrComp(bigRadiusText.Text, "") And StrComp(smallRadiusText.Text, "")) Then
	    toFactor1 = bigRadiusText.Text
	    toFactor2 = smallRadiusText.Text

	    ' Get the factors of each radius
	   
	    factors1 = Factor(toFactor1, factor1List)
	    factors2 = Factor(toFactor2, factor2List)

	    ' Cross off all factors held in common
	    For i = 0 To factors2
		For j = 0 To factors1
		    If (factor1List(j) = factor2List(i)) Then
			factor1List(j) = 1
			factor2List(i) = 1
		    End If
		Next
	    Next
	    cycles = 1
	    ' and multiply together all remaining factors in the small radius
	    For i = 0 To factors2
		cycles = cycles * factor2List(i)
	    Next
	    ' to determine the number of cycles
	    cyclesText.Text = cycles
	End If

End Sub

' ****************************************************************

' Create the path items

Private Sub CreateLine(inGroupItem, inStartX, inStartY, inEndX, inEndY, inOpacity)
	
	Set pathItem = inGroupItem.PathItems.Add
	pathItem.SetEntirePath Array(Array(inStartX, inStartY), Array(inEndX, inEndY))
	pathItem.Opacity = inOpacity

End Sub

' ****************************************************************

' Calculate the required number of cycles every time the Big radius is changed

Private Sub bigRadiusText_Change
	
	CalculateCycles

End Sub

' ****************************************************************

' Calculate the required number of cycles every time the Small radius is changed

Private Sub smallRadiusText_Change
	
	CalculateCycles

End Sub



'' SIG '' Begin signature block
'' SIG '' MIIYxAYJKoZIhvcNAQcCoIIYtTCCGLECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFLOSR+6QqZQl
'' SIG '' d65lljUeS2rfifN0oIITrzCCA+4wggNXoAMCAQICEH6T
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
'' SIG '' BDEWBBSgRIp+Dw6PQNWYgu43BmXplzXeCTBGBgorBgEE
'' SIG '' AYI3AgEMMTgwNqA0gDIAQQBkAG8AYgBlACAASQBsAGwA
'' SIG '' dQBzAHQAcgBhAHQAbwByACAAQwBDACAAMgAwADEANTAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQAIsvr+SsDy7FN89Kt30H+G
'' SIG '' kA+FK+9Fzqt9tlCNVGBILBabXMhmytfKuDFuHzTzXq40
'' SIG '' EdHp/UGI1Zp8g/hpgLnWiR8uKx4KoEjXdQ/GEPrIkY9W
'' SIG '' X8l5+r07JMDXd6KZt+zqc1IPG95Pe08SguwHHf/4fTCb
'' SIG '' vIVD/XdS4+lsqf+Mjz2/LRJgG+h0lonvtQcJD/NmoIXu
'' SIG '' RbR7UP4iP55JQHrgPcxYrzquKzNQ524BsvNLcslJZt7+
'' SIG '' ByNW/go+5ptAK5f3+a2nwmqf+a65svu8g6dkLnE9lLu+
'' SIG '' TFMa+i5WMGccjVuQathLEIzaQrqZcCQKFSAi6eUrKNPi
'' SIG '' BQOS1QFg5n26oYICCzCCAgcGCSqGSIb3DQEJBjGCAfgw
'' SIG '' ggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoT
'' SIG '' FFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
'' SIG '' eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENB
'' SIG '' IC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYFKw4DAhoF
'' SIG '' AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
'' SIG '' KoZIhvcNAQkFMQ8XDTE1MDUyODIxNTAyMVowIwYJKoZI
'' SIG '' hvcNAQkEMRYEFGP1Z/MVJiBExP4PFhH+YbgH+0gmMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBAIEcB+2dhV0G804CgJUCADSy
'' SIG '' Oi/PY+JsuGBQJwkPY9vYHWNUsrW3/3KY8xhoaw+bJSmz
'' SIG '' JWXY/2zMQ8Sd2mq9ptZL80sZIvn9qdpwNDt+UnMgiQ/5
'' SIG '' ZSYozaYX32D832UX/nZF4o/nxRVHyXd3Q/tsOZfbHMyW
'' SIG '' ttv+OvnzYksKMt9PsSqf6n1GY5xBur306oXPVgqNjAQe
'' SIG '' TzeIFc7Vx+zDQiSsvHYsNqbdOrgMZF3xDx35kWoDSx6r
'' SIG '' dmUjxZNe3rxe4iImScg+phSselPVsterecHyv+Z1/rbs
'' SIG '' JBOib93qok+ZEUi+RmdMEoXt9IWBWVNP9bGLH8PC/m/M
'' SIG '' qDuwlUFAPao=
'' SIG '' End signature block
