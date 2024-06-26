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

' WebGallery.vbs

' DESCRIPTION

' This example demonstrates how a gallery can be created from 
' several Illustrator compatible files.
' The script first creates .ai and .pdf files then saves them 
' in a 'WebGallery' folder on the desktop, Illustrator 
' will then scan the WebGallery folder for files 
' and load them into Illustrator. Once in Illustrator the 
' documents are exported to JPEG and HTML format and an HTML file 
' is created which displays the files in a gallery format.
'
'*************************************************************
    Dim objApp
    Dim objExportOptions
    Dim objFileSys
    Dim objFile
    Dim objFiles
    Dim objFolder
    Dim objFolderGall
    Dim objFolderImg
    Dim objFolderPage
    Dim objFolderThumb
    Dim theFiles
    Dim f1 
    Dim myName
    Dim myPath
    Dim htmlFrame
    Dim htmlPage
    Dim htmlIndex
    Dim htmlIndexRow
    Dim htmlIndexRows
    Dim success
    Dim objFolderSamp
    Dim newItem
    Dim itemColor
    Dim pdfSvOpts
    Dim textRef
    
    const DESKTOP = &H0&
	set myShell = CreateObject("Shell.Application")
	set myDesktopFolder = myShell.Namespace(DESKTOP)
	set myDesktopFolderItem = myDesktopFolder.Self
	myDesktopPath = myDesktopFolderItem.Path
    
    Set objApp = CreateObject("Illustrator.Application")
    Set objExportOptions = CreateObject("Illustrator.ExportOptionsJPEG")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSys.GetFolder(myDesktopPath)
    Set pdfSvOpts = CreateObject("Illustrator.PDFSaveOptions")
    Set aiSvOpts = CreateObject("Illustrator.IllustratorSaveOptions")
    myMsg = "alert(""Creates 3 documents, saves the documents to a folder " & _
    "called \'WebGallery\' on the Desktop, exports each document as jpeg " & _
    "and html then creates the html gallery file - index.html."")"
    objApp.DoJavaScript myMsg
    
    ' check for folders and create if they don't exist
    If (objFileSys.FolderExists(objFolder & "\WebGallery")) Then
        Set objFolderSamp = objFileSys.GetFolder(objFolder & "\WebGallery")
	Else
        Set objFolderSamp = objFolder.SubFolders.Add("WebGallery")
    End If
    
	If (objFileSys.FolderExists(objFolderSamp & "\gallery")) Then
		Set objFolderGall = objFileSys.GetFolder(objFolderSamp & "\gallery")
	Else
		Set objFolderGall = objFolderSamp.SubFolders.Add("gallery")
	End If
	
	If (objFileSys.FolderExists(objFolderGall & "\images")) Then
		Set objFolderImg = objFileSys.GetFolder(objFolderGall & "\images")
	Else
		Set objFolderImg = objFolderGall.SubFolders.Add("images")
	End If

	If (objFileSys.FolderExists(objFolderGall & "\pages")) Then
		Set objFolderPage = objFileSys.GetFolder(objFolderGall & "\pages")
	Else
		Set objFolderPage = objFolderGall.SubFolders.Add("pages")
	End If
	
	If (objFileSys.FolderExists(objFolderGall & "\thumbnails")) Then
		Set objFolderThumb = objFileSys.GetFolder(objFolderGall & "\thumbnails")
	Else
		Set objFolderThumb = objFolderGall.SubFolders.Add("thumbnails")
	End If    
    
    'reserved characters (can't be in filenames processed)
    ' ^ (repeating index rows)
    ' ~ (image number)
    
    'create and save a document with a star in it
    Set docRef = objApp.Documents.Add
    Set newItem = docRef.PathItems.Star(300,400,100,50,5)
    Set itemColor = CreateObject("Illustrator.CMYKColor")
    itemColor.Magenta = 100
    newItem.FillColor = itemColor
    docRef.SaveAs (objFolderSamp & "\star.ai")

    'create and save a document with an ellipse in it
    Set docRef = objApp.Documents.Add
    Set newItem = docRef.PathItems.Ellipse(200,200,200,100)
    Set itemColor = CreateObject("Illustrator.CMYKColor")
    itemColor.Yellow = 100
    newItem.FillColor = itemColor
    docRef.SaveAs (objFolderSamp & "\ellipse.ai")

    'create and save a pdf file with some text
    Set docRef = objApp.Documents.Add
    Set textRef = docRef.TextFrames.Add()
    textRef.Contents = "Here is text from the PDF file."
    textRef.Top = 400
    textRef.Left = 100
    textRef.TextRange.CharacterAttributes.Size = 18
    objApp.Redraw
    pdfSvOpts.AcrobatLayers = true
    docRef.SaveAs objFolderSamp & "\text.pdf",pdfSvOpts

    'standard frameset html
    htmlFrame = ""
    htmlFrame = htmlFrame & "<HTML>" & vbCrLf
    htmlFrame = htmlFrame & "<HEAD>" & vbCrLf
    htmlFrame = htmlFrame & "<TITLE>Web Gallery</TITLE>" & vbCrLf
    htmlFrame = htmlFrame & "<META name=""generator"" content=""Adobe Illustrator Script-Generated Web Gallery"">" & vbCrLf
    htmlFrame = htmlFrame & "<META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
    htmlFrame = htmlFrame & "</HEAD>" & vbCrLf
    htmlFrame = htmlFrame & "<FRAMESET frameborder=1 cols=""20%,80%"">" & vbCrLf
    htmlFrame = htmlFrame & "<FRAME src=""index.html""  NAME=""LeftFrame""  scrolling=YES>" & vbCrLf
    htmlFrame = htmlFrame & "<FRAME src=""pages/1.html""  name=""RightFrame"" scrolling=YES>" & vbCrLf
    htmlFrame = htmlFrame & "<NOFRAMES>" & vbCrLf
    htmlFrame = htmlFrame & "<BODY>" & vbCrLf
    htmlFrame = htmlFrame & "Viewing this page requires a browser capable of displaying frames." & vbCrLf
    htmlFrame = htmlFrame & "</BODY>" & vbCrLf
    htmlFrame = htmlFrame & "</NOFRAMES>" & vbCrLf
    htmlFrame = htmlFrame & "</FRAMESET>" & vbCrLf
    htmlFrame = htmlFrame & "</HTML>" & vbCrLf
    'write frameset
    Set objFile = objFileSys.CreateTextFile(objFolderGall.Path & "\" & "frameset.html", True, False)
    objFile.Write htmlFrame
    objFile.Close
    
    'standard thumb index html
    htmlIndex = ""
    htmlIndex = htmlIndex & "<HTML>" & vbCrLf
    htmlIndex = htmlIndex & "<HEAD>" & vbCrLf
    htmlIndex = htmlIndex & "<TITLE>Web Gallery</TITLE>" & vbCrLf
    htmlIndex = htmlIndex & "<META name=""generator"" content=""Adobe Illustrator Script-Generated Web Gallery"">" & vbCrLf
    htmlIndex = htmlIndex & "<META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
    htmlIndex = htmlIndex & "</HEAD>" & vbCrLf
    htmlIndex = htmlIndex & "<BODY bgcolor=""#FFFFFF""  text=""#000000""  link=""#FF0000""  vlink=""#CCCC99""  alink=""#0000FF"" >" & vbCrLf
    htmlIndex = htmlIndex & "Thumbnails with hyperlinks" & vbCrLf
    htmlIndex = htmlIndex & "<P><TABLE border=""0"" cellspacing=""2"" cellpadding=""0"">" & vbCrLf
    htmlIndex = htmlIndex & "^" & vbCrLf
    'repeating row
    htmlIndexRow = ""
    htmlIndexRow = htmlIndexRow & "<TR><TD>" & vbCrLf
    htmlIndexRow = htmlIndexRow & "<P><A href=""pages/~.html"" target=""RightFrame""><IMG src=""thumbnails/~.jpg"" border=""0"" alt=1 align=""BOTTOM""></A><BR><BR>" & vbCrLf
    htmlIndexRow = htmlIndexRow & "<A href=""pages/~.html"" target=""RightFrame""><FONT size=""2""  face=""Arial"" >~.jpg</FONT></A><FONT size=""2""  face=""Arial"" ></FONT><BR><BR>" & vbCrLf
    htmlIndexRow = htmlIndexRow & "</TD></TR>" & vbCrLf
    '
    htmlIndex = htmlIndex & "</TABLE>" & vbCrLf
    htmlIndex = htmlIndex & "</BODY>" & vbCrLf
    htmlIndex = htmlIndex & "</HTML>" & vbCrLf
    'clear running string for final insertion
    htmlIndexRows = ""
    
    'standard image page html
    htmlPage = ""
    htmlPage = htmlPage & "</HTML>" & vbCrLf
    htmlPage = htmlPage & "<HTML>" & vbCrLf
    htmlPage = htmlPage & "<HEAD>" & vbCrLf
    htmlPage = htmlPage & "<TITLE>~</TITLE>" & vbCrLf
    htmlPage = htmlPage & "<META name=""generator"" content=""Adobe Illustrator Script-Generated Web Gallery"">" & vbCrLf
    htmlPage = htmlPage & "<META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
    htmlPage = htmlPage & "</HEAD>" & vbCrLf
    htmlPage = htmlPage & "<BODY bgcolor=""#FFFFFF""  text=""#000000""  link=""#FF0000""  vlink=""#CCCC99""  alink=""#0000FF"" >" & vbCrLf
    htmlPage = htmlPage & "<TABLE border=""0"" cellpadding=""5"" cellspacing=""2"" width=""100%"" bgcolor=""#F0F0F0"" >" & vbCrLf
    htmlPage = htmlPage & "<TR>" & vbCrLf
    htmlPage = htmlPage & "<TD><FONT size=""2"" face=""Arial"" >Web Gallery / ~<BR><BR>" & Date & "</FONT>" & vbCrLf
    htmlPage = htmlPage & "</TD>" & vbCrLf
    htmlPage = htmlPage & "</TR>" & vbCrLf
    htmlPage = htmlPage & "</TABLE>" & vbCrLf
    htmlPage = htmlPage & "<P><CENTER><IMG src=""../images/~.jpg""  border=""0"" alt=1></CENTER></P>" & vbCrLf
    htmlPage = htmlPage & "</BODY>" & vbCrLf
    
    'loop thru all files that Illustrator can open
        Set theFiles = objFolderSamp.Files
        Dim i
        i = 0
        For Each f1 In theFiles
        i = i + 1
            myPath = f1.Path
            'open document in illustrator
            objApp.Open (myPath)
            
            If objApp.Documents.Count > 0 Then
                objExportOptions.HorizontalScale = 20
                objExportOptions.VerticalScale = 20
                objApp.Documents(1).Export objFolderThumb.Path & "\" & i & ".jpg", 1, objExportOptions
                objExportOptions.HorizontalScale = 75
                objExportOptions.VerticalScale = 75
                objApp.Documents(1).Export objFolderImg.Path & "\" & i & ".jpg", 1, objExportOptions
                objApp.Documents(1).Close (2) 'aiDoNotSaveChanges
                htmlIndexRows = htmlIndexRows & Replace(htmlIndexRow, "~", i)
                'and creating a page html file
                Set objFile = objFileSys.CreateTextFile(objFolderPage.Path & "\" & i & ".html", True, False)
                objFile.Write Replace(htmlPage, "~", i)
                objFile.Close
            End If
        Next
    
        'save thumbnail index
        htmlIndex = Replace(htmlIndex, "^", htmlIndexRows)
        Set objFile = objFileSys.CreateTextFile(objFolderGall.Path & "\" & "index.html", True, False)
        objFile.Write Replace(htmlIndex, "~", i)
        objFile.Close
'' SIG '' Begin signature block
'' SIG '' MIIYxAYJKoZIhvcNAQcCoIIYtTCCGLECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFNlCYHyaBjXj
'' SIG '' QaKgulyfTpGtF18loIITrzCCA+4wggNXoAMCAQICEH6T
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
'' SIG '' BDEWBBTCzDovzxDkUP/WhytiWYZws++GcDBGBgorBgEE
'' SIG '' AYI3AgEMMTgwNqA0gDIAQQBkAG8AYgBlACAASQBsAGwA
'' SIG '' dQBzAHQAcgBhAHQAbwByACAAQwBDACAAMgAwADEANTAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQBh6eo4hJ9/+0RWDwGS/aG+
'' SIG '' LhM4FMTfme36IvYU6bJZvYlxkevUEWcVRUzaGq23e68e
'' SIG '' FpqWB+19zdveiAXD64HkYRo5bxvfjoQHV8FUhrQKYA67
'' SIG '' u+cacn1uVVZ77gBvBGVb5bacKoOkntms6QALQlog1y2x
'' SIG '' xZPIGsGIx1wUgN+AupwZ75Iqd2LcOW0CMKA8EfM/wtjk
'' SIG '' i3edGkxB8eq8adNlnJjh04hwwR0g5yb12W5oG+EwS8Vn
'' SIG '' c5ekoq3UOPIOBBTuMKcf+aeNINDdzQXDUw3A2EXiBKQq
'' SIG '' dvgR1J/sSD3elueCQaSyv4Tdq4sRvlDAIOoYrEVIVfim
'' SIG '' asIxb71+0L65oYICCzCCAgcGCSqGSIb3DQEJBjGCAfgw
'' SIG '' ggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoT
'' SIG '' FFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
'' SIG '' eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENB
'' SIG '' IC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYFKw4DAhoF
'' SIG '' AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
'' SIG '' KoZIhvcNAQkFMQ8XDTE1MDUyODIxNTIzN1owIwYJKoZI
'' SIG '' hvcNAQkEMRYEFLYgmf9r/4/xTnD4p5g6EjxjsGstMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBAB48epVNe+L2679nSBotfmCc
'' SIG '' 0/9cWe5dzI1756OCniD+kqHl6nTMQWsHAU1ju3ujoC2/
'' SIG '' JLwyWIs3o+WiuRqWZqTbUuIczoyH9K/D+nd3TPdq4JFf
'' SIG '' ZW3n6fvFUwKr5J4BCQ813VduLoAo7BtKKJkShsqjp5ZJ
'' SIG '' XcxzZbaLl17gNwF8ZLzQIzKxXC18Bs64ajbUcWMBJTRI
'' SIG '' ZcuL2HMdvfN4tBlPP47c/j9U+MXmGaGWNIn2AC22wFKZ
'' SIG '' S5tLp4bswFSPBeq0vpqp3ELPle+semWUexLPS5h6IZVr
'' SIG '' N8+6V8wrxMFvqVdeUiB7FG+ylxs8Z0GTjYgz3oq4pNNE
'' SIG '' 03LzzPRFfLA=
'' SIG '' End signature block
