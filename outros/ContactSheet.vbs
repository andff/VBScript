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

' ContactSheet.vbs

' DESCRIPTION

' This sample iterates through all images in a folder and 
' creates a named thumb nail for each of these images. The 
' number of rows and columns are hard-coded as the InputBox 
' VBScript function cannot be called within Illustrator.
' 
'*************************************************************

' Main Code - execution begins here

' Create the Application object and get a set of image files to work with
Set appRef = CreateObject("Illustrator.Application")
Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

' Creating a folder browser in VBScript can be a problem (relying on either Windows API calls
' or use of ActiveX controls which may not be present on a given system). Instead, use
' Illustrator's built-in JavaScript to display a file browser. DoJavaScript can return a value,
' in this example it's the platform specific full path to the chosen folder.

doJavaScript = "var imagesFolder = Folder.selectDialog(""Select images folder:""); if (imagesFolder) folderPath = imagesFolder.fsName;"
imagesFolderPath = appRef.DoJavaScript(doJavaScript)

If (fileSystemObject.FolderExists(imagesFolderPath)) Then

	doJavaScript = "var numRows = prompt(""Enter number of rows:"", ""4""); if (numRows) numRows = parseInt(numRows); else numRows = 0;"
	numRows = appRef.DoJavaScript(doJavaScript)

	doJavaScript = "var numCols = prompt(""Enter number of columns:"", ""4""); if (numCols) numCols = parseInt(numCols); else numCols = 0;"
	numCols = appRef.DoJavaScript(doJavaScript)
	
	If (numRows > 0 And numCols > 0) Then
		imagesFolderPathTemp = Replace( imagesFolderPath, "\", "\\")
		doJavaScript = "alert(""Path to the test files used: " & CStr(imagesFolderPathTemp) & "    Number of rows: " & numRows & "  Number of columns: " & numCols & "  "")"
		appRef.DoJavaScript doJavaScript

		' Call function DoIt to iterate through all images in the folder 
		' and create the contact sheet
		Call DoIt(imagesFolderPath, CStr(numRows), CStr(numCols), 24, 22)
	Else
		doJavaScript = "alert(""Script requires at least 1 row and 1 column!"");"
		appRef.DoJavaScript doJavaScript
	End If

End If
'*************************************************************

' This routine iterates through all images in the folder and creates a named
' thumb nail for each of these

Sub DoIt(sDirName, verticalCount, horizontalCount, horizontalSpacing, verticalSpacing)
          
        ' Get the test files from the folder sDirName
        Set fso = CreateObject("Scripting.FileSystemObject")
	Set fldr = fso.GetFolder(sDirName)
	Set fls = fldr.Files
        numImages = verticalCount * horizontalCount
        Dim numFiles
        numFiles = fls.count
        Dim fileNames(5)
        
	
	' Add a new document
	Set docRef = appRef.Documents.Add
        
        ' Determine the dimensions of the document
        docPageOrigin = docRef.PageOrigin
        docLeft = docPageOrigin(0)
  	docTop = docRef.Height - docPageOrigin(1)
        printableWidth = docRef.Width - docLeft * 2
        printableHeight = docRef.Height - docPageOrigin(1) * 2
        
        ' Calculate the size of the individual grid spaces that the images will be placed in
        ' Dim gridWidth, gridHeight As Double
       
        gridWidth = (printableWidth + horizontalSpacing) / horizontalCount
        gridHeight = (printableHeight + verticalSpacing) / verticalCount
        
        ' Calculate the size of the individual images based on the printable document area,
        ' and on the number of images to be placed across and down the page
       
        imageWidth = gridWidth - horizontalSpacing
        imageHeight = gridHeight - verticalSpacing
           
        ' Normalize the image size so we end up with a square image
        
        If imageWidth < imageHeight Then
        	imageSize = imageWidth
                xAdjustment = 0
                yAdjustment = (imageHeight - imageWidth) / 2
        Else
                imageSize = imageHeight
                xAdjustment = (imageWidth - imageHeight) / 2
                yAdjustment = 0
        End If
            
        ' Calculate the bounding box for the first image to be placed
        ' Dim imageLeft, imageRight, imageTop, imageBottom As Double
        imageLeft = docLeft + xAdjustment
        imageRight = imageLeft + imageSize
            
        imageTop = docTop - yAdjustment
        imageBottom = imageTop - imageSize
            
        ' Iterate over the images in the list, positioning then one at a time
        ' After they are positioned, place a border around then to make them stand out better
        
        i = 1
        num = 0
        Dim aFile
        For each aFile in fls
		' Ignore system files.
		If not aFile.attributes and 4 Then
			num = num + 1
            		fileNames(num) = CStr(aFile)
		End If
        Next
        For j = 1 to numImages
        	
                Call AddRasterItemToPage(docRef,  fileNames(i), imageLeft, imageTop, (imageSize))
                    
                Call MakeFramingRectangle(docRef, imageLeft, imageTop, imageRight, imageBottom)
                Set fileObject = fso.getFile(fileNames(i))
                
                Call MakeTextLabel(docRef, CStr(fileObject.Name), imageLeft + imageSize / 2, imageBottom - 2)
                
                ' Calculate a new image position for the next iteration of the loop
                If j Mod horizontalCount <> 0 Then
                    ' Not at the end of row yet, move to next position in the row
                    imageLeft = imageLeft + gridWidth
                    imageRight = imageRight + gridWidth
                Else
                    ' If at the end of a row, first check to see if we have run out of rows
                    If j / horizontalCount >= verticalCount Then
                        Exit For
                    End If
                    
                    imageLeft = docLeft + xAdjustment
                    imageRight = imageLeft + imageSize
                        
                    imageTop = imageTop - gridHeight
                    imageBottom = imageBottom - gridHeight
                End If
                if (i = num) then
                    i = 1
                else
                    i = i + 1
                end if
        Next
            
        Set docRef = Nothing
        Set appRef = Nothing
End Sub

'*************************************************************

' This routine makes the labels for the thumbnails

Sub MakeTextLabel(aDocument, theText, xCenter, yVertPos)
                          
        ' Create the new text label
        Set aTextItem = aDocument.TextFrames.Add
        aTextItem.Contents = theText
        
        ' Calculate the final position and move the text label there
        aTextItem.Position = Array(xCenter - (aTextItem.Width / 2), yVertPos)
        
        ' Set the color of the text to default Illustrator color:
        ' No stroke & blank fill
        Set blackCMYK = CreateObject("Illustrator.CMYKColor")
        blackCMYK.Black = 100
        blackCMYK.Cyan = 0
        blackCMYK.Magenta = 0
        blackCMYK.Yellow = 0
        
        Set aTextRange = aTextItem.TextRange
        Set noStroke = CreateObject("Illustrator.NoColor")
        aTextRange.CharacterAttributes.StrokeColor = noStroke
        aTextRange.CharacterAttributes.FillColor = blackCMYK

End Sub

'*************************************************************

' This routine adds all images in the folder as RasterItems 
' to the Document

Sub AddRasterItemToPage(aDocument, theFile, imageLeft, imageTop, imageSize)
        
        ' Create a new raster item and link it to the image file
        Set aPlacedItem = aDocument.PlacedItems.Add
   	aPlacedItem.File = theFile
        
        ' Get the raster item dimensions
        itemWidth = aPlacedItem.Width
        itemHeight = aPlacedItem.Height
        
        ' Calculate the scale factors so the raster item can be fitted to the grid
        X_ScaleFactor = (imageSize / itemWidth) * 100
        Y_ScaleFactor = (imageSize / itemHeight) * 100
        
        ' Determine which of the scale factors to use
        If X_ScaleFactor < Y_ScaleFactor Then
            scaleFactor = X_ScaleFactor
        Else
            scaleFactor = Y_ScaleFactor
        End If
        
        ' Actually resize the raster item
        Call aPlacedItem.Resize(scaleFactor, scaleFactor)
                
        ' Get the new raster item dimensions
        itemWidth = aPlacedItem.Width
        itemHeight = aPlacedItem.Height
        
        ' Move the raster item to the middle of its grid cell
        aPlacedItem.Position = Array(imageLeft + (imageSize - itemWidth) / 2, _
                                     imageTop - (imageSize - itemHeight) / 2)
        
        Exit Sub

End Sub

'*************************************************************

Sub MakeFramingRectangle(aDocument, left, top, right, bottom)
        
        ' Add but not show the rectangle yet
       
        Set myPath = aDocument.PathItems.Rectangle(top, left, right - left, top - bottom)
        myPath.Hidden = True
        
        ' Make a black color
        Set blackCMYK = CreateObject("Illustrator.CMYKColor")
        
        blackCMYK.Black = 100
        blackCMYK.Cyan = 0
        blackCMYK.Magenta = 0
        blackCMYK.Yellow = 0
        
        ' Set the Color to 100% black
        myPath.StrokeColor = blackCMYK
        
        ' Make sure it is not filled
        myPath.Filled = False
        
        ' Show the resulting path
        myPath.Hidden = False
End Sub

'*************************************************************

'' SIG '' Begin signature block
'' SIG '' MIIYxAYJKoZIhvcNAQcCoIIYtTCCGLECAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFHAor4OrQCzw
'' SIG '' Lk7+i/OlMnNH5JenoIITrzCCA+4wggNXoAMCAQICEH6T
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
'' SIG '' BDEWBBQAyHJyRvt1hcy280luyKqZmeq+0DBGBgorBgEE
'' SIG '' AYI3AgEMMTgwNqA0gDIAQQBkAG8AYgBlACAASQBsAGwA
'' SIG '' dQBzAHQAcgBhAHQAbwByACAAQwBDACAAMgAwADEANTAN
'' SIG '' BgkqhkiG9w0BAQEFAASCAQB8kwNdZCko+bX9UplyQTyS
'' SIG '' +XDeT5Ya0iCN/qVcqRrxIXWqbmv2Ol9rZTz8EAgxSkYp
'' SIG '' rK+WPg28fW64VP77317YDelpuHThZ36LIpZHXGB9TeqI
'' SIG '' JhrnIKcooTvPdqmOQ9RbRieXdhBOnnmovex3kpFbc38E
'' SIG '' 5gFG934OXXAUzqnTIkTwxzQiaHP7M/VfAylbsXVdgC3+
'' SIG '' pHE3spyOQUOCRT29rJt7NQNESIyAl+UAm2XwoyrdVTWY
'' SIG '' exF9OE3/eTU5HnlaYeDry2u1S5W0BP3E9ZxUj4WAEtpz
'' SIG '' YbfcQxBP9S9BZrEoqxiw+wWwm/d8oaCGEGrJB3nTS4A6
'' SIG '' go/VxCw5UVDmoYICCzCCAgcGCSqGSIb3DQEJBjGCAfgw
'' SIG '' ggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoT
'' SIG '' FFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
'' SIG '' eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENB
'' SIG '' IC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYFKw4DAhoF
'' SIG '' AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
'' SIG '' KoZIhvcNAQkFMQ8XDTE1MDUyODIxNTAxNFowIwYJKoZI
'' SIG '' hvcNAQkEMRYEFJ1aKYlhKSimmnGEtpZweAaHiytmMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBAI8MjxvL7bhwW7ETEFyyn1Bn
'' SIG '' WEiih/GBZqxco2KghPGBygmUyj/QDne0DoppsmoXlDDy
'' SIG '' KMAlDk9xO9xmpftkXRL8cg0upEPwKNRAps6qB3H4WceK
'' SIG '' OK/pYbrMY+X3lb7ax79rOJz5awlHEPydSXh63OVvBx0m
'' SIG '' 9QjbecAgKUiU1BffqTUTG6KxAr3wwDqxi2wdUIZfYA/m
'' SIG '' ohX8TxaUgip1+8ojlw5yUwPpHkrsFVe6p8vKoSRjeYQv
'' SIG '' w2sismYFGC89vqzSuMSJNYjOu5hOqUo7Y6NmVS+fnpZD
'' SIG '' 7W0o6q6U9yJGZ3Ge6SbDgFbxOwS/qPfqDn2IR/pRGlC3
'' SIG '' tqMxJL8b2AU=
'' SIG '' End signature block
