'Took the code From  https://www.robvanderwoude.com/vbstech_files_encoding.php
'Ive just made it into a function 
'And made the return values to fit OpenTextFile() function parameters for encoding
Function GetFileEncoding(filePath)
    Const adTypeBinary = 1

    Dim dicBOMs, objStream, strHead, strBOM, strType, strUTF7
    Dim i

    strType = "Unknown"
    strUTF7 = "38;39;2B;2F"

    Set dicBOMs = CreateObject("Scripting.Dictionary")
    dicBOMs.Add "0000FEFF", "UTF-32 (BE)"
    dicBOMs.Add "0EFEFF",   "SCSU"
    dicBOMs.Add "2B2F76",   "UTF-7"
    dicBOMs.Add "84319533", "GB-18030"
    dicBOMs.Add "DD736673", "UTF-EBCDIC"
    dicBOMs.Add "EFBBBF",   "UTF-8"
    dicBOMs.Add "F7644C",   "UTF-1"
    dicBOMs.Add "FBEE28",   "BOCU-1"
    dicBOMs.Add "FEFF",     "UTF-16 (BE)"
    dicBOMs.Add "FFFE",     "UTF-16 (LE)"
    dicBOMs.Add "FFFE0000", "UTF-32 (LE)"

    On Error Resume Next
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = adTypeBinary
    objStream.LoadFromFile filePath
    If Err.Number <> 0 Then
        GetFileEncoding = 0 ' Default to ASCII
        Exit Function
    End If

    objStream.Position = 0
    strHead = ""
    For i = 0 To 3
        strHead = strHead & UCase(Right("0" & Hex(AscB(objStream.Read(1))), 2))
        If Err.Number <> 0 Then
            GetFileEncoding = 0
            Exit Function
        End If
    Next
    objStream.Close
    Set objStream = Nothing
    On Error GoTo 0

    For i = 8 To 4 Step -2
        If strType = "Unknown" Then
            strBOM = Left(strHead, i)
            If dicBOMs.Exists(strBOM) Then
                If dicBOMs(strBOM) = "UTF-7" Then
                    If InStr(strUTF7, Right(strHead, 2)) Then strType = "UTF-7"
                Else
                    strType = dicBOMs(strBOM)
                End If
            End If
        End If
    Next

    ' Return encoding for VBScript's OpenTextFile
	Select Case strType
		Case "UTF-8"
			GetFileEncoding = -2 ' VBScript UTF-8
		Case "UTF-16 (LE)"
			GetFileEncoding = -1 ' VBScript Unicode
		Case "UTF-16 (BE)"
			GetFileEncoding = -1 ' Best handled as -1 (VBScript assumes LE)
		Case Else
			GetFileEncoding = 0 ' Default ANSI
	End Select

End Function

