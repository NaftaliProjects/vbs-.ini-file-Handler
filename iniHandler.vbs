'import other functions
Sub Include (strFile)
	'Create objects for opening text file
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(strFile, 1)

	'Execute content of file.
	ExecuteGlobal objTextFile.ReadAll

	'CLose file
	objTextFile.Close

	'Clean up
	Set objFSO = Nothing
	Set objTextFile = Nothing
End Sub


'includes
Include("getFileEncoding.vbs")


'-----------------------Functions--------------------------------------------------------
'Function: isLabel(line)
'Checks if the line is a label (starts with [ and ends with ])
'Returns 1 if true, 0 if false
Function isWantedLabel(line, wantedLabel)
    Dim firstChar, lastChar, middleString
    line = Trim(line)
   
    firstChar = Left(line, 1)
    lastChar = Right(line, 1)
    middleString = Replace(line, "[", "")
	middleString = Replace(middleString, "]", "")
   
    If firstChar = "[" And lastChar = "]" And Trim(middleString) = Trim(wantedLabel) Then
        isWantedLabel = true
    Else
        isWantedLabel = false
    End If
End Function



'---------------------------------------------------------------------
'Function: getKeyValue(line, key)
'gets the value of a given key
'Returns value if found, returns empty string if did not found
Function getKeyValue(line, key)
    Dim parts, currentKey

    line = Trim(line)
    If InStr(line, "=") = 0 Then
        getKeyValue = ""
        Exit Function
    End If

    parts = Split(line, "=")
    If UBound(parts) = 1 Then
        currentKey = Trim(parts(0))
        If LCase(currentKey) = LCase(Trim(key)) Then
            getKeyValue = line ' return full line
        Else
            getKeyValue = ""
        End If
    Else
        getKeyValue = ""
    End If
End Function



'---------------------------------------------------------------------
'Function: readFromIni(path, label, key)
'open path file, search for key in the given label 
'Returns value if found, returns empty string if did not found
Function readFromIni(path, label, key)
	section = trim(section)
	key = trim(key)
	
	Dim line, char, asciiCode
    Dim fso1, file
 
    Const ForReading = 1
    
	Dim encoding, filePath
	encoding = GetFileEncoding(path)
	
	Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set file  = fso1.OpenTextFile(path, ForReading, False, encoding) 'if needed open file as unicode
	
	searchingForLabel = 1
	readFromIni = "" 
	
	While Not file.AtEndOfStream
		line = file.ReadLine
		If searchingForLabel = 1 Then
            If isWantedLabel(line, label) = true Then
                searchingForLabel = 0
            End If
        Else
            wantedKey = getKeyValue(line, key)
            If wantedKey <> "" Then
                readFromIni = split(wantedKey,"=")(1)
                Exit Function
            End If
        End If
	Wend
	file.Close
End Function




'---------------------------------------------------------------------
'Function: WriteToIni(path, section, key, newValue)
'open path file, search for key in the given label and overWrite it with newValue
'Returns true if success, else returns false
Function WriteToIni(path, section, key, newValue)
	section = trim(section)
	key = trim(key)
	newValue = trim(newValue)
	
    Dim line, char, asciiCode
    Dim fso1, fso2, file
    Const ForAppending = 8
    Const ForReading = 1

	Dim encoding, filePath
	encoding = GetFileEncoding(path)

   
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set file  = fso1.OpenTextFile(path, ForReading, False, encoding)
	
	' Create FileSystemObject for writing
	Set fso2 = CreateObject("Scripting.FileSystemObject")
	Set newFile = fso2.CreateTextFile(path & ".temp.ini", True, encoding) ' Overwrite=True, Unicode=True
	
	'check if key exist before writing to it
	keyRed = readFromIni(path,  section,  key)
	
	If keyRed = "" Then 
		MsgBox "Error couldn't write to ini - key wasn't found"
		Exit Function
	End If 

	
	foundKey = false
	foundSection = false
    While Not file.AtEndOfStream
        line = file.ReadLine
		
		If foundSection = false Then
            If isWantedLabel(line, section) = true Then
                foundSection = true
            End If
        Else
            wantedKey = getKeyValue(line, key)
            If wantedKey <> "" Then
                newLine = key + "=" + newValue
				newFile.WriteLine newLine
				foundKey = True

            End If
        End If
		
		if foundKey = false then 
			newFile.writeLine line
		end if 
		
		foundKey = false
    Wend
	
	file.Close
	newFile.Close

	' Delete original file
	fso1.DeleteFile path

	' Rename/move temp file to original path
	fso2.MoveFile path & ".temp.ini", path

	' Remove read-only attribute if it exists
	fso2.GetFile(path).Attributes = 0
	
	WriteToIni = foundKey
	
End Function
