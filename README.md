# vbs iniHandler

## About  
Robust VBScript functions to **read from** and **overwrite values** in `.ini` files, supporting any kind of file encoding.  
Most of the code is a collection of snippets that have been modified and assembled together. :)

## Requirements  
- VBScript (VBS) only

## Project Structure  
This project includes two files:

1. **`getFileEncoding.vbs`**  
   Determines the file's encoding type.  
   _(Logic adapted from [Rob van der Woude’s script](https://www.robvanderwoude.com/vbstech_files_encoding.php) and turned into a function that returns the correct parameter for `OpenTextFile()`.)_

2. **`iniHandler.vbs`**  
   Provides two functions:
   - `WriteToIni(path, section, key, newValue)`
   - `readFromIni(path, section, key)`

## How to Use  
1. Copy both files into your working directory.

2. In your script, use the following to include the functions:

    ```vbscript
    '--------------------------
    ' Import other functions
    Sub Include(strFile)
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objTextFile = objFSO.OpenTextFile(strFile, 1)
        ExecuteGlobal objTextFile.ReadAll
        objTextFile.Close
        Set objFSO = Nothing
        Set objTextFile = Nothing
    End Sub

    ' Include handlers
    Include("iniHandler.vbs")
    '--------------------------
    ```

3. Example usage:

    ```vbscript
    value = readFromIni("file.ini", "section", "key")
    WriteToIni("file.ini", "section", "key", "newValue")
    ```

## Notes  
- Currently does **not** support appending new sections or keys. This may be added in the future.

Enjoy! 😃  

## References  
- https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/  
- https://www.robvanderwoude.com/vbstech_files_encoding.php  
- https://www.robvanderwoude.com/vbstech_files_ini.php
"# vbs-.ini-file-Handler" 
