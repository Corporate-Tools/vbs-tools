
' add_prefix $FOLDER $Filnmame_Substring $Prefix_to_add

folderPath        = WScript.Arguments(0)
substringToMatch  = WScript.Arguments(2)
prefixToAdd       = WScript.Arguments(1)


 Set objFSO    = CreateObject("Scripting.FileSystemObject")
 Set objFolder = objFSO.GetFolder(folderPath)

 For Each objFile In objFolder.Files
 
    If    InStr(1, objFile.Name, substringToMatch, vbTextCompare) > 0 _ 
    Then
            ' Generate the new file name with the prefix
            newFileName = prefixToAdd & objFile.Name
            ' Rename the file with the new file name
            objFile.Name = newFileName
        
            ' Print a message indicating the renaming

    End If
Next

' Clean up objects
Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing
