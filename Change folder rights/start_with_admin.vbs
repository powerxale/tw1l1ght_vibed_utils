Set Shell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Gets the path to the main batch script (assumes it is in the same directory)
Dim batchFilePath
batchFilePath = FSO.GetParentFolderName(WScript.ScriptFullName) & "\edit_permissions_folder.bat"

' Collects the passed arguments (the folder path)
Dim args
If WScript.Arguments.Count > 0 Then
    For i = 0 To WScript.Arguments.Count - 1
        ' Encapsulates arguments with spaces in quotation marks
        If InStr(WScript.Arguments(i), " ") > 0 Then
            args = args & Chr(34) & WScript.Arguments(i) & Chr(34) & " "
        Else
            args = args & WScript.Arguments(i) & " "
        End If
    Next
    ' Removes excess end space
    args = RTrim(args)
End If

' Start the batch script with elevated privileges (as administrator)
' This will pop up the User Account Control (UAC) prompt if needed.
Shell.ShellExecute batchFilePath, args, "", "runas", 1

' Exit VBScript script
WScript.Quit