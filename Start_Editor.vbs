Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = scriptDir

portablePython = scriptDir & "\runtime\Scripts\pythonw.exe"
localPython = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe"

If fso.FileExists(portablePython) Then
    WshShell.Run """" & portablePython & """ web_editor.py", 0, False
ElseIf fso.FileExists(localPython) Then
    WshShell.Run """" & localPython & """ web_editor.py", 0, False
Else
    WshShell.Run "pythonw web_editor.py", 0, False
End If
