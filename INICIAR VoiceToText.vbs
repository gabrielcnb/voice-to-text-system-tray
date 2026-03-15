' Voice to Text - Launcher Simples

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

scriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\voice_to_text.py"

' Verificar se o script existe
If Not FSO.FileExists(scriptPath) Then
    MsgBox "Arquivo voice_to_text.py nao encontrado!", vbExclamation, "Voice to Text"
    WScript.Quit
End If

' Executar com pythonw (sem janela de console)
WshShell.CurrentDirectory = scriptDir
WshShell.Run "pythonw """ & scriptPath & """", 0, False
