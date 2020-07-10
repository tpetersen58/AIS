Option Explicit
Const LAUNCHER_FILE = "C:\Users\Taylor Petersen\Desktop\test.bat"
Dim objWSHShell
Dim sLaunchFile
Dim FSO

Set FSO = CreateObject("Scripting.FileSystemObject")
If Not (FSO.FolderExists("c:\temp")) Then
FSO.CreateFolder "c:\temp"
End If

Set objWSHShell = CreateObject("Wscript.Shell")
objWSHShell.run chr(34) & LAUNCHER_FILE & chr(34), 0, True