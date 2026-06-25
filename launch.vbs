Dim WShell, FSO, appDir, port
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

appDir = FSO.GetParentFolderName(WScript.ScriptFullName)
port = 5002

WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":"" & port & """"""') do taskkill /f /pid %a", 0, True

WShell.CurrentDirectory = appDir
WShell.Run "cmd /c set PYTHONIOENCODING=utf-8 && py app.py", 0, False

WScript.Sleep 8000
WShell.Run "http://localhost:" & port & "/"