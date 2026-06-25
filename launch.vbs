Dim WShell, FSO, appDir, pythonPath, port
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

appDir = FSO.GetParentFolderName(WScript.ScriptFullName)
port = 5002

WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":"" & port & """"""') do taskkill /f /pid %a", 0, True

Dim oExec, sLine
pythonPath = "python"
Set oExec = WShell.Exec("cmd /c where python")
Do While Not oExec.StdOut.AtEndOfStream
    sLine = Trim(oExec.StdOut.ReadLine())
    If InStr(sLine, "python.exe") > 0 Then
        pythonPath = sLine
        Exit Do
    End If
Loop

WShell.CurrentDirectory = appDir
WShell.Run "cmd /c set PYTHONIOENCODING=utf-8 && set PYTHONUTF8=1 && """ & pythonPath & """ app.py", 0, False

WScript.Sleep 8000
WShell.Run "http://localhost:" & port & "/"