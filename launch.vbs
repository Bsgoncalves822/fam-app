Dim WShell, FSO, http, port, appDir, configPath, pythonPath
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

appDir = FSO.GetParentFolderName(WScript.ScriptFullName)

port = 5002
configPath = appDir & "\config.json"
If FSO.FileExists(configPath) Then
    Dim cfgFile, cfgText, re, matches
    Set cfgFile = FSO.OpenTextFile(configPath, 1, False, -1)
    cfgText = cfgFile.ReadAll
    cfgFile.Close
    Set re = New RegExp
    re.Pattern = """port""\s*:\s*(\d+)"
    re.Global = False
    Set matches = re.Execute(cfgText)
    If matches.Count > 0 Then port = CInt(matches(0).SubMatches(0))
End If

Dim baseUrl
baseUrl = "http://localhost:" & port & "/"

WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":"" & port & """"""') do taskkill /f /pid %a", 0, True

' Find python dynamically
Dim oExec, sLine
pythonPath = ""
Set oExec = WShell.Exec("cmd /c where python")
Do While Not oExec.StdOut.AtEndOfStream
    sLine = Trim(oExec.StdOut.ReadLine())
    If InStr(sLine, "python.exe") > 0 And InStr(sLine, "WindowsApps") = 0 Then
        pythonPath = sLine
        Exit Do
    End If
Loop
If pythonPath = "" Then
    Set oExec = WShell.Exec("cmd /c where python")
    Do While Not oExec.StdOut.AtEndOfStream
        sLine = Trim(oExec.StdOut.ReadLine())
        If InStr(sLine, "python.exe") > 0 Then
            pythonPath = sLine
            Exit Do
        End If
    Loop
End If
If pythonPath = "" Then pythonPath = "python"

WShell.CurrentDirectory = appDir
WShell.Run "cmd /c set PYTHONIOENCODING=utf-8 && set PYTHONUTF8=1 && """ & pythonPath & """ app.py", 0, False

WScript.Sleep 3000
Dim attempts
attempts = 0
Do While attempts < 30
    WScript.Sleep 1500
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", baseUrl, False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        WShell.Run baseUrl
        WScript.Quit
    End If
    On Error GoTo 0
    attempts = attempts + 1
Loop

MsgBox "Erro ao iniciar FAM App." & Chr(13) & Chr(10) & "Verifique se o Python esta instalado e tente novamente.", 16, "FAM App"