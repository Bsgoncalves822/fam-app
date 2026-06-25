' FAM App Launcher
' Kills any existing instance, starts Flask, waits for ready, opens browser

Dim WShell, FSO, http, port, appDir, configPath, pythonCmd
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Derive install directory from this script's own location (no hardcoded path)
appDir = FSO.GetParentFolderName(WScript.ScriptFullName)

' Set encoding environment variables
WShell.Environment("Process")("PYTHONIOENCODING") = "utf-8"
WShell.Environment("Process")("PYTHONUTF8") = "1"

' Default port; override from config.json if present
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
    If matches.Count > 0 Then
        port = CInt(matches(0).SubMatches(0))
    End If
End If

Dim baseUrl
baseUrl = "http://localhost:" & port & "/"

' Kill any existing instance on the configured port
WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":"" & port & """"""') do taskkill /f /pid %a", 0, True

' Find Python â€” try py launcher first (Windows standard), then python
pythonCmd = "python"
Dim testResult
testResult = WShell.Run("cmd /c py --version >nul 2>&1", 0, True)
If testResult = 0 Then pythonCmd = "py"

' Start Flask silently from the app directory
WShell.CurrentDirectory = appDir
WShell.Run "cmd /c set PYTHONIOENCODING=utf-8 && set PYTHONUTF8=1 && " & pythonCmd & " app.py", 0, False

' Wait for Flask to be ready (health check loop)
WScript.Sleep 5000
Dim attempts
attempts = 0
Do While attempts < 15
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

MsgBox "Erro ao iniciar FAM App." & Chr(13) & Chr(10) & _
       "Verifique se o Python esta instalado e tente novamente.", 16, "FAM App"
