' FAM App Launcher
' Kills any existing instance, starts Flask, waits for ready, opens browser

Dim WShell, FSO, http, port, appDir, configPath
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Derive install directory from this script's own location (no hardcoded path)
appDir = FSO.GetParentFolderName(WScript.ScriptFullName)

' Default port; override from config.json's "port" key if present
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
WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":" & port & """') do taskkill /f /pid %a", 0, True

' Start Flask silently from the app directory
WShell.CurrentDirectory = appDir
WShell.Run "cmd /c python app.py", 0, False

' Wait for Flask to be ready (health check loop)
WScript.Sleep 2000
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

' If we get here Flask didn't start - show error
MsgBox "Erro ao iniciar FAM App. Verifique se o Python esta instalado corretamente.", 16, "FAM App"
