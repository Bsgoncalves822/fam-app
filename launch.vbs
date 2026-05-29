' FAM App Launcher
' Kills any existing instance, starts Flask, waits for ready, opens browser

Dim WShell, http, port, appDir
Set WShell = CreateObject("WScript.Shell")

port = 5002
appDir = "C:\fam-app"

' Kill any existing instance on port 5002
WShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| find "":5002""') do taskkill /f /pid %a", 0, True

' Start Flask silently
WShell.Run "cmd /c cd /d " & appDir & " && python app.py", 0, False

' Wait for Flask to be ready (health check loop)
WScript.Sleep 2000
Dim attempts
attempts = 0
Do While attempts < 15
    WScript.Sleep 1500
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "http://localhost:5002/", False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        WShell.Run "http://localhost:5002"
        WScript.Quit
    End If
    On Error GoTo 0
    attempts = attempts + 1
Loop

' If we get here Flask didn't start - show error
MsgBox "Erro ao iniciar FAM App. Verifique se o Python está instalado corretamente.", 16, "FAM App"
