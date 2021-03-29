Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c launcherForPowerShellScript.bat"
oShell.Run strArgs, 0, false