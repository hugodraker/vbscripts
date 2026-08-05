Option Explicit

Dim objShell, objExec, strOutput, isSafeBoot, msg, userChoice

Set objShell = CreateObject("WScript.Shell")

' 1. Self-elevate the script to run as Administrator if not already elevated
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

' 2. Execute bcdedit to inspect the current boot entry configuration
Set objExec = objShell.Exec("bcdedit /enum {current}")
strOutput = objExec.StdOut.ReadAll()

' 3. Determine current mode (returns true if safeboot parameter is found)
isSafeBoot = (InStr(1, strOutput, "safeboot", vbTextCompare) > 0)

' 4. Toggle boot configuration based on current state
If isSafeBoot Then
    ' --- REVERT TO NORMAL MODE ---
    objShell.Run "bcdedit /deletevalue {current} safeboot", 0, True
    objShell.Run "bcdedit /deletevalue {current} safebootalternateshell", 0, True
    
    msg = "Boot configuration set to: NORMAL MODE" & vbCrLf & vbCrLf & "Would you like to restart now?"
Else
    ' --- SET TO SAFE MODE (Network + Alternate Shell) ---
    objShell.Run "bcdedit /set {current} safeboot network", 0, True
    objShell.Run "bcdedit /set {current} safebootalternateshell yes", 0, True
    
    msg = "Boot configuration set to: SAFE MODE (Network + Command Prompt)" & vbCrLf & vbCrLf & "Would you like to restart now?"
End If

' 5. Prompt to reboot immediately
userChoice = MsgBox(msg, vbYesNo + vbQuestion, "Boot Mode Toggled")

If userChoice = vbYes Then
    objShell.Run "shutdown /r /t 0", 0, False
End If