Dim strFilter, strComputer, objWMIService, colServices, objService
Dim errReturn

' Mention admin requirement in the output
WScript.Echo "Note: Deleting services requires this script to be run as an Administrator."

' Check for command line arguments
If WScript.Arguments.Count > 0 Then
    strFilter = LCase(WScript.Arguments(0))
    WScript.Echo "Filter specified: '" & strFilter & "'"
    WScript.Echo "WARNING: Any service containing this text in its Name or Display Name will be DELETED."
    WScript.Echo "Press CTRL+C immediately to abort if this is a mistake..."
    WScript.Sleep 5000 ' Give the user 5 seconds to abort
Else
    strFilter = ""
    WScript.Echo "No filter specified. Listing all services..."
    WScript.Echo "----------------------------------------------------"
End If

strComputer = "."
' Connect to WMI
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService In colServices
    If strFilter = "" Then
        ' No argument provided: Just list the services one by one as command output
        WScript.Echo "Name: " & objService.Name & " | Display Name: " & objService.DisplayName
    Else
        ' Argument provided: Look for matches in the Name or Display Name
        If InStr(LCase(objService.Name), strFilter) > 0 Or InStr(LCase(objService.DisplayName), strFilter) > 0 Then
            WScript.Echo vbCrLf & "MATCH FOUND: " & objService.Name & " (" & objService.DisplayName & ")"
            WScript.Echo "Attempting to delete..."
            
            ' Attempt to stop the service first if it is running
            If objService.State = "Running" Then
                WScript.Echo "  -> Service is running. Attempting to stop it..."
                objService.StopService()
                WScript.Sleep 2000 ' Wait 2 seconds for it to stop
            End If

            ' Delete the service
            errReturn = objService.Delete()
            
            If errReturn = 0 Then
                WScript.Echo "  -> SUCCESS: Service deleted."
            Else
                WScript.Echo "  -> FAILED: Could not delete service. Error code: " & errReturn
            End If
        End If
    End If
Next

WScript.Echo vbCrLf & "Script complete."