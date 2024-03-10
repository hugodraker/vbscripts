'LICENCE: PUBLIC DOMAIN
'NO WARRANTY
'mostly tested under 32bit GS9.5.4, some testing with newer GS versions
Dim arg, gswinexe, cmdline, logging, gspath, papersize, pageoffset, ticks, waittime, oldticks, writedelay, temppath, savepath, temppdfpath, tempmovedpath, temptagpath, tickspath, temptxtpath, pname, scanid
waittime = 15
pdfversion = 1.2 'may crash if PDF Version is too low
writedelay = 1500 '1.5 seconds increase if corrupted pdfs occur
temppath ="C:\temp"
savepath ="C:\Patient Reports"
pname="none"
scanid="Report"
tickspath =temppath & "\ticks"
temppdfpath =temppath & "\pdf.pdf"
tempmovedpath =temppath & "\pdfmoved.pdf"
temptagpath =temppath & "\pdftag.pdf"
temptxtpath =temppath & "\pdf.txt"
papersize = "letter"
pageoffset = "21 10" 'x,y
logging = False
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")

if WScript.Arguments.Named.Exists("elevated")=False AND IsSystemUser=False Then 
 '   CreateObject ("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """"&" /elevated", "","runas", 1
 '   WScript.Quit
End If

'On error resume next 'comment out to debug
gspath = ReadInstallPath("ghostscript")
gswinexe = GetLastWildFile(gspath & "\bin" , "c.exe",0)  'gswin32c.exe gswin64c.exe
'gswinexe = """" & gspath & "\bin\" & GetLastWildFile(gspath & "\bin" , "c.exe",0) &"""" 'use full path

If gspath=False then
    Msgbox "Can't find ghostscript path from installed programs, quitting"
    WScript.Quit
End If

CheckCreateFolder(temppath)

Dim windowStyle
Dim waitOnReturn
windowStyle = 1
waitOnReturn = True

CheckCreateFolder(savepath)

ticks = DateDiff("s", "01/01/1900 00:00:00", Now())
'ticks = 1919040953 'comment out to never delete old .ps files

Dim colVolEnvVars, fso
Dim sPath, your_path

set fso = CreateObject("Scripting.FileSystemObject")
'Set myLog = fso.OpenTextFile(temppath & "\" & "pdf.log", 3)

'if WScript.Arguments.Count = 0 then 
'arg =WScript.Arguments(0)
arg=GetLastWildFile(temppath, ".ps",1)

If NOT FileExists(arg) Then
    MsgBox  "This script will merge all print jobs within " & waittime &" seconds, and output a PDF to " & savepath & ", parsing out a patient name and ID from the PDF with Ghostscript." & vbCrLf &"Should be called with at least 1 postscript file saved in " & temppath & vbCrLf & "typically in Multi File Port Monitor, or Ghostgum Redmon" & vbCrLf & "filename pattern: %i.ps" & vbCrLf & "output folder: " & temppath & vbCrLf & "cscript.exe " & WScript.ScriptFullName & vbCrLf & "Please add a Color Postscript Printer, then set the port to port monitor software" & vbCrLf  & vbCrLf &"SORRY Doesn't work on Windows 11 or 10",64,"Usage"
    WScript.Quit
end if

If FileExists(tickspath) Then
    if DeleteTicks() > 0 then
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.MoveFile arg, temppath & "\" & "0001.p"
        DeleteAFile(tickspath)
        DeleteAFile(temppath & "\*.ps")
        DeleteAFile(temppath & "\pdf*.txt")
        fs.MoveFile temppath & "\" & "0001.p",temppath & "\" & "0001.ps"
        'arg = temppath & "\" & "0001.ps"
    end if    
End If

WScript.Sleep(500)
WriteTicks()

'10.02.1 ghostscript likes this:
'oShell.run gswinexe & " -dNOPAUSE -dBATCH -q -dSAFER -dPDFSETTINGS=/printer -sOutputFile=""" & temppdfpath & """ " & GetWildFile(temppath, ".ps",1) & " -sDEVICE=pdfwrite",windowStyle, waitOnReturn
'9.54.0 ghostscript likes this:
''oShell.run gswinexe & " -dDisplayFormat=198788 -dDisplayResolution=96 -sDEVICE=pdfwrite -c ""<</NeverEmbed []>> setdistillerparams"" ""-o " & temppdfpath & """" & " -dNOPAUSE -dBATCH -q -sPAPERSIZE=letter -dCompatibilityLevel=" & pdfversion & " " & GetWildFile(temppath, ".ps",1),windowStyle, waitOnReturn
'cmdline = "echo 1"
cmdline = "cd """ & gspath & "\" & "bin"""
cmdline = cmdline & " & " & gswinexe & " -dDisplayFormat=198788 -dDisplayResolution=96 -sDEVICE=pdfwrite ""-sOutputFile=" & temppdfpath & """" & " -c ""<</NeverEmbed []>> setdistillerparams"" -dNOPAUSE -dBATCH -q -sPAPERSIZE=letter -dCompatibilityLevel=" & pdfversion & " " & GetWildFile(temppath, ".ps",1) 
cmdline = cmdline & " & timeout 0 " & " & " & gswinexe & " -dNOPAUSE -dBATCH -q -dSAFER -sDEVICE=txtwrite -dTextFormat=3 -sOutputFile=""" & temptxtpath & """ " & """" & temppdfpath & """"
cmdline = "cmd /K """ & cmdline  &  " & exit"""
oShell.run cmdline
if logging then Call oShell.LogEvent(1,cmdline)
WScript.Sleep(writedelay)

If FileExists(temptxtpath) Then
    set re = new RegExp
    re.Pattern="in"
    re.Global=True
    Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(temptxtpath,1)
    Dim strLine
    do while not objFileToRead.AtEndOfStream
        strLine = objFileToRead.ReadLine()
        If InStr(strLine,"Name:") > 0 Then
            re.Pattern = "Sex.*|Name:|\s"
            pname =re.Replace(strLine,"")
            pname =Replace(pname,","," ")
        end if

        If InStr(strLine,"Patient ID") = False Then
            If InStr(strLine,"ID:") > 0 Then
                re.Pattern = "[a-zA-Z][0-9]{8}"
                Set Matches = re.Execute(strLine)
                For Each Match in Matches
                    scanid=Match.Value
                    Exit For
                Next
                'MsgBox scanid
                Exit Do
             end if
         end if
    loop
    objFileToRead.Close
    Set objFileToRead = Nothing
end if

'cmdline = "echo 1"
cmdline = "cd """ & gspath & "\" & "bin"""
'add margins to the page, as some sofware doesn't
cmdline = cmdline & " & " & gswinexe & " -sDEVICE=pdfwrite -sOutputFile=""" & tempmovedpath & """ " & " -dNOPAUSE -dBATCH -q -sPAPERSIZE=" & papersize & " -dCompatibilityLevel=" & pdfversion & " -c ""<</PageOffset [" & pageoffset & "]>> setpagedevice"" -f " & """" & temppdfpath & """" & ""
'add subject and keyword tags
'oShell.run gswinexe & " -sDEVICE=pdfwrite -sOutputFile=""" & temppdfpath & """ " & " -dNOPAUSE -dBATCH -q -sPAPERSIZE=" & papersize & " -dCompatibilityLevel=" & pdfversion & " -c ""[ /Title (TitleHere) /Author (AuthorHere) /Subject (" & pname & ") /Creator (CreatorHere) /Producer (AuthorHere) /Keywords (" & scanid & ") /DOCINFO pdfmark"" -f " & """" & tempmovedpath & """", windowStyle, waitOnReturn
cmdline = cmdline & " & " & gswinexe & " -sDEVICE=pdfwrite -sOutputFile=""" & temppdfpath & """ " & " -dNOPAUSE -dBATCH -q -sPAPERSIZE=" & papersize & " -dCompatibilityLevel=" & pdfversion & " -c ""[ /Title (TitleHere) /Author (AuthorHere) /Subject (" & pname & ") /Creator (CreatorHere) /Producer (AuthorHere) /Keywords (" & scanid & ") /DOCINFO pdfmark"" -f " & """" & tempmovedpath & """" & ""
cmdline = cmdline & " & copy """ & temppdfpath & """ " & """" & savepath & "\" & pname & "_" & scanid & ".pdf"""
'cmd /K "echo "1" & echo "2" & echo "3" & exit" '%systemroot%\system32 must be in the path
cmdline = "cmd /K """ & cmdline  &  " & exit"""
oShell.run cmdline
if logging then Call oShell.LogEvent(1,cmdline)

'WScript.Sleep(100)
'if fso.FileExists(temppdfpath) then fso.CopyFile temppdfpath, savepath & "\" & pname & "_" & scanid & ".pdf", TRUE
Set fs=nothing


Set oShell = Nothing
Set fso = Nothing
'myLog.Close

Function FileExists(FilePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        FileExists=CBool(1)
    Else
        FileExists=CBool(0)
    End If
End Function

Function WriteTicks()
    Set fst=createobject("Scripting.FileSystemObject")
    Set qfile=fst.OpenTextFile(tickspath,2,True)
    qfile.Write ticks
    qfile.Close
    Set qfile=nothing
    Set fst=nothing
End Function

Function DeleteTicks()
    Set fst=createobject("Scripting.FileSystemObject")
    Set qfile=fst.OpenTextFile(tickspath,1,True)
    Do while qfile.AtEndOfStream <> true
        oldticks=qfile.ReadLine
        if ticks-int(oldticks) >waittime then DeleteTicks=1
    Loop
    qfile.Close
    Set qfile=nothing
    Set fst=nothing
End Function

Sub DeleteAFile(filespec)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On error resume next
    fso.DeleteFile(filespec)
End Sub


Function GetWildFile(strFolder, strWild,returnFolder) 
Dim objfs 
Dim objFolder
dim objFiles
Dim objFile
Dim strDesc
Dim strspace

strspace =""
Set objfs = CreateObject("Scripting.FileSystemObject")
On error resume next ' Intercept No Folder
    Set objFolder = objFS.GetFolder(strFolder)
    if Err.Number <> 0 then strDesc = Err.Description
    On error goto 0
    If Len(strDesc) = 0 then
        Set objFiles = objFolder.Files
        For Each objFile in ObjFiles
            if instr(1,objFile.Name, strWild, 1) > 0 then 
                Select Case returnFolder
                    Case 1
                        GetWildFile = GetWildFile & strspace & """" & strFolder & "\" & objFile.Name & """"
                    Case Else
                        GetWildFile = GetWildFile & strspace & """"  & objFile.Name & """"
                    End Select
            End if
            strspace =" "
        Next
    End if 

Set objfs = nothing
End Function

Function GetLastWildFile(strFolder, strWild,returnFolder) 
Dim objfs 
Dim objFolder
dim objFiles
Dim objFile
Dim strDesc
Set objfs = CreateObject("Scripting.FileSystemObject")
On error resume next ' Intercept No Folder
    Set objFolder = objFS.GetFolder(strFolder)
    if Err.Number <> 0 then strDesc = Err.Description
    On error goto 0
    If Len(strDesc) = 0 then
        Set objFiles = objFolder.Files
        For Each objFile in ObjFiles
            if instr(1,objFile.Name, strWild, 1) > 0 then 
                Select Case returnFolder
                    Case 1
                        GetLastWildFile = strFolder & "\" & objFile.Name
                    Case Else
                        GetLastWildFile = objFile.Name
                    End Select
            End if
        Next
    End if 

Set objfs = nothing
End Function

Public Function CheckCreateFolder(path)
    path=path & "\"
    Dim TempPath
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    pos = 0
    While pos < Len(path)
        pos = InStr(pos + 1, path, "\")
        TempPath = Left(path, pos)
        If Not (fso.FolderExists(TempPath)) Then
            fso.CreateFolder TempPath
        End If
    Wend
     set fso = nothing
End Function

Function ReadInstallPath(strKeyName)
    Dim oReg, oFSO 
    Dim UninstallString, ProductCode
    Dim strComputer, colItems, objWMIService, objItem
    Dim strKeyPath(1), subkey, arrSubKeys
    strComputer = "." 
    'InputBox(UninstallString)

    Const HKEY_LOCAL_MACHINE = &H80000002
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    strKeyPath(0) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    strKeyPath(1) = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

    For i=0 to 1
         if oReg.EnumKey(HKEY_LOCAL_MACHINE, strKeyPath(i), arrSubKeys) = False then
         'if err.number <> 0 Then
            For Each subkey In arrSubKeys 
               IF instr(1,subkey,strKeyName,1) Then
               str = ReadfromRegistry("HKLM" & "\" & strKeyPath(i) & "\" & subkey & "\" &"UninstallString","1")
               With CreateObject("scripting.filesystemobject")
                   str = Replace(str,"""","")
                   wFolder = .GetParentFolderName(str)
                   ReadInstallPath=wFolder
                 End With
               End If
            Next
        End If
    Next

    Set oReg = Nothing
End Function

Function ReadFromRegistry(strRegistryKey, strDefault)
    Dim value
    On Error Resume Next
    value = oShell.RegRead( strRegistryKey )
    If err.number <> 0 Then
        readFromRegistry = "" 'strDefault
    Else
        readFromRegistry = value
    End If
End Function

Function CreateTempFile 
   Dim tfolder, tname, tfile
   Const TemporaryFolder = 2
   CreateTempFile = fso.GetSpecialFolder(TemporaryFolder)
   'tname = fso.GetTempName   
   'Set tfile = tfolder.CreateTextFile(tname)
   'Set CreateTempFile = tfile
End Function