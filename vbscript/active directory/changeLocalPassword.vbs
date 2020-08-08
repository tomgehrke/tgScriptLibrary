' ************************************************ '
' Change Local Account Password Script
'
' ChangeLocalPassword.vbs
'
' by Thomas Gehrke
' ************************************************ '
Option Explicit

ON ERROR RESUME NEXT

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objFileSystem
Dim objWshNetwork
Dim objWshShell
Dim objDomain
Dim objScriptExec
Dim fileLog
Dim strLogFileName
Dim objItem
Dim strWorkstation
Dim strTargetFileName
Dim objFile
Dim strTargetUser
Dim strNewPassword

' ================================================ '
' CUSTOMIZATION SECTION
' ================================================ '

strTargetDomain = "Domain"
strTargetUser = "Administrator"
strNewPassword = "n3wp455w0rd"
strTargetPrefix = ""

' ================================================ '
' END OF CUSTOMIZATION SECTION
' ================================================ '

strLogFileName = "ChangeLocalPassword.log"

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objWshNetwork = WScript.CreateObject("WScript.Network")
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objDomain = GetObject("WinNT://" + strTargetDomain)
Set fileLog = objFileSystem.OpenTextFile(strLogFileName, ForAppending, True)

Dim User

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strComputer = UCase(objItem.Name)

    If Left(LCase(strComputer), LEN(strTargetPrefix)) = strTargetPrefix Then
      fileLog.WriteLine("===============================================" )
      fileLog.WriteLine( strComputer )

      Set User = GetObject("WinNT://" & strComputer & "/" & strTargetUser)

      User.SetPassword(strNewPassword)
      If err.number = 0 Then
        fileLog.WriteLine("- Changed password for " & strTargetUser )

        If User.AccountDisabled = True Then
          User.AccountDisabled = False
          User.SetInfo
          If err.number = 0 Then
            fileLog.WriteLine("- Enabled account" )
          Else
            fileLog.WriteLine("- ERROR: " & err.Description)
            Err.Clear
          End If
        End If
      Else
        fileLog.WriteLine("- ERROR: " & err.Description)
        Err.Clear
      End If

      Set User = Nothing
    End If
  End If
Next

fileLog.WriteLine(chr(13) & chr(10) & "*** JOB COMPLETED " & Now() & chr(13) & chr(10) )
fileLog.Close

