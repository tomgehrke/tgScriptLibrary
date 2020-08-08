' Workstation Administrator Report Script
Option Explicit

On Error Resume Next

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
Dim strDomain
Dim strWorkstation
Dim strTargetFileName
Dim objFile
Dim datTargetFileThreshold
Dim intDateDifference

strDomain = "DOMAIN"
strLogFileName = "YourDomainAdministratorReport.log"

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objWshNetwork = WScript.CreateObject("WScript.Network")
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objDomain = GetObject("WinNT://" & strDomain)
Set fileLog = objFileSystem.OpenTextFile(strLogFileName, ForAppending, True)

Dim Group
Dim Member

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strWorkstation = UCase(objItem.Name)
    
    fileLog.WriteLine("===============================================" )
    fileLog.WriteLine( strWorkstation )
    Set Group = GetObject("WinNT://" & strWorkstation & "/Administrators")
    If err.number = 0 Then
      fileLog.WriteLine("-----------------------------------------------" )
      fileLog.WriteLine("Administrators" )
      fileLog.WriteLine("-----------------------------------------------" )
      For Each Member in Group.Members
        If Member.Class = "User" And LCase(Member.Name) <> "administrator" And LCase(Member.Name) <> "usec-admin" Then
          fileLog.WriteLine("- USER: " & Member.Name & " (" & Member.FullName & ") Disabled=" & Member.AccountDisabled )
        End If
        If Member.Class = "Group" And LCase(Member.Name) <> "domain admins" And LCase(Member.Name) <> "interactive" Then
          fileLog.WriteLine("- GROUP: " & Member.Name )
        End If
      Next
    Else
      Err.Clear
    End If
    Set Group = Nothing
    Set Group = GetObject("WinNT://" & strWorkstation & "/Power Users")
    If err.number = 0 Then
      fileLog.WriteLine("-----------------------------------------------" )
      fileLog.WriteLine("Power Users" )
      fileLog.WriteLine("-----------------------------------------------" )
      For Each Member in Group.Members
        If Member.Class = "User" Then
          fileLog.WriteLine("- USER: " & Member.Name & " (" & Member.FullName & ") Disabled=" & Member.AccountDisabled )
        End If
        If Member.Class = "Group" And LCase(Member.Name) <> "interactive" Then
          fileLog.WriteLine("- GROUP: " & Member.Name )
        End If
      Next
    Else
      err.Clear
    End If
    Set Group = Nothing
    fileLog.WriteLine( chr(13) & chr(10))
  End If
Next

fileLog.WriteLine(chr(13) & chr(10) & "*** JOB COMPLETED " & Now() & chr(13) & chr(10) )
fileLog.Close
