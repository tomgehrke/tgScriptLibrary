' ===============================================
' generateWorkstationGroupReport.vbs
' -- August 14, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' Lists local groups and group membership on
' domain computers
' -----------------------------------------------
' RECORD OF REVISION:
' 08/14/2007 [tcg] Script created.
' -----------------------------------------------
' USAGE:
'
' ===============================================
Option Explicit

On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


' *****************************************************
' BASE
' *****************************************************
Dim strErrorMessage, objShell, intShellExecStatus
strErrorMessage = ""
Set objShell = CreateObject("WScript.Shell")

' *****************************************************
' PARAMETERS
' *****************************************************
Dim strDomain, strComputerNamePrefix

' *****************************************************
' DEFAULTS
' *****************************************************
strDomain = ""
strComputerNamePrefix = ""

' *****************************************************
' SET UP LOGGING
' *****************************************************
Dim strLogFolder
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Local Group Reporter\"

' *****************************************************
' READ ARGUMENTS
' *****************************************************
Dim objArguments
Set objArguments = WScript.Arguments

Dim intArgCount
intArgCount = objArguments.Count

Dim intCounter
Dim strThisArgument
Dim strThisOption
Dim strThisValue
For intCounter = 0 to intArgCount - 1
  strThisArgument = LCase( objArguments( intCounter ) )
  If inStr( strThisArgument, "=" ) > 0 Then
    strThisOption = Trim(Left( strThisArgument, InStr( strThisArgument, "=" ) - 1 ))
    strThisValue = Trim(Mid( strThisArgument, InStr( strThisArgument, "=" ) + 1 ))
  Else
    strThisValue = Trim(strThisArgument)
  End If

  Select Case strThisOption
  Case "domain"
    strDomain = UCase(strThisValue)
  Case "prefix"
    strComputerNamePrefix = LCase(strThisValue)
  Case Else
    strDomain = LCase(strThisValue)
  End Select
Next

' ===============================================

Dim objDomain, objFileSystem, objLogFile, strLogFile, objItem, strComputer

strLogFile = strLogFolder & "Local Group Memberships for " & strDomain & " - " & Replace( Replace( CStr( Now()), "/", "_"), ":", "_") & ".csv"

Set objDomain = GetObject("WinNT://" & strDomain)
Set objFileSystem = CreateObject("Scripting.FileSystemObject")

If Not objFileSystem.FolderExists( strLogFolder ) Then
  objFileSystem.CreateFolder( strLogFolder )
End If

Set objLogFile = objFileSystem.OpenTextFile(strLogFile, ForAppending, True)

objLogFile.WriteLine( """Computer"",""Group"",""Account""")

Dim colGroups, objGroup, objUser, bolNoMembers

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strComputer = UCase(objItem.Name)

    If Left(LCase(strComputer), LEN(strComputerNamePrefix)) = strComputerNamePrefix Then
      Set colGroups = GetObject("WinNT://" & strComputer)
      colGroups.Filter = Array("group")

      For Each objGroup In colGroups
        bolNoMembers = True
        For Each objUser in objGroup.Members
          objLogFile.WriteLine( """" & strComputer & """,""" & objGroup.Name & """,""" & objUser.ADSPath & """" )
          bolNoMembers = False
        Next
        If bolNoMembers Then
          objLogFile.WriteLine( """" & strComputer & """,""" & objGroup.Name & """,""""" )
        End If
      Next
    End If

  End If
Next

objLogFile.Close

'======================================================
' EOF: generateWorkstationGroupReport.vbs
'======================================================
