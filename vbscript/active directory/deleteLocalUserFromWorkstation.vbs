' ===============================================
' deleteLocalUserFromWorkstation.vbs
' -- August 14, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' Deletes local accounts on workstations
' -----------------------------------------------
' RECORD OF REVISION:
' 08/14/2007 [tcg] Script rewritten.
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
Dim strDomain, strComputerNamePrefix, strLocalAccount, bolVerbose

' *****************************************************
' DEFAULTS
' *****************************************************
strDomain = ""
strComputerNamePrefix = ""
strLocalAccount = ""
bolVerbose = False

' *****************************************************
' SET UP LOGGING
' *****************************************************
Dim strLogFolder
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Delete Local Account Logs\"

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
  Case "account"
    strLocalAccount = LCase(strThisValue)
  Case "verbose"
    If LCase(strThisValue) = "on" or LCase(strThisValue) = "1" or LCase(strThisValue) = "yes" or LCase(strThisValue) = "true" Then
      bolVerbose = True
    End If
  Case Else
    strLocalAccount = LCase(strThisValue)
  End Select
Next

' ===============================================

Dim objDomain, objFileSystem, objLogFile, strLogFile, objItem, strComputer

strLogFile = strLogFolder & "Delete '" & strLocalAccount & "' from " & strDomain & ".csv"

Set objDomain = GetObject("WinNT://" & strDomain )
Set objFileSystem = CreateObject("Scripting.FileSystemObject")

If Not objFileSystem.FolderExists( strLogFolder ) Then
  objFileSystem.CreateFolder( strLogFolder )
End If

Set objLogFile = objFileSystem.OpenTextFile(strLogFile, ForAppending, True)

' objLogFile.WriteLine( """Computer"",""Account"",""Result"",""Time""")

Dim Computer

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strComputer = UCase(objItem.Name)

    If Left(LCase(strComputer), LEN(strComputerNamePrefix)) = strComputerNamePrefix Then
      Set Computer = GetObject("WinNT://" & strComputer & ",computer")
      Computer.Delete "user", strLocalAccount

      If err.number = 0 Then
        objLogFile.WriteLine( """" & strComputer & """,""" & strLocalAccount & """,""deleted"",""" & Now() & """" )
      Else
        If bolVerbose Then
          objLogFile.WriteLine( """" & strComputer & """,""" & strLocalAccount & """,""ERROR: " & Err.Number & " - " & Replace( Err.Description, vbCrLf, "") & """,""" & Now() & """" )
        End If
        Err.Clear
      End If
    End If

    Set Computer = Nothing
  End If
Next

objLogFile.Close

'======================================================
' EOF: deleteLocalUserFromWorkstation.vbs
'======================================================