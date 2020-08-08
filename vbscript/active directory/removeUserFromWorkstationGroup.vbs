' ===============================================
' removeUserFromWorkstationGroup.vbs
' -- August 15, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' Removes account or group from local groups on
' domain workstations
' -----------------------------------------------
' RECORD OF REVISION:
' 08/15/2007 [tcg] Script created.
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
Dim strDomain, strComputerNamePrefix, strTargetAccount, strTargetGroup, bolVerbose

' *****************************************************
' DEFAULTS
' *****************************************************
strDomain = ""
strComputerNamePrefix = ""
strTargetAccount = ""
strTargetGroup = ""
bolVerbose = False

' *****************************************************
' SET UP LOGGING
' *****************************************************
Dim strLogFolder
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Remove User from Local Group Logs\"

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
    strTargetAccount = Replace(LCase(strThisValue),"\","/")
  Case "group"
    strTargetGroup = LCase(strThisValue)
  Case "verbose"
    If LCase(strThisValue) = "on" or LCase(strThisValue) = "1" or LCase(strThisValue) = "yes" or LCase(strThisValue) = "true" Then
      bolVerbose = True
    End If
  Case Else
    strTargetAccount = LCase(strThisValue)
  End Select
Next

' ===============================================

Dim objDomain, objFileSystem, objLogFile, strLogFile, objItem, strComputer

strLogFile = strLogFolder & "Remove '[" & Replace(strTargetAccount,"/","]") & "' from '" & strTargetGroup & "' in " & strDomain & ".csv"

Set objDomain = GetObject("WinNT://" & strDomain )
Set objFileSystem = CreateObject("Scripting.FileSystemObject")

If Not objFileSystem.FolderExists( strLogFolder ) Then
  objFileSystem.CreateFolder( strLogFolder )
End If

Set objLogFile = objFileSystem.OpenTextFile(strLogFile, ForAppending, True)

' objLogFile.WriteLine( """Computer"",""Group"",""Account"",""Result"",""Time""")

Dim objTargetGroup, objTargetAccount

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strComputer = UCase(objItem.Name)

    If Left(LCase(strComputer), LEN(strComputerNamePrefix)) = strComputerNamePrefix Then
      Set objTargetGroup = GetObject("WinNT://" & strComputer & "/" & strTargetGroup)

      If err.number = 0 Then
        Set objTargetAccount = GetObject("WinNT://" & strTargetAccount)

        If err.number = 0 Then
          objTargetGroup.Remove(objTargetAccount.ADsPath)

          If err.number = 0 Then
            objLogFile.WriteLine( """" & strComputer & """,""" & strTargetGroup & """,""" & strTargetAccount & """,""removed"",""" & Now() & """" )
          End If
        End If
      End If

      If Err.Number <> 0 Then
        If bolVerbose Then
          objLogFile.WriteLine( """" & strComputer & """,""" & strTargetGroup & """,""" & strTargetAccount & """,""ERROR: " & Err.Number & " - " & Replace( Err.Description, vbCrLf, "") & """,""" & Now() & """" )
        End If

        Err.Clear
      End If

    End If

    Set objTargetGroup = Nothing
    set objTargetAccount = Nothing
  End If
Next

objLogFile.Close

'======================================================
' EOF: removeUserFromWorkstationGroup.vbs
'======================================================