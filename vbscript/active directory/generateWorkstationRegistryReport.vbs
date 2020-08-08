' ===============================================
' generateWorkstationRegistryReport.vbs
' -- July 24, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' Creates a report of a particular registry key
' for all domain workstations.
' -----------------------------------------------
' RECORD OF REVISION:
' 07/24/2006 [tcg] Script created.
' -----------------------------------------------
' USAGE:
'
' ===============================================
Option Explicit

On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS=&H80000003

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7

' *****************************************************
' BASE
' *****************************************************
Dim strErrorMessage, objShell, intShellExecStatus
strErrorMessage = ""
Set objShell = CreateObject("WScript.Shell")

' *****************************************************
' PARAMETERS
' *****************************************************
Dim strDomain, strRegistryKey, strRegistryValueName

' *****************************************************
' DEFAULTS
' *****************************************************
strDomain = ""
strRegistryKey = ""
strRegistryValueName = ""

' *****************************************************
' SET UP LOGGING
' *****************************************************
Dim strLogFolder
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Domain Registry Reports\"

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
  Case "key"
    strRegistryKey = LCase(strThisValue)
  Case "value"
    strRegistryValueName = LCase(strThisValue)
  Case Else
    strRegistryKey = LCase(strThisValue)
  End Select
Next

' ===============================================

Dim objDomain, objItem, objFileSystem, objLogFile, strLogFile, strRegistryRootNode, strRootNode, strRegistryValue

strLogFile = strLogFolder & "Report for " & strDomain & " for value '" & strRegistryValueName & "' - " & Replace( Replace( CStr( Now()), "/", "_"), ":", "_") & ".csv"
strRegistryRootNode = Left( strRegistryKey, InStr(1, strRegistryKey, "\") -1 )
strRegistryKey =  Mid( strRegistryKey, InStr(1, strRegistryKey, "\") + 1 )

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

If Not objFileSystem.FolderExists( strLogFolder ) Then
  objFileSystem.CreateFolder( strLogFolder )
End If

Set objLogFile = objFileSystem.OpenTextFile(strLogFile, ForAppending, True)
Set objDomain = GetObject( "WinNT://" & strDomain )

Dim strComputerName, objRegistry, arrValueNames, arrValueTypes, intValueCounter

objLogFile.WriteLine( """Computer"",""Key"",""Value Name"",""Value""")

For Each objItem In objDomain
  If objItem.Class = "Computer" Then
    strComputerName = UCase( objItem.Name )
    strRegistryValue = ""
    arrValueNames = Null

    Set objRegistry = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\default:StdRegProv" )

    If Err.number = 0 Then
      Select Case strRegistryRootNode
      Case "hklm"
        strRootNode = HKEY_LOCAL_MACHINE
      Case "hkey_local_machine"
        strRootNode = HKEY_LOCAL_MACHINE
      Case "hkcu"
        strRootNode = HKEY_CURRENT_USER
      Case "hkey_current_user"
        strRootNode = HKEY_CURRENT_USER
      Case "hku"
        strRootNode = HKEY_USERS
      Case "hkey_users"
        strRootNode = HKEY_USERS
      Case Else
        strRootNode = HKEY_LOCAL_MACHINE
      End Select

      objRegistry.EnumValues strRootNode, strRegistryKey, arrValueNames, arrValueTypes

      If Err.number = 0 And NOT IsNull(arrValueNames) Then
        For intValueCounter = 0 To UBound( arrValueNames )
          If LCase( arrValueNames( intValueCounter ) ) = strRegistryValueName Then

            Select Case arrValueTypes(intValueCounter)
            Case REG_SZ
              objRegistry.GetStringValue strRootNode, strRegistryKey, strRegistryValueName, strRegistryValue

            Case REG_EXPAND_SZ
              objRegistry.GetExpandedStringValue strRootNode, strRegistryKey, strRegistryValueName, strRegistryValue

            Case REG_BINARY
              Dim strBinaryValue, intBinaryCounter
              objRegistry.GetBinaryValue strRootNode, strRegistryKey, strRegistryValueName, strBinaryValue

              For intBinaryCounter = lBound(strBinaryValue) To uBound(strBinaryValue)
                strRegistryValue = strRegistryValue & " " & strBinaryValue(intBinaryCounter)
              Next

              strRegistryValue = Trim( strRegistryValue )

            Case REG_DWORD
              objRegistry.GetDWORDValue strRootNode, strRegistryKey, strRegistryValueName, strRegistryValue

            Case REG_MULTI_SZ
              Dim arrValues, strValue
              objRegistry.GetMultiStringValue strRootNode, strRegistryKey, strRegistryValueName, arrValues

              For Each strValue In arrValues
                strRegistryValue = strRegistryValue & " " & strValue
              Next

              strRegistryValue = Trim( strRegistryValue )

            End Select
          End If
        Next

        If strRegistryValue = "" Then
          strRegistryValue = "** Value not found **"
        End If
      Else
        strRegistryValue = "** Key not found **"
      End If
    End If

    If Err.Number > 0 Then
      strRegistryValue = "** " & Err.Description & " **"
      Err.Clear
    End If

    objLogFile.WriteLine( """" & strDomain & "\" & strComputerName & """,""" & UCase( strRegistryRootNode ) & "\" & strRegistryKey & """,""" & strRegistryValueName & """,""" & Trim( strRegistryValue ) & """")

  End If
Next

objLogFile.WriteLine( "** EOF **")

objLogFile.Close

'****************************************************************************
' FUNCTIONS
'****************************************************************************

'======================================================
' EOF: generateWorkstationRegistryReport.VBS
'======================================================
