' ===============================================
' exportUsers.vbs
' -- September 4, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' Exports User objects from Active Directory
' -----------------------------------------------
' RECORD OF REVISION:
' 09/04/2007 [tcg] Script created.
' -----------------------------------------------
' USAGE:
'
' ===============================================
Option Explicit

' On Error Resume Next

Const ADS_SCOPE_BASE = 0
Const ADS_SCOPE_ONELEVEL = 1
Const ADS_SCOPE_SUBTREE = 2
Const ADS_SECURE_AUTHENTICATION = 1

' *****************************************************
' PARAMETERS
' *****************************************************
Dim strContainer

' *****************************************************
' DEFAULTS
' *****************************************************
strContainer = ""

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
  strThisOption = ""
  strThisValue = ""
  strThisArgument = LCase( objArguments( intCounter ) )

  If inStr( strThisArgument, "=" ) > 0 Then
    strThisOption = Trim(Left( strThisArgument, InStr( strThisArgument, "=" ) - 1 ))
    strThisValue = Trim(Mid( strThisArgument, InStr( strThisArgument, "=" ) + 1 ))
  Else
    strThisValue = Trim(strThisArgument)
  End If

  Select Case strThisOption
  Case "container"
    strContainer = strContainer & strThisValue
  Case Else
    strContainer = strContainer & strThisArgument
  End Select
Next

' ===============================================

Dim objFile, objFileSystem, objConnection, objCommand, objRecordSet

Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSystem.CreateTextFile("export.txt", True)
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 500
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCommand.Properties("Chase referrals") = &H40

objCommand.CommandText = "Select ADsPath, cn, displayName, givenName, initials, sn, mail from 'LDAP://" & strContainer &"' Where objectClass='user'"
Set objRecordSet = objCommand.Execute

EnumerateUsers(objRecordSet)

objFile.Close
Set objFileSystem = Nothing
Set objConnection = Nothing
Set objCommand = Nothing

' ===============================================
Sub EnumerateUsers(RecordSet)
  RecordSet.MoveFirst
  Do Until RecordSet.EOF
    objFile.WriteLine "adspath: " & RecordSet.Fields("ADsPath").Value
    objFile.WriteLine "display name: " & RecordSet.Fields("displayName").Value
    objFile.WriteLine "first name: " & RecordSet.Fields("givenName").Value
    objFile.WriteLine "initials: " & RecordSet.Fields("initials").Value
    objFile.WriteLine "last name: " & RecordSet.Fields("sn").Value
    objFile.WriteLine "email: " & RecordSet.Fields("mail").Value
    RecordSet.MoveNext
  Loop
End Sub