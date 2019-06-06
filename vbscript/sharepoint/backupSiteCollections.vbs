' ===============================================
' BackupSiteCollections.vbs
' -- May 15, 2007
' -- Thomas Gehrke
' -----------------------------------------------
' For the automated backup of MOSS 2007 site
' collections.  Includes emailed backup report
' and resource utilization.
' -----------------------------------------------
' RECORD OF REVISION:
' 03/15/2006 [tcg] Script created.
' -----------------------------------------------
' USAGE:
'
' ===============================================
Option Explicit

On Error Resume Next

' *****************************************************
' BASE
' *****************************************************
Dim strErrorMessage, objShell, intShellExecStatus
strErrorMessage = ""
Set objShell = CreateObject("WScript.Shell")

' *****************************************************
' PARAMETERS
' *****************************************************
Dim strBinPath, strBackupPath, strBaseURL, strNotification, strFromAddress

' *****************************************************
' DEFAULTS
' *****************************************************
strBinPath = ""
strBackupPath = ".\"
strURL = ""
strNotification = ""
strFromAddress = ""

' *****************************************************
' SET UP LOGGING
' *****************************************************
Dim strLogFolder
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Site Collection Backup Logs"

' *****************************************************
' READ ARGUMENTS
' *****************************************************
Dim objArguments
Set objArguments = WScript.Arguments
If Err.Number > 0 Then
  strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, Err.Description & " [receiving arguments]")
  Err.Clear
End If

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
  Case "notify"
    strNotification = strThisValue
  Case "from"
    strFromAddress = strThisValue
  Case "baseurl"
    strBaseUrl = strThisValue
  Case "backuppath"
    strBackupPath = strThisValue
    If Right(strBackupPath,1) <> "\" Then
      strBackupPath = strBackupPath & "\"
    End If
  Case "binpath"
    strBinPath = strThisValue
  Case Else
    strBaseUrl = strThisValue
  End Select
Next

' *****************************************************
'    Begins building response to be sent as an e-mail.
'    Note that it's being formatted as HTML.
' *****************************************************
Dim datStart
datStart = Now()

Dim strReportHeader
Dim strReportBody
Dim strReportFooter

strReportHeader = ""
strReportBody    = ""
strReportFooter  = ""

strReportHeader = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">" & vbCrLf & _
  "<html><head>" & vbCrLf & _
  "<title>Folder Cleanup Report</title>" & vbCrLf & _
  "<style type=""text/css"">" & vbCrLf & _
    "  body {background-color:#ffffff;font-family:calibri,arial;}" & vbCrLf & _
    "  table {width:100%;border-collapse:collapse;}" & vbCrLf & _
    "  .status {font-family:consolas,courier new;border:2px #c0c0c0 solid;background-color:#e0e0e0;color:#808080;padding:4px;font-size:12px;}" & vbCrLf & _
    "  .error {border:2px #800000 solid;background-color:#ffff00;color:#800000;padding:4px;font-size:12px;font-weight:bolder;}" & vbCrLf & _
    "  .log {border:1px #808080 solid;}" & vbCrLf & _
    "  .log th {border:1px #808080 solid;background-color:#808080;font-size:12px;font-weight:bolder;}" & vbCrLf & _
    "  .log td {border:1px #808080 solid;font-size:11px;padding:2px;}" & vbCrLf & _
  "</style>" & vbCrLf & _
  "</head><body>"

strReportHeader = strReportHeader & "<h1>SITE COLLECTION BACKUP</h1>" & vbCrLf & _
  "<table><tr><td class=""status"">" & vbCrLf & _
  "Starting Task at <em>" & CStr( datStart ) & "</em><br />" & vbCrLf & _
  "<hr>" & vbCrLf & _
  "Base URL......: <em>" & strBaseURL & "</em><br />" & vbCrLf & _
  "Backup Path ..: <em>" & strBackupPath & "</em><br />" & vbCrLf & _
  "BIN path......: <em>" & strBinPath & "</em><br />" & vbCrLf & _
  "</td></tr></table>"

' ===============================================

Dim strBaseCommandLine

strBaseCommandLine = ""
If Trim(strBinPath) <> "" Then
  strBaseCommandLine = strBaseCommandLine & Trim(strBinPath)
  If Right(strBaseCommandLine,1) <> "\" Then
    strBaseCommandLine = strBaseCommandLine & "\"
  End If
End If
strBaseCommandLine = strBaseCommandLine & "stsadm.exe"

Dim objFileSystem, objFolder, objFiles, objFile, objExec, strResult, objXml, objSiteCollection, objURL, strURL, strFileName, strCommandLine

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

WScript.Echo vbCrLf & "|==> Retrieving site collection list..."
Set objExec = objShell.Exec(strBaseCommandLine & " -o enumsites -url " & strBaseURL)
Do While objExec.Status = 0
  WScript.Sleep 100
Loop
strResult = objExec.StdOut.ReadAll
WScript.Echo strResult

If InStr( 1, LCase(strResult), "</sites>") = 0 Then
  strErrorMessage = strErrorMessage & "Unable to retrieve list of site collections from " & strBaseUrl
Else
  Set objFolder = objFileSystem.GetFolder(strBackupPath)
  Set objFiles = objFolder.Files

  WScript.Echo vbCrLf & "|==> Deleting old backup files..."
  For Each objFile in objFiles
    If LCase(Right(objFile.Name,9)) = ".scbackup" Then
      objFile.Delete(True)
      If Err.Number > 0 Then
        strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, "[Cleanup] " & Err.Description & " [deleting old backup files]")
        Err.Clear
      End If
    End If
  Next

  WScript.Echo vbCrLf & "|==> Loading XML..."
  Set objXml = CreateObject("MSXML2.DOMDocument")
  objXml.LoadXML(strResult)

  WScript.Echo vbCrLf & "|==> Processing XML..."

  Dim datBackupStartTime

  strReportBody = strReportBody & "<table class=""log""><tr><th>Site</th><th>Size<sup>1</sup></th><th>Owner</th><th>Database</th><th>Time<sup>2</sup></th><th>Result</th></tr>" & vbCrLf

  For Each objSiteCollection in objXml.DocumentElement.ChildNodes
      strUrl = objSiteCollection.Attributes.GetNamedItem("Url").Text
      strFileName = strBackupPath & Replace(Replace(Replace(strUrl, "http://", ""),"https://",""), "/", "_") & ".scbackup"
      strCommandLine = strBaseCommandLine & " -o backup -url """ + strUrl + """ -filename """ + strFileName + """"
      WScript.Echo vbCrLf & "|==> Backing up site collection " & strUrl

      strReportBody = strReportBody & "<tr><td><a href=""" & strUrl & """>" & strUrl & "</a></td><td align=""right"">" & objSiteCollection.Attributes.GetNamedItem("StorageUsedMB").Text & "</td><td>" & objSiteCollection.Attributes.GetNamedItem("Owner").Text & "</td><td>" & objSiteCollection.Attributes.GetNamedItem("ContentDatabase").Text & "</td>"

      datBackupStartTime = Now()

      Set intShellExecStatus = objShell.Exec(strCommandLine)
      Do While intShellExecStatus.Status = 0
        WScript.Sleep 100
      Loop
      strResult = objExec.StdOut.ReadAll

      strReportBody = strReportBody & "<td align=""right"">" & DateDiff( "s", datBackupStartTime, Now() ) & "</td><td>" & strResult & "</td></tr>"
      ' WScript.Echo strResult
  Next

  strReportBody = strReportBody & "<tr><td colspan=""6""><sup>1</sup> Size shown in MB<br /><sup>2</sup> Time shown in seconds</td></tr></table>"

End If

WScript.Echo vbCrLf & "|==> Backup of site collections completed!"

If strErrorMessage <> "" Then
  strReportBody = strReportBody & "<table><tr><td class=""error"">[ERROR MESSAGE COLLECTION]<ul>" & strErrorMessage & "</ul></td></tr></table>"
End If

strReportFooter = strReportFooter & "<table><tr><td class=""status"">" & vbCrLf & _
  "Done in " & DateDiff( "s", datStart, Now() ) & " seconds" & vbCrLf & _
  "</td></tr></table>" & vbCrLf & _
  "</body></html>"

' *****************************************************
'    Report the results of the operation.
' *****************************************************
Dim strSendMailError
Dim bolMailSent
strSendMailError = ""
bolMailSent      = False

If strNotification <> "" And strFromAddress <> "" Then
  strSendMailError = SendMail( strNotification, strFromAddress, "[BackupSiteCollections] " & LCase( strBaseUrl ) & " Backed Up!", strReportHeader & strReportBody & strReportFooter )
  If strSendMailError = "" Then
    bolMailSent = True
  Else
    strReportBody = strReportBody & strSendMailError
  End If
End If

If Not bolMailSent Then
  Dim strLogName
  Dim objLogFile
  strLogName = strLogFolder & "\BackupSiteCollections [" & Replace( Replace( Replace( LCase( Replace(Replace(Replace(strBaseUrl,"http://",""),"https://",""),"/","_") ) & "] " & CStr( Year(datStart)) & Right( "0" & CStr( Month(datStart)), 2) & Right( "0" & CStr( Day(datStart)), 2), "/", "][" ), ":", "" ), "\", "][" ) + ".htm"
  If Not objFileSystem.FolderExists( strLogFolder ) Then
    objFileSystem.CreateFolder( strLogFolder )
  End If
  Set objLogFile = objFileSystem.CreateTextFile( strLogName, True )
  objLogFile.Write( strReportHeader & strReportBody & strReportFooter )
  objLogFile.Close
  Set objLogFile = Nothing
End If

'****************************************************************************
' FUNCTIONS
'****************************************************************************
'-------------------------------------------------------------
' SendMail
'-------------------------------------------------------------
Function SendMail( strTo, strFrom, strSubject, strBody )

  ON ERROR RESUME NEXT

  Dim strMailComponent
  strMailComponent = ""

  Dim strErrorMessage
  strErrorMessage = ""

  Dim objMail
  Set objMail = CreateObject( "CDONTS.NewMail" )
  If Err.Number > 0 Then
    ' strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [creating CDONTS object]")
    Err.Clear

    Set objMail = CreateObject( "CDO.Message" )
    If Err.Number > 0 Then
      ' strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [creating CDOSYS object]")
      Err.Clear
    Else
      strMailComponent = "CDOSYS"
    End If
  Else
    strMailComponent = "CDONTS"
  End If


  If strMailComponent <> "" Then
    objMail.To = strTo
    objMail.Subject = strSubject

    Select Case strMailComponent
    Case "CDONTS"
      objMail.From = strFrom
      objMail.Body = strBody
      objMail.BodyFormat = 0
      objMail.MailFormat = 0

      objMail.Send
      If Err.Number > 0 Then
        strErrorMessage = FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [sending mail]")
        Err.Clear
      End If

      Set objMail = Nothing
    Case "CDOSYS"
      objMail.Sender = strFrom
      objMail.HTMLBody = strBody

      objMail.Send
      If Err.Number > 0 Then
        strErrorMessage = FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [sending mail]")
        Err.Clear
      End If

      Set objMail = Nothing
    End Select
  End If

  SendMail = strErrorMessage
End Function

'-------------------------------------------------------------
' FormatErrorMessage
'-------------------------------------------------------------
Function FormatErrorMessage( intErrorNumber, strErrorDescription )
  FormatErrorMessage = "<li><em>ERROR:</em> (" & intErrorNumber & ") " & strErrorDescription & "</li>"
End Function

'======================================================
' EOF: BackupSiteCollections.vbs
'======================================================
