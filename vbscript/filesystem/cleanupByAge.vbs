'======================================================
' Script for Deleting Files/Folders by Age
'   by Tom Gehrke
'   March 30, 2004
'======================================================
' USAGE:
'
' CleanupByAge
'   <path to folder> - REQUIRED
'     This is the starting folder to clean.  All
'     subfolders will be subject to cleaning/deletion
'     based on the age threshhold.
'
'   [notify=<smtp address>]
'     The e-mail address to receive reporting information
'
'   [from=<smtp address>]
'     The originating e-mail account
'
'   [days=<age threshhold>] - defaults to 30
'     If a file or folder is equal to or greater in age
'     than the number of days passed, it will be deleted.
'     If not passed, this defaults to 30 days.
'
'   [dateattribute=<created|modified|accessed>] - defaults to created
'     If a file or folder is equal to or greater in age
'     than the number of days passed, it will be deleted.
'     If not passed, this defaults to 30 days.
'
'   [testmode=<on|yes|true|1>] - defaults to "off"
'     Turning this on turns this into a reporting only
'     script.  No files/folders will be deleted, however
'     the report will reflect information on what would
'     have been processed if things had been deleted.
'
'   [showok=<on|yes|true|1>] - defaults to "off"
'     When this setting is "on", the resulting report
'     will list ALL files processed and not just the
'     files that were deleted.
'
'   [ignorefolders=<on|yes|true|1>] - defaults to "off"
'     When this setting is "on" folders will not be
'     removed even if their age is older than the
'     threshold.
'
'   [onlyfolders=<on|yes|true|1>] - defaults to "off"
'     When this setting is "on" only folders will be
'     removed as long as they are empty and their
'     age exceeds or equals the threshold.
'
'======================================================
' RECORD OF REVISION:
' -----------------------------------------------------
' March 30, 2004                             Tom Gehrke
' -----------------------------------------------------
' - Cosmetic changes
' -----------------------------------------------------
' March 29, 2004                             Tom Gehrke
' -----------------------------------------------------
' - Added option to only process folders
' - Enhanced error messages to include script section
'   that caused the error
' -----------------------------------------------------
' June 5, 2003                               Tom Gehrke
' -----------------------------------------------------
' - Added option to not remove folders
' -----------------------------------------------------
' December 15, 2001                          Tom Gehrke
' -----------------------------------------------------
' - If email address is provided, but for some reason
'   an address cannot be sent, a log file will be
'   created so that reporting information is not lost.
' - Error handling finally added.  Errors will be
'   logged inline in the report.
' - Logfiles are not written to a "Cleanup Logs"
'   folder that is created in "My Documents".  If the
'   folder does not exists, it will be created.
' - Log file naming has changed to list the path
'   before the date/time.  Should allow for easier
'   grouping in Windows Explorer.
' -----------------------------------------------------
' December 13, 2001                          Tom Gehrke
' -----------------------------------------------------
' - Fixed problem where files/folders were deleted
'   only if age was exceeded by a day.  Now files are
'   deleted if their age is greater than or equal to
'   the threshold number of days.
' - Cosmetic changes.
' - Added options for date attributes other than
'   only allowing deletion based on file creation
'   date.
' -----------------------------------------------------
' December 12, 2001                          Tom Gehrke
' -----------------------------------------------------
' - Modified command line options to allow for
'   drag-and-drop folder cleaning.
' - Checks to see if what was passed as the starting
'   folder exists and is indeed a folder.
' - E-mail address arguments are now optional.
' - If no addresses are passed, the script
'   will generate an HTML file locally.  This means that
'   the script is now usable on workstations and not
'   just servers.
' - If only "notify" or "from" are passed, the
'   option not passed uses the value from the
'   option that was passed.
' -----------------------------------------------------
' December 11, 2001                          Tom Gehrke
' -----------------------------------------------------
' - Folder deletion added.
' - Documentation in the form of comments added.
' - Added graphical usage information at end of report.
' -----------------------------------------------------
' December 10, 2001                          Tom Gehrke
' -----------------------------------------------------
' - Modified for command line use instead of requiring
'   the script to be modified for every instance.
' -----------------------------------------------------
' November 27, 2001                          Tom Gehrke
' -----------------------------------------------------
' - Script created.
'======================================================

' *****************************************************
'    Sets up operating parameters
' *****************************************************
Dim bolTestMode
Dim bolShowOKFiles
Dim strRootFolder
Dim intMaxDays
Dim strTestAddress
Dim strNotification
Dim strDateAttribute
Dim bolIgnoreFolders
Dim bolOnlyFolders

' -----------------------------------------------------
' Defaults
' -----------------------------------------------------
  bolTestMode      = False
  bolShowOKFiles   = False
  strRootFolder    = ""
  strNotification  = ""
  strFromAddress   = ""
  intMaxDays       = 30
  strDateAttribute = ""
  bolIgnoreFolders = False
  bolOnlyFolders   = False

' -----------------------------------------------------

' *****************************************************
'    Set up Error Handler
' *****************************************************
Dim strErrorMessage
strErrorMessage = ""

ON ERROR RESUME NEXT

' *****************************************************
'    Set up Log File Location
' *****************************************************
Dim strLogFolder
Dim objShell
set objShell = WScript.CreateObject( "WScript.Shell" )
strLogFolder = objShell.SpecialFolders( "MyDocuments" ) & "\Cleanup Logs"

' *****************************************************
'    Read command line options
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
    strThisOption   = Left( strThisArgument, InStr( strThisArgument, "=" ) - 1 )
    strThisValue    = Mid( strThisArgument, InStr( strThisArgument, "=" ) + 1 )
  Else
    strThisValue    = strThisArgument
  End If
  Select Case strThisOption
  Case "days"
    intMaxDays = CInt( strThisValue )
  Case "notify"
    strNotification = strThisValue
  Case "from"
    strFromAddress = strThisValue
  Case "testmode"
    If strThisValue = "on" Or strThisValue = "yes" Or strThisValue = "1" or strThisValue = "true" Then
      bolTestMode = True
    End If
  Case "showok"
    If strThisValue = "on" Or strThisValue = "yes" Or strThisValue = "1" or strThisValue = "true" Then
      bolShowOKFiles = True
    End If
  Case "dateattribute"
    strDateAttribute = strThisValue
  Case "ignorefolders"
    If strThisValue = "on" Or strThisValue = "yes" Or strThisValue = "1" or strThisValue = "true" Then
      bolIgnoreFolders = True
    End If
  Case "onlyfolders"
    If strThisValue = "on" Or strThisValue = "yes" Or strThisValue = "1" or strThisValue = "true" Then
      bolOnlyFolders = True
    End If
  Case Else
    strRootFolder = Trim(strThisValue)
    If Right(strRootFolder,1) = "\" Then
      strRootFolder = Left(strRootFolder, Len(strRootFolder)-1)
    End If
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

strReportHeader = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">" & _
  "<html><head>" & _
  "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & _
  "<TITLE>Folder Cleanup Report</TITLE>" & _
  "<style>" & _
    "body{background-color:#ffffff;}" & _
    ".Heading{text-align:center;font-size:12pt;color:#ffffff;font-family:Arial;font-weight:bolder;background-color:#000000;padding:2px;}" & _
    ".Error{text-align:center;border: #800000 3px solid;font-size: 12pt;color:#800000;font-family:'Courier New';background-color: #ffff00;padding:2px;font-weight:bolder;margin:10px;}" & _
    ".Status{border-right: #000000 2px solid;border-top: #808080 1px solid;font-size:8pt;border-left: #000000 2px solid;color:#000000;border-bottom: #808080 1px solid;font-family: 'Courier New';background-color:#d0d0d0;padding:10px;}" & _
    ".FileOK{border-right: #000000 1px solid;border-top: #000000 1px solid;FONT-SIZE: 8pt;border-left: #000000 1px solid;COLOR: #0000ff;border-bottom: #000000 1px solid;font-family: 'Courier New';background-color: #c0ffc0; padding:5px;}" & _
    ".FileDeleted{border-right: #000000 1px solid;border-top: #000000 1px solid;FONT-SIZE: 8pt;border-left: #000000 1px solid;COLOR: #000000;border-bottom: #000000 1px solid;FONT-FAMILY: 'Courier New';background-color: #ffc0c0; padding:5px;}" & _
    ".FolderProcessed{border: solid 1px #808080;font-size: 8pt;color: #808080;font-family: 'Courier New';background-color: #ffffff; padding:5px;}" & _
    ".FolderDeleted{border-right: #000000 1px solid;border-top: #000000 1px solid;font-size: 8pt;border-left: #000000 1px solid;color: #ffffff;border-bottom: #000000 1px solid;font-family: 'Courier New';background-color: #800000; padding:5px;}" & _
    ".PercentUsedElsewhere{background-color:#808080;border: 3px outset #ffffff;}" & _
    ".PercentUsedHere{background-color:#800000;border: 3px outset #ff0000;}" & _
    ".PercentRecovered{background-color:#0000FF;border: 3px outset #c0c0ff;}" & _
    ".PercentAvailable{background-color:#008000;border: 3px outset #00ff00;}" & _
    "#Graph{background-color:#FFFFFF;border:2px #000000 solid; padding:10px;font-family:Arial;font-size:8pt;font-weight:bolder;text-align:center;}" & _
  "</style>" & _
  "</head><body>"

' *****************************************************
'    Creation of the FileSystem Object.  This is what
'    does all the work.
' *****************************************************
Dim objFileSystem
Set objFileSystem = CreateObject( "Scripting.FileSystemObject" )
If Err.Number > 0 Then
  strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, Err.Description & " [creating filesystem object]" )
  Err.Clear
End If

' *****************************************************
'    Check to see that all the required information
'    is available.
' *****************************************************
Dim strThisMessage
strThisMessage  = ""

If strRootFolder = "" Then
  strErrorMessage = strErrorMessage & FormatErrorMessage( 0, "NO STARTING FOLDER SPECIFIED!<br />&lt;path to root folder/share to be cleaned&gt;" )
Else
  If Not objFileSystem.FolderExists( strRootFolder ) Then
    strErrorMessage = strErrorMessage & FormatErrorMessage( 0, "STARTING FOLDER DOES NOT EXIST OR IS NOT REALLY A FOLDER!" )
  End If
End If


If strNotification = "" And strFromAddress <> "" Then
  strNotification = strFromAddress
End If

If strFromAddress = "" And strNotification <> "" Then
  strFromAddress = strNotification
End If

If intMaxDays < 0 Then
  strErrorMessage = strErrorMessage & FormatErrorMessage( 0, "AGE OF FILES/FOLDERS TO BE DELETED MUST BE AT LEAST ZERO (0) DAYS OLD!<br />days=&lt;file age threshold in days&gt;" )
End If

If strErrorMessage <> "" Then
  strReportBody = strReportBody & "ERRORS FOUND!<br /><br />"
  strReportBody = strReportBody & strErrorMessage
Else
  ' *****************************************************
  '    Set variables
  ' *****************************************************
  Dim intTotalFolders
  Dim intTotalFiles
  Dim intTotalSize
  Dim intTotalDeleted
  intTotalFolders      = 0
  intTotalFiles        = 0
  intTotalSize         = 0
  intTotalDeletedFiles = 0
  intTotalDeletedSize  = 0

  strReportHeader = strReportHeader & "<div class=""Heading"">FILE/FOLDER CLEANUP TASK</div>" & _
  "<div class=""Status"">" & _
  "Starting Task at <B>" & CStr( datStart ) & "</B><br />" & _
  "<HR>" & _
  "Start.........: <B>" & strRootFolder & "</B><br />" & _
  "Days..........: <B>" & intMaxDays & "</B><br />" & _
  "DateAtrribute.: <B>" & strDateAttribute & "</B><br />" & _
  "TestMode......: <B>" & bolTestMode & "</B><br />" & _
  "Show OK Files.: <B>" & bolShowOKFiles & "</B><br />" & _
  "Ignore Folders: <B>" & bolIgnoreFolders & "</B><br />" & _
  "Only Folders..: <B>" & bolOnlyFolders & "</B><br />" & _
  "</div>"

  ' *****************************************************
  '    Some pre-processing stuff
  ' *****************************************************
  Dim objThisDrive
  Set objThisDrive = objFileSystem.GetDrive( objFileSystem.GetDriveName( strRootFolder ) )
  If Err.Number > 0 Then
    strReportBody = strReportBody & FormatErrorMessage( Err.Number, Err.Description & " [getting drive]")
    Err.Clear
  End If

  Dim intAvailableSpace
  Dim intTotalVolumeSize
  intAvailableSpace       = Int( objThisDrive.AvailableSpace/1024 )
  intTotalVolumeSize      = Int( objThisDrive.TotalSize/1024 )

  Set objThisDrive = Nothing

  ' *****************************************************
  '    Beginning of actual processing.
  ' *****************************************************
  ProcessFolder strRootFolder, True

  ' *****************************************************
  '    Finish up body of the notification e-mail.
  ' *****************************************************
  If intTotalSize > 0 Then
    intTotalSize = intTotalSize/1024
  End If
  If intTotalDeletedSize > 0 Then
    intTotalDeletedSize = intTotalDeletedSize/1024
  End If

  strReportHeader = strReportHeader & "<div class=""Status"">" & _
    "Folders Processed: " & intTotalFolders            & "<br />" & _
    "Files Processed..: " & intTotalFiles              & "<br />" & _
    "Files Deleted....: " & intTotalDeletedFiles       & "<br />" & _
    "Space Used.......: " & Int( intTotalSize )        & "kb<br />" & _
    "Space Recovered..: " & Int( intTotalDeletedSize ) & "kb<br />" & _
    "</div>"

  Dim intPercentUsed
  Dim intPercentRecovered
  Dim intPercentAvailable
  Dim intPercentUsedElsewhere
  Dim intMultiplier
  Dim intBorderWidth

  intPercentRecovered     = Int( ( intTotalDeletedSize / intTotalVolumeSize ) * 100 )
  intPercentUsed          = Int( ( intTotalUsedSize / intTotalVolumeSize ) * 100 )
  intPercentAvailable     = Int( ( intAvailableSpace / intTotalVolumeSize ) * 100 )
  intPercentUsedElsewhere = 100 - ( intPercentRecovered + intPercentUsed + intPercentAvailable )
  intSectionMultiplier    = 4
  intBorderWidth          = 6

  strReportHeader = strReportHeader & "<div ID=""Graph"">"

  If intPercentUsedElsewhere > 0 Then
    strReportHeader = strReportHeader & "<span class=""PercentUsedElsewhere"" STYLE=""height:30px;width:" & ( intPercentUsedElsewhere * intSectionMultiplier + intBorderWidth ) & "px;"">&nbsp;</span>"
  End If

  If intPercentUsed > 0 Then
    strReportHeader = strReportHeader & "<span class=""PercentUsedHere"" STYLE=""height:30px;width:" & ( intPercentUsed * intSectionMultiplier + intBorderWidth ) & "px;"">&nbsp;</span>"
  End If

  If intPercentRecovered > 0 Then
    strReportHeader = strReportHeader & "<span class=""PercentRecovered"" STYLE=""height:30px;width:" & ( intPercentRecovered * intSectionMultiplier + intBorderWidth ) & "px;"">&nbsp;</span>"
  End If

  If intPercentAvailable > 0 Then
    strReportHeader = strReportHeader & "<span class=""PercentAvailable"" STYLE=""height:30px;width:" & ( intPercentAvailable * intSectionMultiplier + intBorderWidth ) & "px;"">&nbsp;</span>"
  End If

  strReportHeader = strReportHeader & _
    "<div STYLE=""text-align:center;"">Total Volume: " & Int( intTotalVolumeSize/1024 ) & "MB</div>" & _
    "<div STYLE=""text-align:center;margin-top:5px;"">"

  If intPercentUsedElsewhere > 0 Then
    strReportHeader = strReportHeader & _
      "<span class=""PercentUsedElsewhere"" STYLE=""height:15px;width:15px;"">&nbsp;</span>" & _
      "&nbsp;Used&nbsp;Elsewhere&nbsp;(" & intPercentUsedElsewhere & "%)&nbsp;&nbsp;&nbsp;&nbsp;"
  End If

  If intPercentUsed > 0 Then
    strReportHeader = strReportHeader & _
      "<span class=""PercentUsedHere"" STYLE=""height:15px;width:15px;"">&nbsp;</span>" & _
      "&nbsp;Used&nbsp;(" & intPercentUsed & "%)&nbsp;&nbsp;&nbsp;&nbsp;"
  End If

  If intPercentRecovered > 0 Then
    strReportHeader = strReportHeader & _
      "<span class=""PercentRecovered"" STYLE=""height:15px;width:15px;"">&nbsp;</span>" & _
      "&nbsp;Recovered&nbsp;(" & intPercentRecovered & "%)&nbsp;&nbsp;&nbsp;&nbsp;"
  End If

  If intPercentAvailable > 0 Then
    strReportHeader = strReportHeader & _
      "<span class=""PercentAvailable"" STYLE=""height:15px;width:15px;"">&nbsp;</span>" & _
      "&nbsp;Available&nbsp;(" & intPercentAvailable & "%)&nbsp;&nbsp;&nbsp;&nbsp;"
  End If

  strReportHeader = strReportHeader & "</div></div>"
End If

strReportFooter = strReportFooter & "<div class=""Status"">" & _
  "Done in " & DateDiff( "s", datStart, Now() ) & " seconds" & _
  "</div>" & _
  "</body></html>"

' *****************************************************
'    Report the results of the operation.
' *****************************************************
Dim strSendMailError
Dim bolMailSent
strSendMailError = ""
bolMailSent      = False

If strNotification <> "" And strFromAddress <> "" Then
  strSendMailError = SendMail( strNotification, strFromAddress, "[CleanupByAge] " & UCase( strRootFolder ) & " Cleaned!", strReportHeader & strReportBody & strReportFooter )
  If strSendMailError = "" Then
    bolMailSent = True
  Else
    strReportBody = strReportBody & strSendMailError
  End If
End If

If Not bolMailSent Then
  Dim strLogName
  Dim objLogFile
  strLogName = strLogFolder & "\CleanupByAge [" & Replace( Replace( Replace( UCase( strRootFolder ) & "] " & CStr( Year(datStart)) & Right( "0" & CStr( Month(datStart)), 2) & Right( "0" & CStr( Day(datStart)), 2), "/", "][" ), ":", "" ), "\", "][" ) + ".htm"
  If Not objFileSystem.FolderExists( strLogFolder ) Then
    objFileSystem.CreateFolder( strLogFolder )
  End If
  Set objLogFile = objFileSystem.CreateTextFile( strLogName, True )
  objLogFile.Write( strReportHeader & strReportBody & strReportFooter )
  objLogFile.Close
  Set objLogFile = Nothing
End If

' *****************************************************
'    Cleanup.
' *****************************************************
Set objFileSystem = Nothing

'****************************************************************************
' PROCEDURES
'****************************************************************************
'----------------------------------------------------------------------------
' ProcessFolder()
'----------------------------------------------------------------------------
Sub ProcessFolder( strThisFolder, bolIsStart )

  ON ERROR RESUME NEXT

  If objFileSystem.FolderExists( strThisFolder ) Then

    intTotalFolders = intTotalFolders + 1

    Dim objSubFolders
    Dim objCurrentFolder
    Dim objNextFolder
    Dim objFiles
    Dim objCurrentFile
    Dim objOriginalFile
    Dim strCurrentStyle

    Set objCurrentFolder = objFileSystem.GetFolder( strThisFolder )
    Set objSubFolders    = objCurrentFolder.SubFolders
    If Err.Number > 0 Then
      strReportBody = strReportBody & FormatErrorMessage( Err.Number, Err.Description & " [loading sub-folders]")
      Err.Clear
    End If

    For Each objNextFolder In objSubFolders
      ProcessFolder strThisFolder & "\" & objNextFolder.Name, False
    Next

    strReportBody = strReportBody & "<div class=""FolderProcessed"">PROCESSING: " & strThisFolder & "</div>"
    Set objFiles = objCurrentFolder.Files
    If Err.Number > 0 Then
      strReportBody = strReportBody & FormatErrorMessage( Err.Number, Err.Description & " [loading files]")
      Err.Clear
    End If

    Dim intDaysOld
    intDaysOld = 0

    If objFiles.Count = 0 And objSubFolders.Count = 0 And Not bolIsStart And Not bolIgnoreFolders Then
      intDaysOld = GetDaysOld( objCurrentFolder, strDateAttribute )
      If intDaysOld >= intMaxDays Then
        strReportBody = strReportBody & "<div class=""FolderDeleted"">Folder Deleted: " & objCurrentFolder.Path & " (" & intDaysOld & " days old)</div>"
        If Not bolTestMode Then
          objCurrentFolder.Delete( True )
          If Err.Number > 0 Then
            strReportBody = strReportBody & FormatErrorMessage( Err.Number, Err.Description & " [deleting folder]")
            Err.Clear
          End If
        End If
      End If
    Else
      If bolOnlyFolders = False Then

        Dim strCurrentFile

        For Each objCurrentFile In objFiles
          intTotalFiles  = intTotalFiles + 1
          intTotalSize   = intTotalSize + objCurrentFile.Size
          intDaysOld     = GetDaysOld( objCurrentFile, strDateAttribute )

          strCurrentFile = objCurrentFile.Path

          If intDaysOld >= intMaxDays Then
            intTotalDeletedSize  = intTotalDeletedSize + objCurrentFile.Size
            intTotalDeletedFiles = intTotalDeletedFiles + 1

            If Not bolTestMode Then
              objCurrentFile.Attributes = 0
              objFileSystem.DeleteFile strCurrentFile, True
              If Err.Number > 0 Then
                strReportBody = strReportBody & FormatErrorMessage( Err.Number, Err.Description & " [deleting file]")
                Err.Clear
              End If
            End If
            strReportBody = strReportBody & "<div class=""FileDeleted"">File Deleted: " & strCurrentFile & " (" & intDaysOld & " days old)</div>"
          Else
            If bolShowOKFiles Then
              strReportBody = strReportBody & "<div class=""FileOK"">File OK: " & strCurrentFile & " (" & intDaysOld & " days old)</div>"
            End If
          End If
        Next

      End If
    End If

    Set objFiles         = Nothing
    Set objOriginalFile  = Nothing
    Set objSubFolders    = Nothing
    Set objCurrentFolder = Nothing
  End If
End Sub

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
    strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [creating CDONTS object]")
    Err.Clear

    Set objMail = CreateObject( "CDO.Message" )
    If Err.Number > 0 Then
      strErrorMessage = strErrorMessage & FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [creating CDOSYS object]")
      Err.Clear
    Else
      strMailComponent = "CDOSYS"
    End If
  Else
    strMailComponent = "CDONTS"
  End If


  If strMailComponent <> "" Then
    objMail.To         = strTo
    objMail.Subject    = strSubject

    Select Case strMailComponent
    Case "CDONTS"
      objMail.From       = strFrom
      objMail.Body       = strBody
      objMail.BodyFormat = 0
      objMail.MailFormat = 0

      objMail.Send
      If Err.Number > 0 Then
        strErrorMessage = FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [sending mail]")
        Err.Clear
      End If

      Set objMail    = Nothing
    Case "CDOSYS"
      objMail.Sender   = strFrom
      objMail.HTMLBody = strBody

      objMail.Send
      If Err.Number > 0 Then
        strErrorMessage = FormatErrorMessage( Err.Number, "[SendMail] " & Err.Description & " [sending mail]")
        Err.Clear
      End If

      Set objMail    = Nothing
    End Select
  End If

  SendMail = strErrorMessage
End Function

'-------------------------------------------------------------
' GetDaysOld
'-------------------------------------------------------------
Function GetDaysOld( objFile, strAttribute )
  Dim intReturnValue
  intReturnValue = 0

  Select Case strAttribute
  Case "modified"
    intReturnValue = Int( datStart - objFile.DateLastModified )
  Case "accessed"
    intReturnValue = Int( datStart - objFile.DateLastAccessed )
  Case Else
    intReturnValue = Int( datStart - objFile.DateCreated )
  End Select

  GetDaysOld = intReturnValue
End Function

'-------------------------------------------------------------
' FormatErrorMessages
'-------------------------------------------------------------
Function FormatErrorMessage( intErrorNumber, strErrorDescription )
  FormatErrorMessage = "<div class=""Error""><B>ERROR:</B> (" & intErrorNumber & ") " & strErrorDescription & "</div>"
End Function

'======================================================
' EOF: CleanupByAge.VBS
'======================================================