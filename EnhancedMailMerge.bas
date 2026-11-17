Attribute VB_Name = "EnhancedMailMerge"
Option Explicit

' ============================================================================
' Enhanced Mail Merge to Email
' Macro created by Richmond Tang
' Please share freely while retaining attribution
' Last Updated 17-Nov-2026
' Please copy all the text in this module - it contains several procedures
' that work together.
'
' Revision notes (17-Nov-2026, v3):
'  - PROGRESS LOG: every email is written to a CSV log as soon as it is
'    sent / saved, so you always know which record was done last - even after
'    an error, a stop, or Word crashing. The log is saved next to this
'    document as "<document name>_MailMergeLog.csv" (or in Documents if the
'    document is unsaved or stored on OneDrive/SharePoint).
'  - RESUME: at start-up you choose the record to start from. The macro reads
'    the log and suggests the record after the last one completed.
'  - BATCH SIZE: you can choose how many emails to process in a run.
'  - STOP: hold Esc to stop after the current email. The key is now checked
'    several times per email, and all prompts are forced to the front.
'  - SHOW_EACH_EMAIL now defaults to False (Outlook windows were covering
'    Word's status bar and prompts, and catching the Esc key).
'
' Earlier revisions: error handling and clean-up; original merge settings
' restored; duplicate Subject no longer wipes other errors; leading ";"
' removed; error drafts set the body directly; past delivery times reported;
' consistent yes/no parsing; multiple attachments separated by ";"; included
' record count; Outlook-not-open warning; end-of-run summary;
' ResetMailMergeSettings repair macro.
' ============================================================================

' Set to True to open an Outlook window for every email (slower). Only needed
' if emails arrive with a missing body - try False first.
Private Const SHOW_EACH_EMAIL As Boolean = False

' MsgBox style that forces the box to the front (vbSystemModal 4096 + vbMsgBoxSetForeground 65536)
Private Const MB_FRONT As Long = 69632

' Windows API used to detect the Esc key while the macro is running
' (requires Office 2010 or later, 32-bit or 64-bit)
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private mStopRequested As Boolean
Private mLogPath As String
Private mLogEnabled As Boolean
Private mLogFailed As Boolean

Sub EnhancedMailMergeToEmail()
    ' declare variables
    Dim outlookApp As Object        'Outlook.Application
    Dim outlookMail As Object       'Outlook.MailItem
    Dim mainDoc As Document
    Dim mm As MailMerge
    Dim df As MailMergeDataField
    Dim singleDoc As Document
    Dim lastRecordNum As Long
    Dim currentRecord As Long
    Dim prevRecord As Long
    Dim firstToProcess As Long
    Dim lastToProcess As Long
    Dim includedCount As Long
    Dim recordIndex As Long
    Dim startRecord As Long
    Dim suggestedStart As Long
    Dim maxCount As Long
    Dim lastDoneRecord As Long
    Dim sendFlag As Boolean
    Dim inputDate As Date
    Dim toString As String
    Dim ccString As String
    Dim bccString As String
    Dim subjectString As String
    Dim errorString As String
    Dim contOnError As Boolean
    Dim fso As Object
    Dim f As Variant
    Dim origFirst As Long
    Dim origLast As Long
    Dim origDest As Long
    Dim stateSaved As Boolean
    Dim outlookWasRunning As Boolean
    Dim sentCount As Long
    Dim draftCount As Long
    Dim errorCount As Long
    Dim userCancelled As Boolean
    Dim runStarted As Boolean
    Dim finished As Boolean
    Dim reachedEnd As Boolean
    Dim recordLogged As Boolean
    Dim stage As String
    Dim errNum As Long
    Dim errDesc As String
    Dim msg As String
    Dim inp As String
    Dim runResult As String

    ' identify the main document and its mail merge
    Set mainDoc = ActiveDocument
    Set mm = mainDoc.MailMerge

    ' check the mail merge is set up and ready to go
    If mm.State <> wdMainAndDataSource Then
        MsgBox "Mail merge is not set up for the active document - cannot perform mail merge. Macro will exit." & _
               vbLf & vbLf & "Please click the Edit Recipient List button to confirm that mail merge is correctly set up for this document.", _
               vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If

    ' remember the original mail merge settings so they can be restored at the end
    origFirst = mm.DataSource.FirstRecord
    origLast = mm.DataSource.LastRecord
    origDest = mm.Destination
    stateSaved = True

    On Error GoTo ErrHandler

    ' progress log location
    mLogPath = GetLogPath(mainDoc)
    mLogEnabled = False
    mLogFailed = False

    ' find the last included record (unticked records are skipped by Word's navigation)
    mm.DataSource.ActiveRecord = wdLastRecord
    lastRecordNum = mm.DataSource.ActiveRecord
    If lastRecordNum <= 0 Then
        MsgBox "No records are included in the recipient list. Macro will exit.", vbOKOnly + vbExclamation, "No records"
        GoTo CleanUp
    End If

    ' ---- choose the start record (suggested from the progress log) ----
    suggestedStart = LastCompletedRecord(mLogPath) + 1
    If suggestedStart > lastRecordNum Then suggestedStart = 1
    msg = "Start from which record number?" & vbLf & vbLf & _
          "Records run from 1 to " & lastRecordNum & ". Unticked records are skipped automatically."
    If suggestedStart > 1 Then
        msg = msg & vbLf & vbLf & "The progress log shows record " & (suggestedStart - 1) & _
              " was the last one completed, so record " & suggestedStart & " is suggested." & _
              " Change it to 1 to start from the beginning."
    End If
    inp = InputBox(msg, "Mail merge - start record", CStr(suggestedStart))
    If Trim(inp) = "" Then GoTo CleanUp                       ' Cancel pressed
    If Not IsNumeric(inp) Then
        MsgBox "'" & inp & "' is not a record number. Macro will exit.", vbExclamation, "Invalid start record"
        GoTo CleanUp
    End If
    startRecord = CLng(inp)
    If startRecord < 1 Or startRecord > lastRecordNum Then
        MsgBox "The start record must be between 1 and " & lastRecordNum & ". Macro will exit.", vbExclamation, "Invalid start record"
        GoTo CleanUp
    End If

    ' ---- choose how many emails to process in this run ----
    inp = InputBox("How many emails do you want to process in this run?" & vbLf & vbLf & _
                   "Leave as ALL to process every remaining record, or enter a number (e.g. 20) to do a batch.", _
                   "Mail merge - batch size", "ALL")
    If Trim(inp) = "" Then GoTo CleanUp                       ' Cancel pressed
    If UCase(Trim(inp)) = "ALL" Then
        maxCount = 0
    ElseIf IsNumeric(inp) Then
        maxCount = CLng(inp)
        If maxCount < 1 Then
            MsgBox "The batch size must be 1 or more. Macro will exit.", vbExclamation, "Invalid batch size"
            GoTo CleanUp
        End If
    Else
        MsgBox "'" & inp & "' is not a valid batch size. Macro will exit.", vbExclamation, "Invalid batch size"
        GoTo CleanUp
    End If

    ' move to the first included record at or after the start record
    mm.DataSource.ActiveRecord = wdFirstRecord
    Do While mm.DataSource.ActiveRecord < startRecord
        prevRecord = mm.DataSource.ActiveRecord
        mm.DataSource.ActiveRecord = wdNextRecord
        If mm.DataSource.ActiveRecord = prevRecord Then Exit Do   ' safety: navigation did not move
    Loop
    firstToProcess = mm.DataSource.ActiveRecord

    ' count the records that will be processed in this run
    Do
        includedCount = includedCount + 1
        lastToProcess = mm.DataSource.ActiveRecord
        If mm.DataSource.ActiveRecord >= lastRecordNum Then Exit Do
        If maxCount > 0 And includedCount >= maxCount Then Exit Do
        prevRecord = mm.DataSource.ActiveRecord
        mm.DataSource.ActiveRecord = wdNextRecord
        If mm.DataSource.ActiveRecord = prevRecord Then Exit Do   ' safety: navigation did not move
    Loop
    mm.DataSource.ActiveRecord = firstToProcess

    ' give the user an opportunity to abort, and the option to save the emails as drafts or send immediately
    Select Case MsgBox("This run will process " & includedCount & " email(s): records " & firstToProcess & " to " & lastToProcess & _
                       " (unticked records are skipped)." & vbLf & vbLf & _
                       "Progress will be logged to:" & vbLf & mLogPath & vbLf & _
                       "(Do not open the log in Excel while the macro is running.)" & vbLf & vbLf & _
                       "To stop early, press and HOLD Esc until a prompt appears." & vbLf & vbLf & _
                       "Click 'Yes' to send the emails immediately, 'No' to save the emails as drafts, and 'Cancel' to abort.", _
                       vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Send Emails")
        Case vbCancel
            GoTo CleanUp
        Case vbYes
            sendFlag = True
        Case Else
            sendFlag = False
    End Select

    ' connect to Outlook (late binding, so no reference to the Outlook library is needed)
    ' use the running copy if there is one, otherwise start it
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    Err.Clear
    On Error GoTo ErrHandler
    outlookWasRunning = Not (outlookApp Is Nothing)
    If Not outlookWasRunning Then Set outlookApp = CreateObject("Outlook.Application")

    ' start the progress log - if it cannot be written, ask before continuing
    mLogEnabled = True
    If Not WriteLog(0, "RUN START", "", "", _
                    IIf(sendFlag, "Mode: send", "Mode: drafts") & "; records " & firstToProcess & " to " & lastToProcess & _
                    "; document: " & mainDoc.FullName) Then
        mLogEnabled = False
        If MsgBox("The progress log could not be written to:" & vbLf & mLogPath & vbLf & vbLf & _
                  "(Is it open in Excel, or is the folder read-only?)" & vbLf & vbLf & _
                  "Continue WITHOUT a progress log?", vbYesNo + vbExclamation + MB_FRONT, "Progress log") = vbNo Then
            GoTo CleanUp
        End If
    End If
    runStarted = True

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    mm.Destination = wdSendToNewDocument

    ' clear any Esc key press made before the loop started
    mStopRequested = False
    Call GetAsyncKeyState(vbKeyEscape)

    ' loop through the records for this run
    Do While Not finished
        currentRecord = mm.DataSource.ActiveRecord
        recordIndex = recordIndex + 1
        recordLogged = False
        stage = "creating the email"

        ' reset fields to be populated for each email
        errorString = ""
        toString = ""
        ccString = ""
        bccString = ""
        subjectString = ""

        ' use mail merge to create a new document for one record (the active record)
        mm.DataSource.FirstRecord = currentRecord
        mm.DataSource.LastRecord = currentRecord
        mm.Execute Pause:=False
        Set singleDoc = ActiveDocument

        Application.StatusBar = "Mail merge: email " & recordIndex & " of " & includedCount & _
                                " (record " & currentRecord & ")  -  hold Esc to stop after this email"
        PollStopKey

        ' extract the "MailEnvelope" from the Word doc, which is an Outlook email containing the contents of the doc
        Set outlookMail = singleDoc.MailEnvelope.Item

        ' run through all the fields in the mail merge data; when an email field is identified add the data to the appropriate property
        For Each df In mm.DataSource.DataFields
            ' only check fields that contain data for the active record
            If Trim(df.Value) <> "" Then

                ' try matching the field name to accepted field names
                Select Case StripToLcaseLetters(df.Name)

                    Case "to"
                        ' if multiple To fields are present they will all be added, separated by a semicolon
                        If InStr(1, df.Value, "@", vbBinaryCompare) > 0 Then
                            toString = toString & ";" & Trim(df.Value)
                        Else
                            errorString = errorString & vbCrLf & "Invalid email address in To field: " & df.Value
                        End If

                    Case "cc"
                        If InStr(1, df.Value, "@", vbBinaryCompare) > 0 Then
                            ccString = ccString & ";" & Trim(df.Value)
                        Else
                            errorString = errorString & vbCrLf & "Invalid email address in CC field: " & df.Value
                        End If

                    Case "bcc"
                        If InStr(1, df.Value, "@", vbBinaryCompare) > 0 Then
                            bccString = bccString & ";" & Trim(df.Value)
                        Else
                            errorString = errorString & vbCrLf & "Invalid email address in BCC field: " & df.Value
                        End If

                    Case "subject"
                        If subjectString = "" Then
                            subjectString = df.Value
                        Else
                            errorString = errorString & vbCrLf & "Second subject field found containing: " & df.Value
                        End If

                    Case "importance"
                        ' accepted values: "high", "normal", "low" (not case sensitive)
                        Select Case StripToLcaseLetters(df.Value)
                            Case ""
                                ' leave as default
                            Case "low", "l"
                                outlookMail.Importance = 0 'olImportanceLow
                            Case "normal", "n"
                                outlookMail.Importance = 1 'olImportanceNormal
                            Case "high", "h"
                                outlookMail.Importance = 2 'olImportanceHigh
                            Case Else
                                errorString = errorString & vbCrLf & "Importance value not recognised: " & df.Value
                        End Select

                    Case "sensitivity"
                        ' accepted values: "normal", "personal", "private", "confidential" (not case sensitive)
                        Select Case StripToLcaseLetters(df.Value)
                            Case ""
                                ' leave as default
                            Case "normal"
                                outlookMail.Sensitivity = 0 'olNormal
                            Case "personal"
                                outlookMail.Sensitivity = 1 'olPersonal
                            Case "private"
                                outlookMail.Sensitivity = 2 'olPrivate
                            Case "confidential"
                                outlookMail.Sensitivity = 3 'olConfidential
                            Case Else
                                errorString = errorString & vbCrLf & "Sensitivity value not recognised: " & df.Value
                        End Select

                    Case "readreceipt"
                        ' "true"/"yes"/"y" requests a read receipt, "false"/"no"/"n" does not (not case sensitive)
                        Select Case StripToLcaseLetters(df.Value)
                            Case ""
                                ' leave as default
                            Case "true", "yes", "y"
                                outlookMail.ReadReceiptRequested = True
                            Case "false", "no", "n"
                                outlookMail.ReadReceiptRequested = False
                            Case Else
                                errorString = errorString & vbCrLf & "Read receipt value not recognised: " & df.Value
                        End Select

                    Case "deliveryreceipt"
                        ' "true"/"yes"/"y" requests a delivery report, "false"/"no"/"n" does not (not case sensitive)
                        Select Case StripToLcaseLetters(df.Value)
                            Case ""
                                ' leave as default
                            Case "true", "yes", "y"
                                outlookMail.OriginatorDeliveryReportRequested = True
                            Case "false", "no", "n"
                                outlookMail.OriginatorDeliveryReportRequested = False
                            Case Else
                                errorString = errorString & vbCrLf & "Delivery receipt value not recognised: " & df.Value
                        End Select

                    Case "deliverytime", "delaydelivery"
                        ' add a delivery time (delay delivery)
                        ' - a future date/time delays delivery until then
                        ' - a date only delays delivery until midnight at the start of that date
                        ' - a time only uses the next occurrence of that time (today or tomorrow)
                        ' - a date/time in the past, or an unrecognised value, is reported as an error
                        ' TIP: to avoid any day/month confusion, supply dates as yyyy-mm-dd hh:mm
                        '      e.g. in Excel use a helper column =TEXT(A2,"yyyy-mm-dd hh:mm")
                        If (IsNumeric(df.Value) Or IsDate(df.Value)) Then
                            ' A date passed from an Excel table through mail merge is formatted in US format (m/d/yyyy),
                            ' but CDate uses the local format (e.g. d/m/yyyy). If the local format is d/m, the day is
                            ' 12 or lower, and the raw text matches the local d/m pattern, swap day and month.
                            ' (ISO yyyy-mm-dd values never match this pattern, so they are never swapped.)
                            inputDate = CDate(df.Value)
                            If Day(inputDate) <= 12 And Month(CDate("1/2/2020")) = 2 And _
                                    (df.Value Like Format(inputDate, "d/m/yyyy") & "*" Or df.Value Like Format(inputDate, "dd/mm/yyyy") & "*") Then
                                inputDate = DateSerial(Year(inputDate), Day(inputDate), Month(inputDate)) + _
                                            TimeSerial(Hour(inputDate), Minute(inputDate), Second(inputDate))
                            End If
                            If inputDate < Now() - Date Then        ' time only, already passed today -> tomorrow
                                outlookMail.DeferredDeliveryTime = Date + 1 + inputDate
                            ElseIf inputDate < 1 Then               ' time only, still to come today -> today
                                outlookMail.DeferredDeliveryTime = Date + inputDate
                            ElseIf inputDate > Now() Then           ' date or date/time in the future
                                outlookMail.DeferredDeliveryTime = inputDate
                            Else                                    ' date or date/time in the past
                                errorString = errorString & vbCrLf & "Delivery time is in the past: " & df.Value
                            End If
                        Else
                            errorString = errorString & vbCrLf & "Delivery time value not recognised: " & df.Value
                        End If

                    Case "sendas", "sendonbehalfof"
                        ' send on behalf of another address (only added if it looks like an email address)
                        ' if the account does not have permission, the email will be rejected by the Exchange server when sent
                        If InStr(1, df.Value, "@", vbTextCompare) > 0 Then
                            outlookMail.SentOnBehalfOfName = Trim(df.Value)
                        Else
                            errorString = errorString & vbCrLf & "Send-as email not recognised: " & df.Value
                        End If

                    Case "attachment", "attachments"
                        ' one or more full file paths, separated by ";"
                        ' files that cannot be found are not attached and are reported as errors
                        If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
                        For Each f In Split(df.Value, ";")
                            If Trim(f) <> "" Then
                                If fso.FileExists(Trim(f)) Then
                                    outlookMail.Attachments.Add CStr(Trim(f))
                                Else
                                    errorString = errorString & vbCrLf & "Attachment could not be found: " & Trim(f)
                                End If
                            End If
                        Next f

                End Select  ' end test for the field names
            End If      ' end check for the data value being blank
        Next df     ' move on to the next field
        PollStopKey

        ' remove the leading ";" from the address lists
        If Len(toString) > 0 Then toString = Mid(toString, 2)
        If Len(ccString) > 0 Then ccString = Mid(ccString, 2)
        If Len(bccString) > 0 Then bccString = Mid(bccString, 2)

        If Len(toString & ccString & bccString) = 0 Then errorString = errorString & vbCrLf & "No valid email addresses provided in To, CC and BCC fields"
        If Len(subjectString) = 0 Then errorString = errorString & vbCrLf & "No subject provided"

        ' set the collected To, CC and BCC fields in one go (some systems do not allow the To field to be read)
        outlookMail.To = toString
        outlookMail.CC = ccString
        outlookMail.BCC = bccString
        outlookMail.Subject = subjectString
        If SHOW_EACH_EMAIL Then outlookMail.Display

        ' send, save as draft, or save as an error draft - then log the result straight away
        If Len(errorString) > 0 Then
            stage = "saving the error draft"
            outlookMail.Subject = "**Errors in mail merge: " & subjectString
            outlookMail.BodyFormat = 1 'olFormatPlain
            outlookMail.Body = "Errors found:" & errorString
            outlookMail.Close 0 'olSave
            Set outlookMail = Nothing
            errorCount = errorCount + 1
            lastDoneRecord = currentRecord
            WriteLog currentRecord, "ERROR DRAFT", toString, subjectString, Replace(Mid(errorString, 3), vbCrLf, " | ")
            recordLogged = True
            If Not contOnError Then
                If MsgBox("Errors were found when creating the email for record " & currentRecord & _
                          ". It has been saved as a draft with the subject """ & _
                          "**Errors in mail merge: " & subjectString & """ and a full list of errors in the body of the email." & vbLf & vbLf & _
                          "Click OK to continue with the remaining records (further errors will be saved as drafts without asking), or Cancel to stop.", _
                          vbOKCancel + vbExclamation + MB_FRONT, "Errors in email") = vbCancel Then
                    userCancelled = True
                    GoTo CleanUp
                End If
                contOnError = True
                mStopRequested = False
                Call GetAsyncKeyState(vbKeyEscape)
            End If
        ElseIf sendFlag Then
            stage = "sending the email"
            outlookMail.Send
            Set outlookMail = Nothing
            sentCount = sentCount + 1
            lastDoneRecord = currentRecord
            WriteLog currentRecord, "SENT", toString, subjectString, ""
            recordLogged = True
        Else
            stage = "saving the draft"
            outlookMail.Close 0 'olSave
            Set outlookMail = Nothing
            draftCount = draftCount + 1
            lastDoneRecord = currentRecord
            WriteLog currentRecord, "DRAFT", toString, subjectString, ""
            recordLogged = True
        End If

        stage = "closing the merged document"
        singleDoc.Close False
        Set singleDoc = Nothing
        PollStopKey

        ' finished the run, stop requested, or move to the next included record
        If mm.DataSource.ActiveRecord >= lastRecordNum Then
            reachedEnd = True
            finished = True
        ElseIf recordIndex >= includedCount Then
            finished = True                                   ' batch size reached
        Else
            If mStopRequested Then
                If MsgBox("Stop the mail merge now?" & vbLf & vbLf & _
                          recordIndex & " of " & includedCount & " emails in this run have been processed " & _
                          "(last completed: record " & lastDoneRecord & ")." & vbLf & vbLf & _
                          "Click Yes to stop, or No to carry on.", _
                          vbYesNo + vbQuestion + vbDefaultButton1 + MB_FRONT, "Stop mail merge?") = vbYes Then
                    userCancelled = True
                    GoTo CleanUp
                End If
                mStopRequested = False
                Call GetAsyncKeyState(vbKeyEscape)            ' clear the key press if the user chose to carry on
            End If
            mm.DataSource.ActiveRecord = wdNextRecord
        End If
    Loop

    GoTo CleanUp

ErrHandler:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp

CleanUp:
    On Error Resume Next

    ' log the record that failed (if it was not already logged)
    If errNum <> 0 And runStarted And currentRecord > 0 And Not recordLogged Then
        WriteLog currentRecord, "FAILED", toString, subjectString, _
                 "Error " & errNum & ": " & errDesc & " (while " & stage & ")" & _
                 IIf(stage = "sending the email", " - this email MAY have been sent: check Outbox and Sent Items", "")
    End If

    ' close anything left half-finished (an unfinished email is discarded, not saved)
    If Not outlookMail Is Nothing Then outlookMail.Close 1 'olDiscard
    Set outlookMail = Nothing
    If Not singleDoc Is Nothing Then singleDoc.Close False
    Set singleDoc = Nothing

    ' restore the original mail merge settings
    If stateSaved Then
        mm.DataSource.FirstRecord = origFirst
        mm.DataSource.LastRecord = origLast
        mm.Destination = origDest
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = ""

    ' close the run in the log
    If runStarted Then
        If errNum <> 0 Then
            runResult = "RUN END - ERROR"
        ElseIf userCancelled Then
            runResult = "RUN END - STOPPED BY USER"
        ElseIf reachedEnd Then
            runResult = "RUN END - COMPLETE"
        Else
            runResult = "RUN END - BATCH LIMIT"
        End If
        WriteLog 0, runResult, "", "", "Sent: " & sentCount & "; drafts: " & draftCount & _
                 "; error drafts: " & errorCount & "; last completed record: " & lastDoneRecord
    End If
    On Error GoTo 0

    If errNum <> 0 Then
        MsgBox "The macro stopped because of an error " & _
               IIf(currentRecord > 0, "at record " & currentRecord & " (while " & stage & ")", "before any emails were created") & "." & vbLf & vbLf & _
               "Error " & errNum & ": " & errDesc & _
               IIf(stage = "sending the email", vbLf & vbLf & "This email MAY have been sent - check your Outbox and Sent Items.", ""), _
               vbCritical + MB_FRONT, "Mail merge stopped"
    End If

    ' summary (only if the run actually started)
    If runStarted Then
        msg = "Sent: " & sentCount & vbLf & _
              "Saved as drafts: " & draftCount & vbLf & _
              "Saved as error drafts: " & errorCount & vbLf & vbLf
        If lastDoneRecord > 0 Then
            msg = msg & "Last completed record: " & lastDoneRecord & vbLf
        Else
            msg = msg & "No records were completed." & vbLf
        End If
        If reachedEnd And errNum = 0 And Not userCancelled Then
            msg = msg & "All records have been processed."
        Else
            msg = msg & "Not finished - next time, run the macro again and it will suggest starting at record " & _
                  IIf(lastDoneRecord > 0, lastDoneRecord + 1, firstToProcess) & "."
        End If
        If mLogEnabled And Not mLogFailed Then
            msg = msg & vbLf & vbLf & "Progress log: " & mLogPath
        ElseIf mLogFailed Then
            msg = msg & vbLf & vbLf & "WARNING: some entries could not be written to the progress log (" & mLogPath & _
                  "). Check Sent Items / Drafts to confirm progress."
        Else
            msg = msg & vbLf & vbLf & "No progress log was written for this run."
        End If
        If sentCount > 0 And Not outlookWasRunning Then
            msg = msg & vbLf & vbLf & "Outlook was not open when the macro started. " & _
                  "Open Outlook so the emails waiting in your Outbox are sent."
        End If
        MsgBox msg, vbInformation + MB_FRONT, "Mail merge finished"
    End If
End Sub

Sub ResetMailMergeSettings()
    ' Use this ONLY if the main macro was force-stopped (e.g. Ctrl+Break then "End"),
    ' which skips the automatic clean-up. It sets the merge back to all records.
    ' (The progress log is still accurate up to the last email logged.)
    With ActiveDocument.MailMerge
        If .State = wdMainAndDataSource Then
            .DataSource.FirstRecord = wdDefaultFirstRecord
            .DataSource.LastRecord = wdDefaultLastRecord
            .Destination = wdSendToNewDocument
        End If
    End With
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Mail merge settings reset to all records.", vbInformation, "Reset complete"
End Sub

Private Sub PollStopKey()
    ' records an Esc press; checked several times per email so a held key is not missed
    DoEvents
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then mStopRequested = True
End Sub

Private Function GetLogPath(ByVal doc As Document) As String
    ' log goes next to the main document; if the document is unsaved or on OneDrive/SharePoint
    ' (web address), use the Documents folder instead
    Dim folder As String
    Dim baseName As String
    folder = doc.Path
    If folder = "" Or LCase(Left(folder, 4)) = "http" Then
        On Error Resume Next
        folder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
        On Error GoTo 0
        If folder = "" Or LCase(Left(folder, 4)) = "http" Then folder = Environ("USERPROFILE") & "\Documents"
    End If
    baseName = doc.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    GetLogPath = folder & "\" & baseName & "_MailMergeLog.csv"
End Function

Private Function WriteLog(ByVal recordNum As Long, ByVal status As String, ByVal toAddr As String, _
                          ByVal subj As String, ByVal details As String) As Boolean
    ' appends one line to the progress log and closes the file straight away,
    ' so the log stays accurate even if Word stops unexpectedly
    Dim fNum As Integer
    Dim isNew As Boolean
    If Not mLogEnabled Then Exit Function
    On Error GoTo Failed
    isNew = (Dir(mLogPath) = "")
    fNum = FreeFile
    Open mLogPath For Append As #fNum
    If isNew Then Print #fNum, "Timestamp,Record,Status,Data row (if list not sorted/filtered),To,Subject,Details"
    Print #fNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & "," & _
                 IIf(recordNum > 0, CStr(recordNum), "") & "," & _
                 status & "," & _
                 IIf(recordNum > 0, CStr(recordNum + 1), "") & "," & _
                 CsvField(toAddr) & "," & CsvField(subj) & "," & CsvField(details)
    Close #fNum
    WriteLog = True
    Exit Function
Failed:
    mLogFailed = True
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    WriteLog = False
End Function

Private Function LastCompletedRecord(ByVal logPath As String) As Long
    ' reads the progress log and returns the last record that was sent / saved,
    ' or 0 if there is no log or the last run completed all records
    Dim fNum As Integer
    Dim ln As String
    Dim parts() As String
    Dim lastDone As Long
    On Error GoTo Done
    If Dir(logPath) = "" Then Exit Function
    fNum = FreeFile
    Open logPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, ln
        parts = Split(ln, ",")
        If UBound(parts) >= 2 Then
            Select Case parts(2)
                Case "SENT", "DRAFT", "ERROR DRAFT"
                    If IsNumeric(parts(1)) Then lastDone = CLng(parts(1))
                Case "RUN END - COMPLETE"
                    lastDone = 0
            End Select
        End If
    Loop
Done:
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    LastCompletedRecord = lastDone
End Function

Private Function CsvField(ByVal s As String) As String
    ' wraps a value in quotes for CSV, doubling any quotes and removing line breaks
    CsvField = """" & Replace(Replace(Replace(s, """", """"""), vbCr, " "), vbLf, " ") & """"
End Function

Private Function StripToLcaseLetters(ByVal inputString As String) As String
    ' strips out all non-letters and converts to lowercase
    Dim i As Long
    Dim s As String
    For i = 1 To Len(inputString)
        Select Case Asc(Mid(inputString, i, 1))
            Case 65 To 90, 97 To 122
                s = s & Mid(inputString, i, 1)
        End Select
    Next i
    StripToLcaseLetters = LCase(s)
End Function
