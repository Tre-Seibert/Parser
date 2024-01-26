Attribute VB_Name = "HELP04_MakeTicket_Time_1W"
'############################################################################################################
'## HELP03_MakeTicket_Time - VBA Code to create a Time appointment item                                    1W
'############################################################################################################

Option Explicit

'-- *** HELP00_Public = Module must be loaded to use Public Constants and Variables ***

'--------------------------------------------------------------------------------------------------
' Help_MakeTicket_Time - Makes the Ticket Time Entry (Appointment) by a Time submission
'--------------------------------------------------------------------------------------------------
Public Function Help_MakeTicket_Time(ByRef oAppt As AppointmentItem, _
                                     ByRef sClient As String, _
                                     ByRef sTicketNum As String, _
                                     ByRef sTopic As String, _
                                     ByRef sSubject As String, _
                                     ByRef sSender As String, _
                                     ByRef sMsgClass As String, _
                                     ByRef fTarget As Folder) As AppointmentItem
   
    Dim ii              As Integer
    Dim sInitials       As String
    Dim dteWorkDate     As Date
    Dim dblHours        As Double
    Dim sInvDesc        As String
    Dim oApptCopy       As AppointmentItem
    Dim lBodyLen        As Long
    Dim jdx             As Integer  '-- used to prevent infinite looping while remove blank trailing lines
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set Help_MakeTicket_Time = Nothing
    
    '--------------------------------------------------------------------------------
    '-- Create a Ticket
    '--------------------------------------------------------------------------------
    '-- Change the Message Class so the custom fields will show and the correct forms
    '-- are use for Replies, Forwards and Posts
    oAppt.MessageClass = sMsgClass  '-- maybe ineffective!!!
    
    '-- Add new fields
    oAppt.UserProperties.Add TIME_BILLABLE, olYesNo
    oAppt.UserProperties.Add TIME_BILLEND, olDateTime
    oAppt.UserProperties.Add TIME_BILLHOURS, olNumber
    oAppt.UserProperties.Add TIME_BILLSTART, olDateTime
    oAppt.UserProperties.Add TIME_CLIENT, olText
    oAppt.UserProperties.Add TIME_DATE_CREATED, olDateTime
    oAppt.UserProperties.Add TIME_GRATIS, olYesNo
    oAppt.UserProperties.Add TIME_HOURS, olNumber
    oAppt.UserProperties.Add TIME_INVOICE_DESC, olText
    oAppt.UserProperties.Add TIME_INVOICE_NUM, olText
    oAppt.UserProperties.Add TIME_JOB, olText
    oAppt.UserProperties.Add TIME_REVIEWED, olYesNo
    oAppt.UserProperties.Add TIME_TECH, olText
    oAppt.UserProperties.Add TIME_TICKETNUM, olText
    oAppt.UserProperties.Add TIME_TOPIC, olText
    oAppt.UserProperties.Add TIME_WORKDATE, olDateTime
    oAppt.UserProperties.Add TIME_UniqueID, olText

    '-- Initialize new fields
    oAppt.UserProperties(TIME_BILLABLE) = False
    oAppt.UserProperties(TIME_BILLEND) = NODATE
    oAppt.UserProperties(TIME_BILLHOURS) = 0
    oAppt.UserProperties(TIME_BILLSTART) = NODATE
    oAppt.UserProperties(TIME_CLIENT) = UCase(sClient)
    oAppt.UserProperties(TIME_DATE_CREATED) = Now()
    oAppt.UserProperties(TIME_GRATIS) = False
    oAppt.UserProperties(TIME_HOURS) = 0
    oAppt.UserProperties(TIME_INVOICE_DESC) = ""
    oAppt.UserProperties(TIME_INVOICE_NUM) = ""
    oAppt.UserProperties(TIME_JOB) = ""
    oAppt.UserProperties(TIME_REVIEWED) = False
    oAppt.UserProperties(TIME_TECH) = ""
    oAppt.UserProperties(TIME_TICKETNUM) = sTicketNum
    oAppt.UserProperties(TIME_TOPIC) = Trim(sTopic)
    oAppt.UserProperties(TIME_WORKDATE) = NODATE
    oAppt.UserProperties(TIME_UniqueID) = oAppt.EntryID

    '-- Reset importance to Normal even though the incoming email is set to high
    oAppt.Importance = olImportanceNormal
    oAppt.ReminderSet = False
    oAppt.ReminderOverrideDefault = False
    
    '--------------------------------------------------------------------------------
    '-- Get the Tech's initials
    sInitials = TableFind(m_asAssignees, ASSIGNEE_NAME, sSender, ASSIGNEE_INITIALS)
    
    If sInitials = "" Then
        sInitials = "??"
    End If
    
    oAppt.UserProperties(TIME_TECH) = sInitials

    '--------------------------------------------------------------------------------
    '- Create and load the .Work Date with just the Date
    dteWorkDate = DateSerial(Year(oAppt.Start), Month(oAppt.Start), Day(oAppt.Start))
    oAppt.UserProperties(TIME_WORKDATE) = dteWorkDate
    
    '--------------------------------------------------------------------------------
    '- Create and load the Duration in Hours
    dblHours = Round(oAppt.Duration / 60, 3)
    oAppt.UserProperties(TIME_HOURS) = dblHours
    
    oAppt.UserProperties(TIME_BILLSTART) = Round15Minutes(oAppt.Start)
    oAppt.UserProperties(TIME_BILLEND) = Round15Minutes(oAppt.End)
    If DateDiff("n", oAppt.UserProperties(TIME_BILLSTART), oAppt.UserProperties(TIME_BILLEND)) = 0 Then
        oAppt.UserProperties(TIME_BILLEND) = DateAdd("n", 15, oAppt.UserProperties(TIME_BILLEND))
    End If
    oAppt.UserProperties(TIME_BILLHOURS) = DateDiff("n", oAppt.UserProperties(TIME_BILLSTART), oAppt.UserProperties(TIME_BILLEND)) / 60
    
    '--------------------------------------------------------------------------------
    '-- Create the Invoice Description
    sInvDesc = Format(dteWorkDate, "mm/dd/yyyy") & " " & _
               Format(oAppt.UserProperties(TIME_BILLSTART), "h:mm AM/PM") & " - " & _
               Format(oAppt.UserProperties(TIME_BILLEND), "h:mm AM/PM") & " " & _
               "(" & sInitials & ") " & _
               Trim(sTopic)
    oAppt.UserProperties(TIME_INVOICE_DESC) = sInvDesc
   
    '--------------------------------------------------------------------------------
    '-- Clean Up the Body
    If Len(oAppt.Body) > 0 Then
        oAppt.Body = Trim(oAppt.Body)
        
        '-- Get rid of the header for the appointment
        ii = InStr(1, oAppt.Body, "*~*~*~*~*~*~*~*~*~*", vbTextCompare)
        If ii > 0 Then
            ii = ii + Len("*~*~*~*~*~*~*~*~*~*") + 2
            oAppt.Body = Trim(Mid(oAppt.Body, ii, Len(oAppt.Body) - ii + 1))
        End If
        
        '-- Get rid of preceding blank lines
        Do While Left(oAppt.Body, 2) = vbCrLf
            lBodyLen = Len(oAppt.Body)
            oAppt.Body = Trim(Mid(oAppt.Body, 3, Len(oAppt.Body) - 2))
            If lBodyLen = Len(oAppt.Body) Then Exit Do
        Loop
        
        '-- Get rid of trailing blank lines
        jdx = 10
        Do While Right(oAppt.Body, 2) = vbCrLf
            lBodyLen = Len(oAppt.Body)
            oAppt.Body = Trim(Mid(oAppt.Body, 1, Len(oAppt.Body) - 2))
            jdx = jdx - 1
            If lBodyLen = Len(oAppt.Body) Then Exit Do
            If jdx = 0 Then Exit Do     '-- Preven infinite looping
        Loop

        
        '-- Replace double blanks lines
        oAppt.Body = Replace(oAppt.Body, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)

    End If
       
    '--------------------------------------------------------------------------------
    '-- Fill in the subject
    oAppt.Subject = TKTDELIM & sClient & TKTDELIM & sTicketNum & TKTDELIM & " (" & sInitials & ") " & sTopic
    
    '-- KLUDGE: Trying to prevent conflicts from occuring in Outlook 2010 Public Folder TB Mail
    oAppt.UnRead = False
    
    '--------------------------------------------------------------------------------
    Set Help_MakeTicket_Time = oAppt
    
    '-- This saves it in HELP Calendar
    oAppt.Save
    
    '-- This copies it to the target Folder (This may go away)
    '-- 2 Prob: 1) put "COPY:" in front of subject 2) can't delete a calendar entry
    Set oApptCopy = oAppt.Copy
    oApptCopy.Move fTarget
    
    '-- Used to cleanup the time entry that gets copied.
    '-- Will eventually be used to update the total time
    CleanUp_Time fTarget
    
    CleanUp_Time m_NS.GetDefaultFolder(olFolderCalendar) '-- Clean up COPY:
   
    Exit Function
ERRORHANDLER:
    HandleError "HELP04_MakeTicket_Time:Help_MakeTicket_Time [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' CleanUp_Time - When an appointment is Copied, Outlook prefixes "Copy:" in the subject, that needs
'                to be removed.
'--------------------------------------------------------------------------------------------------
Sub CleanUp_Time(fTarget As Folder)

    Dim sFilter As String
    Dim citems As Items
    Dim oAppt As AppointmentItem
    Dim idx As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- DASL search string
    sFilter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%Copy:%'"
    Set citems = fTarget.Items.Restrict(sFilter)
    citems.Sort "[Subject]", True
    
    For idx = citems.Count To 1 Step -1
        Set oAppt = citems(idx)
        If Left(oAppt.Subject, 6) = "Copy: " Then
            oAppt.Subject = Trim(Mid(oAppt.Subject, Len("Copy: ") + 1, Len(oAppt.Subject) - Len("Copy: ")))
            oAppt.Save
        End If
        DoEvents
    Next

    Set citems = Nothing
    Set oAppt = Nothing
    
    Exit Sub
ERRORHANDLER:
    HandleError "HELP04_MakeTicket_Time:CleanUp_Time [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Sub


'----------------------------------------------------------------------------------------
' Round15Minutes - Rounds the time to the nearest 15 minute interval
'----------------------------------------------------------------------------------------
'Function Round15Minutes(dteTime As Date) As Date
Function Round15Minutes(dteTime)

    Dim iHours      'As Integer
    Dim iMinutes    'As Integer
    Dim iSeconds    'As Integer
    Dim dteNewTime  'As Date
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Break out the time
    iHours = Hour(dteTime)
    iMinutes = Minute(dteTime)
    iSeconds = Second(dteTime)
    dteNewTime = dteTime

    '-- If the minutes are not at the 15min intervals, then round to the nearest 15min interval
    If iMinutes <> 0 Or iMinutes <> 15 Or iMinutes <> 30 Or iMinutes <> 45 Then
    
        '-- First remove the Seconds and Minutes to get just the hour
        dteNewTime = DateAdd("s", -1 * iSeconds, dteNewTime)
        dteNewTime = DateAdd("n", -1 * iMinutes, dteNewTime)
        
        '-- Round to the nearest 15min interval
        iMinutes = Int((iMinutes + 7) / 15) * 15
        
        '-- Add the Minutes to the hour
        dteNewTime = DateAdd("n", iMinutes, dteNewTime)
    End If
    
    '-- Return the Rounded time
    Round15Minutes = dteNewTime

    Exit Function
ERRORHANDLER:
    HandleError "HELP04_MakeTicket_Time:Round15Minutes [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function
