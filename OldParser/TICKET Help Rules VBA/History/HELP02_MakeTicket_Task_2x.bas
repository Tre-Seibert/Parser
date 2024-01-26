Attribute VB_Name = "HELP02_MakeTicket_Task_2x"
'############################################################################################################
'## HELP01_MakeTicket_Task - VBA Code to create a Ticket task item                                         2x
'############################################################################################################

Option Explicit

'-- *** HELP00_Public = Module must be loaded to use Public Constants and Variables              ***
'-- *** HELP01_MakeTicket_from_Rule:HELP_MakeTicket_Init must run to initialize global variables ***

'--------------------------------------------------------------------------------------------------
' Help_MakeTicket_Task - Creates a ticket (Task) based on the email the came in.  It puts to the
'                        ticket into the target ticket folder.
'--------------------------------------------------------------------------------------------------
Public Sub Help_MakeTicket_Task(ByRef oMail As MailItem, _
                                     ByRef sClient As String, _
                                     ByRef sTicketNum As String, _
                                     ByRef sTopic As String, _
                                     ByRef sSubject As String, _
                                     ByRef sSender As String, _
                                     ByRef sAssignee As String, _
                                     ByRef sMsgClass As String, _
                                     ByRef fTarget As Folder)
    
    Dim oTask   As TaskItem '-- New Ticket
    Dim oTask2  As TaskItem '-- Copy of new ticket for backup
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '--------------------------------------------------------------------------------
    '-- Create a Ticket
    '--------------------------------------------------------------------------------
    Set oTask = CreateTicket(sClient, sTicketNum, sTopic, sSubject, sSender, sAssignee, _
                              oMail.Body, sMsgClass, oMail.ReceivedTime, oMail.Importance)
                              
    '-------------------------------------------------------------------------------------------
    '-- Change settings for TechBldrs Quote
    If InStr(1, sTopic, "TechBldrs Quote ", vbTextCompare) > 0 Then
        oTask.UserProperties(TKT_REASON) = "Billable"
        oTask.UserProperties(TKT_ASSIGNEE) = "ja"
        oTask.Categories = "Quoted"
    End If
    
    '-------------------------------------------------------------------------------------------
                              
    '-- Save a Backup Ticket in case of random system deletions
    Set oTask2 = oTask.Copy
    oTask2.MessageClass = sMsgClass
    oTask2.Move m_fldrBackupTickets

    '-- Put a new ticket into the Ticket folder.  Move deletes the Task Item
    oTask.UserProperties(TKT_UNIQUEID) = oTask.EntryID '-- $$$ Don't know why this doesn't working in CreateTicket
    
    oTask.Move fTarget

    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP02_MakeTicket_Task:Help_MakeTicket_Task [" & Err.Number & "] " & Err.Description & " >" & oMail.Subject
    If ERR_RESUME Then Resume Next
End Sub


'--------------------------------------------------------------------------------------------------
'-- CreateTicket - Creates the Ticket but does not save or move it. Used by other routines to
'--                create a ticket (e.g. Orphan cleanup)
'--------------------------------------------------------------------------------------------------
Public Function CreateTicket(ByRef sClient As String, _
                             ByRef sTicketNum As String, _
                             ByRef sTopic As String, _
                             ByRef sSubject As String, _
                             ByRef sSender As String, _
                             ByRef sAssignee As String, _
                             ByRef sBody As String, _
                             ByRef sMsgClass As String, _
                             ByRef dteSentOrReceived As Date, _
                             ByRef iImportance As Integer) As TaskItem
    
    Dim oTask           As TaskItem
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set CreateTicket = Nothing
    '--------------------------------------------------------------------------------
    '-- Create a Ticket
    '--------------------------------------------------------------------------------
    Set oTask = Application.CreateItem(olTaskItem)
    
    '-- Change the Message Class so the custom fields will show and the correct forms
    '-- are use for Replies, Forwards and Posts
    oTask.MessageClass = sMsgClass  '-- maybe ineffective!!!
    
    '-- Add new fields
    oTask.UserProperties.Add TKT_ACTION, olText
    oTask.UserProperties.Add TKT_ASSIGNEE, olText
    oTask.UserProperties.Add TKT_CAUSE1, olText
    oTask.UserProperties.Add TKT_CLIENT, olText
    oTask.UserProperties.Add TKT_DATE_CREATED, olDateTime
    oTask.UserProperties.Add TKT_DATE_FIRST_TOUCH, olDateTime
    oTask.UserProperties.Add TKT_DATE_LAST_ACTIVITY, olDateTime
    oTask.UserProperties.Add TKT_DATE_MODIFIED, olDateTime
    oTask.UserProperties.Add TKT_HRS_ACTUAL_TOTAL, olNumber
    oTask.UserProperties.Add TKT_HRS_BILLABLE_TOTAL, olNumber
    oTask.UserProperties.Add TKT_HRS_COMPLETION, olNumber
    oTask.UserProperties.Add TKT_HRS_DURATION, olNumber
    oTask.UserProperties.Add TKT_HRS_GRATIS_TOTAL, olNumber
    oTask.UserProperties.Add TKT_INVOICE_NOTES, olText
    oTask.UserProperties.Add TKT_INVOICE_NUM, olText
    oTask.UserProperties.Add TKT_JOB, olText
    oTask.UserProperties.Add TKT_LOG, olText
    oTask.UserProperties.Add TKT_MACHINE_NAME, olText
    oTask.UserProperties.Add TKT_MACHINE_SUPPORT, olText
    oTask.UserProperties.Add TKT_MAIL_TEMPLATE, olText
    oTask.UserProperties.Add TKT_REASON, olText
    oTask.UserProperties.Add TKT_REQUESTOR, olText
    oTask.UserProperties.Add TKT_STATUS, olText
    oTask.UserProperties.Add TKT_TECHNAME, olText
    oTask.UserProperties.Add TKT_TICKETMONTH, olNumber
    oTask.UserProperties.Add TKT_TICKETNUM, olText
    oTask.UserProperties.Add TKT_TICKETYEAR, olNumber
    oTask.UserProperties.Add TKT_TOPIC, olText
    oTask.UserProperties.Add TKT_UNIQUEID, olText
    oTask.UserProperties.Add TKT_USER, olText

    oTask.UserProperties(TKT_ACTION) = ""
    oTask.UserProperties(TKT_ASSIGNEE) = sAssignee
    oTask.UserProperties(TKT_CAUSE1) = ""
    oTask.UserProperties(TKT_CLIENT) = sClient
    oTask.UserProperties(TKT_DATE_CREATED) = DateValue(dteSentOrReceived) + TimeSerial(Hour(dteSentOrReceived), Minute(dteSentOrReceived), 0)
    oTask.UserProperties(TKT_DATE_FIRST_TOUCH) = NODATE
    oTask.UserProperties(TKT_DATE_LAST_ACTIVITY) = oTask.UserProperties(TKT_DATE_CREATED)
    oTask.UserProperties(TKT_DATE_MODIFIED) = NODATE
    oTask.UserProperties(TKT_HRS_ACTUAL_TOTAL) = 0
    oTask.UserProperties(TKT_HRS_BILLABLE_TOTAL) = 0
    oTask.UserProperties(TKT_HRS_COMPLETION) = 0
    oTask.UserProperties(TKT_HRS_DURATION) = 0
    oTask.UserProperties(TKT_HRS_GRATIS_TOTAL) = 0
    oTask.UserProperties(TKT_INVOICE_NOTES) = ""
    oTask.UserProperties(TKT_INVOICE_NUM) = ""
    oTask.UserProperties(TKT_JOB) = ""
    oTask.UserProperties(TKT_LOG) = "Created: " & oTask.UserProperties(TKT_DATE_CREATED)
    oTask.UserProperties(TKT_MACHINE_NAME) = ""
    oTask.UserProperties(TKT_MACHINE_SUPPORT) = ""
    oTask.UserProperties(TKT_MAIL_TEMPLATE) = ""
    oTask.UserProperties(TKT_REASON) = ""
    oTask.UserProperties(TKT_REQUESTOR) = sSender
    oTask.UserProperties(TKT_STATUS) = TKT_STATUS_NEW
    oTask.UserProperties(TKT_TECHNAME) = ""
    oTask.UserProperties(TKT_TICKETMONTH) = Month(Now)
    oTask.UserProperties(TKT_TICKETNUM) = sTicketNum
    oTask.UserProperties(TKT_TICKETYEAR) = Year(Now)
    oTask.UserProperties(TKT_TOPIC) = sTopic
    oTask.UserProperties(TKT_UNIQUEID) = oTask.EntryID
    oTask.UserProperties(TKT_USER) = ""
    
    '--------------------------------------------------------------------------------
    '-- Set the ticket category to Urgent, if the mail importance is set high
    If iImportance = olImportanceHigh Then
        oTask.Categories = TKT_CAT0_URGENT
    ElseIf InStr(1, sSubject, ALERT_Backup, vbTextCompare) > 0 Or InStr(1, sSubject, ALERT_Backup2, vbTextCompare) > 0 Then
        oTask.Categories = TKT_CAT4_BACKUP
        oTask.UserProperties(TKT_REASON) = TKT_REASON_ALERT
    Else
        oTask.Categories = TKT_CAT2_NORMAL
    End If
    
    If InStr(1, sSubject, ALERT_Tag) > 0 Then
        oTask.UserProperties(TKT_REASON) = TKT_REASON_ALERT
        oTask.Categories = TKT_CAT0_URGENT
    End If
    
    '--------------------------------------------------------------------------------
    oTask.DueDate = AssignDueDate(dteSentOrReceived)
        
    '--------------------------------------------------------------------------------
    '-- Copy the initial email body into the Notes section and replace double blanks lines
    oTask.Body = Replace(sBody, vbCrLf & vbCrLf, vbCrLf, , , vbBinaryCompare)
    
    '--------------------------------------------------------------------------------
    '-- Fill in the subject
    oTask.Subject = sSubject
    
    '-- Turn off Reminders
    oTask.ReminderSet = False
    oTask.ReminderOverrideDefault = False
    oTask.ReminderTime = NODATE
    
    '-- KLUDGE: Trying to prevent conflicts from occuring in Outlook 2010 Public Folder TB Mail
    oTask.UnRead = False
    '--------------------------------------------------------------------------------
    Set CreateTicket = oTask
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP02_MakeTicket_Task:CreateTicket [" & Err.Number & "] " & Err.Description & " >" & sSubject
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- AssignDueDate - Assign the Due Date by adding RESPONSE_DAYS to the date the ticket was sent.
'--                   This routine does not know of any holidays
'--     returns: due date
'--------------------------------------------------------------------------------------------------
Private Function AssignDueDate(ByRef dteSentOn As Date) As Date
    Dim iSentOnDay      As Integer
    Dim iDueOnDay       As Integer

    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- For 2 day Response, if the Start Date is Sunday, the Due Date is Wednesday
    iSentOnDay = Weekday(dteSentOn)
    AssignDueDate = DateAdd("d", RESPONSE_DAYS, dteSentOn)
    iDueOnDay = Weekday(AssignDueDate)

    '-- If the due date is a Saturday or Sunday, then make it the next Monday.
    If iDueOnDay = vbSaturday Then
        AssignDueDate = DateAdd("d", 2, AssignDueDate)
    ElseIf iDueOnDay = vbSunday Then
        AssignDueDate = DateAdd("d", 1, AssignDueDate)
    End If
    
    '-- If the date it was sent in was a weekend, the add days for weekend days
    If iSentOnDay = vbFriday Or iSentOnDay = vbSaturday Then
        AssignDueDate = DateAdd("d", 2, AssignDueDate)
    ElseIf iSentOnDay = vbSunday Then
        AssignDueDate = DateAdd("d", 1, AssignDueDate)
    End If
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP02_MakeTicket_Task:AssignDueDate [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' Update_Ticket - Updates all the flags, dates and status of a ticket.
'                 Called from HELP_ProcessEmail in the HELP00_MakeTicket_from_Rule module.
'--------------------------------------------------------------------------------------------------
Public Sub Update_Ticket(ByRef oTask As TaskItem, ByRef oMail As MailItem)
    Dim oLock As TaskItem
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Look for lock
    Set oLock = Nothing
    Set oLock = Lock_Find(m_fldrLocks.Items, oTask.UserProperties(TKT_CLIENT), oTask.UserProperties(TKT_TICKETNUM))
    
    '-- Change the last activity date
    If oLock Is Nothing Then
'        oTask.UserProperties(TKT_DATE_LAST_ACTIVITY) = oMail.ReceivedTime
        
        FixDates_Created_LastActivity oTask, m_fldrMail.Items, m_fldrTime.Items, _
                "[" & TKT_CLIENT & "] = '" & oTask.UserProperties(TKT_CLIENT) & "' AND [" & TKT_TICKETNUM & "] = '" & oTask.UserProperties(TKT_TICKETNUM) & "'"
    
    ElseIf Not Lock_InUse(oLock) Then
        If oLock.ReminderTime < oMail.ReceivedTime Or oLock.ReminderTime = NODATE Then
           oTask.UserProperties(TKT_DATE_LAST_ACTIVITY) = oMail.ReceivedTime
        Else
            oTask.UserProperties(TKT_DATE_LAST_ACTIVITY) = oLock.ReminderTime
        End If
        oLock.Delete
        
    '-- update date last activity in lock and nothing else if ticket is in use
    ElseIf Lock_InUse(oLock) Then
        If oLock.ReminderTime < oMail.ReceivedTime Or oLock.ReminderTime = NODATE Then
            oLock.ReminderTime = oMail.ReceivedTime
            oLock.Save
        End If
        Exit Sub
    End If
    
        
        '-- Change the status based on who is sending the email
        If InStr(1, oMail.SenderEmailAddress, TB1, vbTextCompare) > 0 _
           Or InStr(1, oMail.SenderEmailAddress, TB2, vbTextCompare) > 0 _
           Or InStr(1, oMail.SenderEmailAddress, TB3, vbTextCompare) > 0 Then
            oTask.UserProperties(TKT_STATUS) = TKT_STATUS_TO_CLIENT
        Else
            oTask.UserProperties(TKT_STATUS) = TKT_STATUS_FROM_CLIENT
            
            '-- 04/14/12 - If an email comes in after the Ticket set to anything lower than 3 Normal,
            '--            change the Category to Re-opened and uncheck the Complete check box.
            If oTask.Categories <> TKT_CAT0_URGENT And _
               oTask.Categories <> TKT_CAT1_HIGH And _
               oTask.Categories <> TKT_CAT2_NORMAL Then
                oTask.Body = Trim(Date & " was " & oTask.Categories & vbCrLf & oTask.Body)
                oTask.Categories = TKT_CAT1_REOPENED
            End If
            
            If oTask.Complete Then
                oTask.Complete = False
            End If
        
        End If
        
        '-- Set the ticket category to Urgent, if the mail importance is set high
        If oMail.Importance = olImportanceHigh Then
            oTask.Categories = TKT_CAT0_URGENT
        End If
        'oTask.Importance = oMail.Importance  '-- Ticket no longer uses importance
        
        oTask.Save
    
    Set oLock = Nothing
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP02_MakeTicket_Task:Update_Ticket [" & Err.Number & "] " & Err.Description & " >" & oMail.Subject
    If ERR_RESUME Then Resume Next
End Sub


'--------------------------------------------------------------------------------------------------
' Lock_InUse - Checks if ticket is currently open
'--------------------------------------------------------------------------------------------------
Private Function Lock_InUse(ByRef oLock As TaskItem) As Boolean

    Lock_InUse = False
    
    '-- If a name is in the BillingInformation, the ticket is in use
    If oLock.BillingInformation <> "" Then Lock_InUse = True

End Function


'----------------------------------------------------------------------------------------------------------
'-- Lock_Find - Find lock if it exist in the TB Locks folder
'--                         '-- Subject - Client|TicketNum
'--                         '-- BillingInformation - User
'--                         '-- ReminderTime - DateLastActivity
'----------------------------------------------------------------------------------------------------------
Private Function Lock_Find(ByRef cLocks As Items, ByRef sClient As String, ByRef sTicketNum As String) As TaskItem
    
    Dim sFilter         'As String
        
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set Lock_Find = Nothing
    
    '-- Abort if blanks are sent in
    If sClient = "" Or sTicketNum = "" Then Exit Function
    
    '-- Must match the client and ticket number
    sFilter = "[Subject] = '" & sClient & TKTDELIM & sTicketNum & "'"
    
    '-- Search for the Lock
    On Error Resume Next
    Set Lock_Find = cLocks.Find(sFilter)
    On Error GoTo 0

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP02_MakeTicket_Task:Lock_Find [" & Err.Number & "] " & Err.Description & " >" & Lock_Find.Subject
    If ERR_RESUME Then Resume Next
End Function


'----------------------------------------------------------------------------------------------------------
'-- FixDates_Created_LastActivity - Corrects the the creation date and the date of last activity
'----------------------------------------------------------------------------------------------------------
Private Sub FixDates_Created_LastActivity(ByRef oTicket As TaskItem, ByRef cMails As Items, ByRef cTimes As Items, ByRef sFilters As String)
    Dim oTime   As AppointmentItem
    Dim oMail   As MailItem
    
    Dim dteEarliest 'As Date
    Dim dteLatest   'As Date
    
    dteLatest = NODATE
    dteEarliest = NODATE
        
    '-- Get first and last Time Entry
    Set oTime = Nothing
    If cTimes.Count > 0 Then
    
        cTimes.Sort "[Start]", False
        Set oTime = cTimes.Item(1)
        If dteEarliest = NODATE Or dteEarliest > oTime.Start Then dteEarliest = oTime.Start
        
        cTimes.Sort "[Start]", True
        Set oTime = cTimes.Item(1)
        If dteLatest = NODATE Or dteLatest < oTime.Start Then dteLatest = oTime.Start
        
    End If
    
    '-- Get first and last mail entry
    Set oMail = Nothing
    If cMails.Count > 0 Then
        
        cMails.Sort "[Received]", False
        Set oMail = cMails.Item(1)
        
        If dteEarliest = NODATE Or dteEarliest > oMail.ReceivedTime Then dteEarliest = oMail.ReceivedTime
        
        cMails.Sort "[Received]", True
        Set oMail = cMails.Item(1)
        
        If dteLatest = NODATE Or dteLatest < oMail.ReceivedTime Then dteLatest = oMail.ReceivedTime
         
    End If
    
    '-- Check against Ticket creation (this could have changed when tickets are moved or merged); It does not matter for the DLA.
    If dteEarliest = NODATE Or dteEarliest > oTicket.CreationTime Then dteEarliest = oTicket.CreationTime
    If dteEarliest = NODATE Or dteEarliest > oTicket.UserProperties(TKT_DATE_CREATED) Then dteEarliest = oTicket.UserProperties(TKT_DATE_CREATED)
    
    '-- Update the date variables.  This gets updated during the next ticket update.
    If dteEarliest <> NODATE Then
    
        '-- Remove the seconds
        dteEarliest = DateValue(dteEarliest) + TimeSerial(Hour(dteEarliest), Minute(dteEarliest), 0)
            
        '-- This will remove any tickets created by hand and not through the parser
        If oTicket.UserProperties(TKT_DATE_CREATED) <> dteEarliest Then oTicket.UserProperties(TKT_DATE_CREATED) = dteEarliest
    End If
    
    If dteLatest <> NODATE Then
    
        '-- Remove the seconds
        dteLatest = DateValue(dteLatest) + TimeSerial(Hour(dteLatest), Minute(dteLatest), 0)
            
        If oTicket.UserProperties(TKT_DATE_LAST_ACTIVITY) <> dteLatest Then oTicket.UserProperties(TKT_DATE_LAST_ACTIVITY) = dteLatest
    End If

    Set oTime = Nothing
    Set oMail = Nothing
End Sub

