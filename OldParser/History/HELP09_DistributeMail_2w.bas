Attribute VB_Name = "HELP09_DistributeMail_2w"
'############################################################################################################
'## HELP09_DistributeMail - VBA Code to distribute incoming notification to the Help Desk and Techs        2s
'##
'##   1) m_fldrCalendar - Must be set to the calendar that has the Help Desk and Duty Week times
'##
'##   2) The calendar can have the following entries in the Folder Properties Description to override the
'##      default settings
'##         starttime=7:30am    - used to determine
'##         endtime=4:00pm      -   the work hours
'##         HDlabel=Help Desk   - label for the weekly Help Desk calendar entry
'##         DWlabel=Duty Week   - label for the weekly Duty Week calendar entry
'##
'############################################################################################################

Option Explicit

'-- Settings for processing Email / Ticket Ditribution (Default Settings)
Const DUTYWEEK_LABEL = "Duty Week"
Const HELPDESK_LABEL = "Help Desk"
Const HELPDESK_STARTTIME = "7:30:00 AM"
Const HELPDESK_ENDTIME = "4:00:00 PM"

'-- Tags for Settings stored in the Folder's Description field
Const FLDPROP_HDLABEL = "HDlabel"
Const FLDPROP_DWLABEL = "DWlabel"
Const FLDPROP_STARTTIME = "starttime"
Const FLDPROP_ENDTIME = "endtime"

'-- Variables to store the settings
Dim m_sDWLabel                  As String
Dim m_sHDLabel                  As String
Dim m_dteHDStartTime            As Date
Dim m_dteHDEndTime              As Date
Dim m_blnDistributeMail_Init    As Boolean


Private Sub mailtest()

    Dim oMail As MailItem
    
    Set oMail = Application.CreateItem(olMailItem)
    HELP_MakeTicket_Init
    HELP_DistributeMail "", oMail
End Sub

'--------------------------------------------------------------------------------------------------
'-- Help_DistributeMail - Distribute the Ticket by Mail to all the Techs
'
'-- Rule 1: If the current time is within the schedule of the assigned tech, then only send to them.
'-- Rule 2: Otherwise, send to everyone that is attached to the HELP Outlook Profile
'--------------------------------------------------------------------------------------------------
Public Sub HELP_DistributeMail(ByRef sAssignee As String, _
                               ByRef oMail As MailItem)

    Dim blnIsItWorkTime As Boolean
    Dim blnNoHelpTechs  As Boolean
    Dim blnNoDutyTechs  As Boolean
    Dim blnDistSuccess  As Boolean
    Dim sHDTechs        As String
    Dim sDWTechs        As String
    Dim sReason         As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If

    '-- Initialize the module variables
    If Not m_blnDistributeMail_Init Then
        m_blnDistributeMail_Init = DistributeMail_Init
    End If
    
    blnIsItWorkTime = False
    blnNoHelpTechs = False
    blnNoDutyTechs = False
    blnDistSuccess = True
    sReason = ""
    
    '--1 Is it after hours
    If m_dteHDStartTime <= Time And Time <= m_dteHDEndTime And Weekday(Now) <> vbSaturday And Weekday(Now) <> vbSunday Then
        blnIsItWorkTime = True
    End If
    
    '--2 If it's work time, then there must be help desk techs, or send it to the duty week person
    If blnIsItWorkTime Then
        sHDTechs = FindAssignedTechs(m_fldrCalendar, m_sHDLabel)
        
        '-- If no help desk techs, set a flag
        If sHDTechs = "" Then
            blnNoHelpTechs = True
        End If
    End If
    
    '--3 If it's after hours or there is no help desk tech, then send the email to the Duty Week Tech
    If Not blnIsItWorkTime Or blnNoHelpTechs Then
        sDWTechs = FindAssignedTechs(m_fldrCalendar, m_sDWLabel)
        
        '-- If no Duty Week tech, then set a flag
        If sDWTechs = "" Then
            blnNoDutyTechs = True
    
        '-- Copy email to the Duty Week Techs
        Else
            If blnNoHelpTechs Then sReason = "NoHelpTech: "
            If Not blnIsItWorkTime Then sReason = "DutyWeek: "
            
            blnDistSuccess = SendMailtoTechs(oMail, sDWTechs, sReason)
            
        End If
    End If
    
    '--4 If no one is assigned, then send to everyone
    If blnNoHelpTechs Or blnNoDutyTechs Or Not blnDistSuccess Then
    
        sReason = "NoAssignedTech:"
        If Not blnDistSuccess Then sReason = "MailErr2Techs: "
        
        SendMailtoAllTechs oMail, sReason
    End If
    
    
    '-- Always copy mail to the person assigned to the ticket
    SendMailtoAssignee sAssignee, oMail, "ClientEmail: "
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:HELP_DistributeMail [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Sub


'--------------------------------------------------------------------------------------------------
'-- DistributeMail_Init - Initalize for this module
'--------------------------------------------------------------------------------------------------
Private Function DistributeMail_Init() As Boolean
    
    Dim arrSettings()   As String
    Dim idx             As Integer
    Dim arrOneSetting() As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
   
    '-- Set the defaults
    m_dteHDStartTime = HELPDESK_STARTTIME
    m_dteHDEndTime = HELPDESK_ENDTIME
    m_sHDLabel = HELPDESK_LABEL
    m_sDWLabel = DUTYWEEK_LABEL
    
    '-- *** m_fldrCalendar is set in HELP_MakeTicket_Init in HELP00_MakeTicket_from_Rule
    '-- If there is a description in the folder, then proceed to override the defaults
    '-- eg. starttime=7:30am
    '--     endtime=4:00pm
    '--     HDlabel=Help Desk
    '--     DWlabel=Duty Week

    If m_fldrCalendar.Description <> "" Then
        
        '-- Split the settings into separate strings
        arrSettings = Split(m_fldrCalendar.Description, vbCrLf)
    
        '-- Split each setting and change the value
        For idx = 0 To UBound(arrSettings)
            arrOneSetting = Split(arrSettings(idx), "=")
            
            '-- If there was a value supplied
            If UBound(arrOneSetting) > 0 Then
            
                '-- Change this list of Cases if there are more items to suck in
                Select Case LCase(Trim(arrOneSetting(0)))
                Case FLDPROP_STARTTIME
                    m_dteHDStartTime = Trim(arrOneSetting(1))
                Case FLDPROP_ENDTIME
                    m_dteHDEndTime = Trim(arrOneSetting(1))
                Case FLDPROP_HDLABEL
                    m_sHDLabel = Trim(arrOneSetting(1))
                Case FLDPROP_DWLABEL
                    m_sDWLabel = Trim(arrOneSetting(1))
                End Select
            End If
        Next
    End If

    '-- Set Initialized variable so everyone knows we have been initilized
    DistributeMail_Init = True

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:DistributeMail_Init [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- FindAssignedTechs - Find the techs that are currently assigned to appointment that is started
'--                     with the label ("Help Desk", "Duty Week").  The string after the label is
'--                     supposed to be the list of Tech Initials e.g. "Help Desk (??, ??, ??)"
'--------------------------------------------------------------------------------------------------
Private Function FindAssignedTechs(ByRef fCalendar As Folder, _
                                   ByRef sLabel As String) As String

    Dim oAppt   As AppointmentItem
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    FindAssignedTechs = ""
    
    '-- Find an appointment that has the Default Tech(s) in the Subject
    Set oAppt = FindApptbyLabel(fCalendar, sLabel)
        
    '-- Return the technician initials
    If Not oAppt Is Nothing Then
    
        '-- Get the string after the label
        FindAssignedTechs = Mid(oAppt.Subject, Len(sLabel) + 1, Len(oAppt.Subject) - Len(sLabel))
        
        '-- Remove common separators
        FindAssignedTechs = Replace(FindAssignedTechs, "(", "")
        FindAssignedTechs = Replace(FindAssignedTechs, ")", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "-", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "]", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "[", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "{", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "}", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "/", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "\", "")
        FindAssignedTechs = Replace(FindAssignedTechs, "|", "")
        FindAssignedTechs = Trim(Replace(FindAssignedTechs, " ", ""))
        
    End If
    
    '-- Cleanup
    Set oAppt = Nothing

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:FindAssignedTechs [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- FindApptbyLabel - Find the calender appointment that matches the label
'--------------------------------------------------------------------------------------------------
Private Function FindApptbyLabel(ByRef fCalendar As Folder, _
                                 ByRef sLabel As String) As AppointmentItem

    Dim sLabel2 As String
    Dim sFilter As String
    Dim idx     As Integer
    Dim cAppts  As Items
    Dim iRecurrenceState As Integer
    Dim RecPat  As RecurrencePattern
    Dim oExcpt  As AppointmentItem
    Dim oAppt   As AppointmentItem
    Dim jdx     As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set FindApptbyLabel = Nothing
    
    '-- If the folder does not exists, exit out
    If fCalendar Is Nothing Then Exit Function
            
    '-- Looking for an appointment that starts before and ends after NOW, and
    '-- begins with the same label
    sLabel2 = Left(sLabel, Len(sLabel) - 1) & Chr(Asc(Right(sLabel, 1)) + 1)
    sFilter = "[Start] <= '" & Format(Now(), "mm/dd/yyyy hh:mm AMPM") & "' AND " & _
              "[End] >= '" & Format(Now(), "mm/dd/yyyy hh:mm AMPM") & "' AND " & _
              "[Subject] > '" & sLabel & "' AND " & _
              "[Subject] < '" & sLabel2 & "'"

    '-- Get all appoinments that match
    Set cAppts = fCalendar.Items.Restrict(sFilter)
    
    '-- If you find more than 1, then use the first non-recurring appointment
    If cAppts.Count > 1 Then
        For idx = 1 To cAppts.Count
            On Error Resume Next                            '*** to eliminate failure
            iRecurrenceState = 0
            iRecurrenceState = cAppts.Item(idx).RecurrenceState
            On Error GoTo 0                                 '*** to eliminate failure
            If ERR_HANDLER Then On Error GoTo ERRORHANDLER  '*** to eliminate failure
            If iRecurrenceState = olApptNotRecurring Then
                Set oAppt = cAppts.Item(idx)
            End If
        Next
    End If
    
    '-- If there were non-recurring appointment, then take the 1st recurring appointment
    '-- but check for an occurance of the series.
    If oAppt Is Nothing Then
        
        For idx = 1 To cAppts.Count
            On Error Resume Next                            '*** to eliminate failure
            iRecurrenceState = 0
            iRecurrenceState = cAppts.Item(idx).RecurrenceState
            On Error GoTo 0
            If iRecurrenceState <> olApptNotRecurring Then
                Set oAppt = cAppts.Item(idx)
            
                '-- Check for a occurance of the recurring series
                On Error Resume Next                            '*** to eliminate failure
                Set RecPat = oAppt.GetRecurrencePattern
                On Error GoTo 0                                 '*** to eliminate failure
                If ERR_HANDLER Then On Error GoTo ERRORHANDLER  '*** to eliminate failure
            
                '-- If there is 1 Exceptions, then use this
                If RecPat.Exceptions.Count > 0 Then
                    
                    For jdx = RecPat.Exceptions.Count To 1 Step -1
'                        On Error Resume Next
                        Set oExcpt = RecPat.Exceptions.Item(jdx).AppointmentItem
                        
                        '-- You can have an occurrence that is delete, so use ON ERROR
                        If oExcpt.Start <= Now And Now <= oExcpt.End Then
                            Set oAppt = oExcpt
                            Exit For
                        End If
'                        On Error GoTo 0
                    Next
                End If
                
                Exit For
            End If
        Next
    End If

    Set FindApptbyLabel = oAppt
    
    Set cAppts = Nothing
    Set RecPat = Nothing
    Set oExcpt = Nothing
    Set oAppt = Nothing
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:FindApptbyLabel [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- SendMailtoTechs - Copies the mail to all the tech initials in sTech separate by commas
'--------------------------------------------------------------------------------------------------
Private Function SendMailtoTechs(ByRef oMail As MailItem, _
                                 ByRef sTechs As String, _
                                 ByRef sReason As String) As Boolean
    Dim arrTechs()  As String
    Dim sEmail      As String
    Dim sTechEmails As String
    Dim idx         As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    SendMailtoTechs = False
    
    sTechEmails = ""
            
    arrTechs = Split(sTechs, ",")
    
    For idx = 0 To UBound(arrTechs)
        
        '-- Get Help Desk Techs
        sEmail = CmTableFind(m_asAssignees, ASSIGNEE_INITIALS, Trim(arrTechs(idx)), ASSIGNEE_EMAIL)
        
        '-- If Tech found, then build the tech email string
        If sEmail <> "" And sTechEmails = "" Then
            sTechEmails = sEmail
        Else
            sTechEmails = sTechEmails & "; " & sEmail
        End If
    Next
    
    '-- Send out the email to all the techs in the list
    If sTechEmails <> "" Then
        SendMailtoTechs = SendMail(oMail, sTechEmails, sReason)
    End If

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:SendMailtoTechs [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function



'--------------------------------------------------------------------------------------------------
'-- SendMailtoAllTechs - Copy the mail to every tech (attached mailboxes)
'--------------------------------------------------------------------------------------------------
Private Function SendMailtoAllTechs(ByRef oMail As MailItem, _
                                    ByRef sReason As String) As Boolean
    Dim sTechEmails As String
    Dim idx         As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    SendMailtoAllTechs = False
    
    '-- Get all the emails in the Assignee table.  This will have some blank emails spaces
    sTechEmails = ""
    
    If m_iRowsAssignees > -1 Then
        For idx = 0 To m_iRowsAssignees
            sTechEmails = sTechEmails & ";" & Trim(m_asAssignees(idx, ASSIGNEE_EMAIL))
        Next
        
        If sTechEmails <> "" Then
            SendMailtoAllTechs = SendMail(oMail, sTechEmails, sReason)
        End If
    End If
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:SendMailtoAllTechs [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- SendMailtoAssignee - Copy the mail to the Assignee's inbox
'--------------------------------------------------------------------------------------------------
Private Function SendMailtoAssignee(ByRef sAssignee As String, _
                                    ByRef oMail As MailItem, _
                                    ByRef sReason As String) As Boolean
    Dim sTechEmails As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    SendMailtoAssignee = False
    
    If sAssignee = "" Then Exit Function
    
    '-- Find the matching email
    sTechEmails = CmTableFind(m_asAssignees, ASSIGNEE_INITIALS, sAssignee, ASSIGNEE_EMAIL)
    
    '-- If there is an email, then prepare to send the email
    If sTechEmails <> "" Then
        
        '-- If the sender is NOT the Tech, then send email
        If InStr(1, sTechEmails, oMail.SenderEmailAddress, vbTextCompare) = 0 Then
        
            '-- Send the email
            SendMailtoAssignee = SendMail(oMail, sTechEmails, sReason)
        End If
    End If

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:SendMailtoAssignee [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
'-- SendMail - Send a copy of the mail into an inbox
'--------------------------------------------------------------------------------------------------
Private Function SendMail(ByRef oMail As MailItem, _
                          ByRef sTechEmails As String, _
                          ByRef sReason As String) As Boolean
    Dim oMsg    As MailItem
    Dim oAttach As Attachment
    Dim sPath   As String
    Dim oFSO
    Dim sTempDir
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    SendMail = False
    
    '-- Start a new mail message
    Set oMsg = CreateItem(olMailItem)
    
    '-- Fill in the message
    With oMsg
        .To = sTechEmails
        
        .BillingInformation = oMail.BillingInformation
        .Categories = oMail.Categories
        .Companies = oMail.Companies
        .Importance = oMail.Importance
        .Mileage = oMail.Mileage
        .Sensitivity = oMail.Sensitivity
        .Subject = oMail.Subject
    End With
    
    '-- Prefix with the REASON
    If sReason <> "" Then
        oMsg.Subject = sReason & oMsg.Subject
    End If

    '-- Copy the body
    oMsg.BodyFormat = oMail.BodyFormat
    Select Case oMsg.BodyFormat
        Case OlBodyFormat.olFormatHTML
            oMsg.HTMLBody = sReason & oMail.SenderName & vbCrLf & _
                        "<hr>" & vbCrLf & _
                        oMail.HTMLBody
        Case Else
            oMsg.Body = sReason & oMail.SenderName & vbCrLf & _
                    "-----------------------------------------------" & vbCrLf & _
                    oMail.Body
    End Select
    
'    '-- Copy Attachments
'    For Each oAttach In oMail.Attachments
'        oMsg.Attachments.Add oAttach., oAttach.Type, oAttach.Position, oAttach.DisplayName
'    Next
    
    '-- Set module level variables if they are blank
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    sTempDir = oFSO.GetSpecialFolder(2)                        ' TemporaryFolder
    
    '-- Copy Attachments
    For Each oAttach In oMail.Attachments
        sPath = sTempDir & oAttach.fileName
        oAttach.SaveAsFile sPath
        oMsg.Attachments.Add sPath, , , oAttach.DisplayName
        oFSO.DeleteFile sPath
        Set oAttach = Nothing
    Next
    
    '-- Send the mail
    oMsg.sEnd
    
    SendMail = True

    Set oMsg = Nothing
    Set oAttach = Nothing

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP09_DistributeMail:SendMail [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function
