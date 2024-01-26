Attribute VB_Name = "HELP01_MakeTicket_from_Rule_5t"
'############################################################################################################
'## HELP01_MakeTicket_from_Rule - VBA Code launched from Outlook Rule                                      5t
'############################################################################################################

Option Explicit

'-- *** HELP00_Public = Module must be loaded to use Public Constants and Variables ***

'--------------------------------------------------------------------------------------------------
' These VBA routines are designed for the HELP email account when it receives new emails.  Each
' email is tested against a list of "Filing Rules", to find the proper Client and Ticket Number.
' If the Ticket exists, then IPM.Note.TB_Mail form is filled out.  If not, then both the
' IPM.Task.TB_Ticket and IPM.Note.TB_Mail forms are created.
'
' HELP_MakeTicket           - Triggered from an Outlook Rule that uses a "Run a Script" as an action.
' HELP_ProcessEmail         - Processes mail that comes in
' HELP_ProcessTime          - Processes time that gets submitted
' HELP_MakeTicket_Init      - Initializes all the module level variables to do once per Outlook session.
' ProcessAllMail            - Loops through all the mail in the inbox to call the HELP_MakeTicket routine.
' NextTicketNum             - Generates the next ticket number for a client ticket
'
' The "Filing Rules" are managed in a Outlook folder or Lists.
'   Client  - Client abbreviation (4 to 8 characters)
'   Subject - Value to look for in the Sender Address and Subject
'   Company - Company to assign to the Task (used only for description)
'
' The Clients and Ticket Numbers are managed in a separate Outlook Task folder.  Each taskitem is a
' Client record with the following fields populated...
'   Client (Subject)                    - Client abbreviation (4 to 8 characters)
'   Ticket Number (Billing Information) - 4 Digit number used for the Ticket Number
'   Help Email (Contact)                - The HELP email to send the reply so that it gets recorded
'                                         back into the Ticket system
'--------------------------------------------------------------------------------------------------
 
'--------------------------------------------------------------------------------------------------
'-- Module Level Constants

'-- Errors for rejecting a time entry
Const ERR_BAD_TICKET_HEADER = "No Ticket Header found." & vbCrLf & "Correct the Subject and add |<Client>|<Ticket#>|"
Const ERR_LONG_DURATION = "The duration exceeds 24 hours." & vbCrLf & "Correct Start & End dates and times"
Const ERR_NO_TIME_DESC = "There is no detailed description for the time entry." & vbCrLf & "Enter the work that was performed."
Const ERR_NO_TICKET = "There is no Ticket matching the ticket header."
Const ERR_NO_TIME_TOPIC = "Topic field must be filled in."

'--------------------------------------------------------------------------------------------------
'-- Module Level Variables because everytime the rule gets triggered, these do not have to be reinitialized.
Dim m_blnMakeTicket_Init    As Boolean  '-- Flag to only run the Init routine once

'-- Array to hold the Filing Rules. Array goes from 0 to N
Dim m_asFilingRules()       As String
Dim m_iRowsRules            As Integer

'--------------------------------------------------------------------------------------------------
' HELP_MakeTicket - Trigged via an Outlook rule.  It checks each new mail and decides what to
'                   process.
'
'   NOTE: oMail argument must be included for Outlook Rule to run this as a script.
'
'  Procedure for making a Rule run a macro, to get around Outlook 2010 bug.
'   1) must be Public or Sub (not private)
'   2) must have (xyz As MailItem) as a argument where xyz is the object that triggers the rule
'   3) create a new rule using "run a script" to execute the script
'   4) after the Rule is created, change to (xyz As Object) to handle none mailitem objects
'--------------------------------------------------------------------------------------------------
'Public Sub HELP_MakeTicket(oItem As MailItem)
Public Sub HELP_MakeTicket(oItem As Object)
    Dim oMail       As MailItem
    Dim oMtgReq     As MeetingItem
    Dim blnRtnList  As Boolean

    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If

    '-- Initialize all the module level variables only the first time
    If Not m_blnMakeTicket_Init Then
        m_blnMakeTicket_Init = HELP_MakeTicket_Init
    End If

'    -- Update the categories for over due tickets.
'    Revive_Overdue_Tickets '-- MOVED TO separate EXE nightly batch run

    '-- Email entries - Make sure the email is a type we can process
    If (InStr(1, oItem.MessageClass, MSGCLS_Note, vbTextCompare) > 0 Or _
        StrComp(oItem.MessageClass, MSGCLS_Mail, vbTextCompare) = 0 Or _
        StrComp(oItem.MessageClass, MSGCLS_Reply, vbTextCompare) = 0) Then
        
        Set oMail = oItem
        
        '-- If the trigger email starts with a ?, return the list of tickets
        blnRtnList = False
        If Left(oMail.Subject, 1) = "?" Then
            blnRtnList = HELP_ReturnTicketList(oMail)
        End If
        
        '-- If a list of tickets were not sent (e.g. Alert email), then process the email
        If Not blnRtnList Then
            '-- Process HELP Emails
            HELP_ProcessEmail oMail
        End If
    
    '-- Time entries
    ElseIf (StrComp(oItem.MessageClass, MSGCLS_MtgRequest, vbTextCompare) = 0) Then
        Set oMtgReq = oItem
        '-- Accept Time emails
        HELP_ProcessTime oMtgReq
    End If
    
'    '-- Do the HeartBeat processing
'    HeartBeat m_fldrTickets    '-- MOVED TO separate EXE nightly batch run
    
    Set oMail = Nothing
    Set oMtgReq = Nothing
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:HELP_MakeTicket [" & Err.Number & "] " & Err.Description & " >" & oItem.Subject
    If ERR_RESUME Then Resume Next
End Sub


'--------------------------------------------------------------------------------------------------
' HELP_ProcessTime - Process all Time entries (meeting appointment requests)
'--------------------------------------------------------------------------------------------------
Private Sub HELP_ProcessTime(ByRef oMtgReq As MeetingItem)
    Dim sClient         As String   '-- Used to store values for ticket creation
    Dim sTicketNum      As String
    Dim sTopic          As String
    Dim sInitials       As String
    Dim sSubject        As String
    Dim sSender         As String
    Dim sError          As String
    
    Dim oTicket           As TaskItem
    Dim oAppt           As AppointmentItem
    Dim oTime           As AppointmentItem
    Dim oReject         As MeetingItem

    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Clear off the key variables that are used to create the Ticket
    sClient = ""
    sTicketNum = ""
    sTopic = ""
    sInitials = ""
    
    '-- Get the Sender's name
    sSender = oMtgReq.SenderName
   
    '-- Clean up the mail Subject
    sSubject = Clean_Up_Subject(oMtgReq.Subject)
    
    '-- Proper time entries must have a ticket header
    Parse_Ticket_Header sSubject, sClient, sTicketNum, sTopic
    
    '-- For valid time entries, find the ticket
    sError = ""
    Set oTicket = Nothing
    If (sClient <> "" And sTicketNum <> "") And _
       Not (StrComp(sClient, NOCLIENT_CLIENT, vbTextCompare) = 0 And sTicketNum = NOCLIENT_TICKETNUM) Then
    
        '-- Find the ticket
        Set oTicket = CmFindMatchingTask(m_fldrTickets.Items, sClient, sTicketNum, False)
    Else
        sError = ERR_BAD_TICKET_HEADER
    End If
        
    '-- If a ticket is found, then accept the time.
    If Not oTicket Is Nothing Then
    
        '-- Accept the appointment, and add it to the calendar (may not be needed)
        Set oAppt = oMtgReq.GetAssociatedAppointment(True)
        oAppt.ReminderOverrideDefault = True
        oAppt.ReminderSet = False

        '-- Error, if the duration is over 24hrs
        If oAppt.Duration / 60 > 24 Then
            sError = ERR_LONG_DURATION
            Set oTicket = Nothing
            
        '-- Error, if no Time Entry description
        ElseIf Trim(oAppt.Body) = "" Or Len(Trim(oAppt.Body)) < 3 Then
            sError = ERR_NO_TIME_DESC
            Set oTicket = Nothing
        
        ElseIf sTopic = "" Then
            sError = ERR_NO_TIME_TOPIC
            Set oTicket = Nothing
        '-- No errors
        Else

            '-- Make a time entry and put it in the Time folder
            Set oTime = Help_MakeTicket_Time(oAppt, sClient, sTicketNum, sTopic, sSubject, sSender, MSGCLS_Time, m_fldrTime)
        End If
    End If
        
    '-- if no Ticket is found (needs to be a separate test to send errors in reject email
    If oTicket Is Nothing Then

        If sError = "" Then
            sError = ERR_NO_TICKET
        End If
    
        '-- Get the appointment, but don't put it on the calendar
        Set oAppt = oMtgReq.GetAssociatedAppointment(False)
        oAppt.ReminderOverrideDefault = True
        oAppt.ReminderSet = False
        
        If Not oAppt Is Nothing Then
            Set oReject = oAppt.Respond(olMeetingDeclined, True)
        
            '-- Send a rejection response
            oReject.Body = sError
            oReject.sEnd
            
            '-- Check for bad time submissions
            If sError = ERR_LONG_DURATION Then
                'm_oSpeech.Speak sSender & TALK_TIME_REJECT1
            ElseIf sError = ERR_NO_TIME_DESC Then
                'm_oSpeech.Speak sSender & TALK_TIME_REJECT2
            ElseIf sError = ERR_BAD_TICKET_HEADER Then
                'm_oSpeech.Speak sSender & TALK_TIME_REJECT3
            End If

        Else
            'm_oSpeech.Speak sSender & TALK_TIME_RESUBMIT
        End If

    End If

    '-- Delete the Time email (meeting request)
    On Error Resume Next
    oMtgReq.Delete
    If Not ERR_IGNORE Then On Error GoTo 0          '--$$$ Ignore Runtime Errors
    If ERR_HANDLER Then On Error GoTo ERRORHANDLER  '--$$$ Use ErrorHandler

    Set oTicket = Nothing
    Set oAppt = Nothing
    Set oTime = Nothing
    Set oReject = Nothing
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:HELP_ProcessTime [" & Err.Number & "] " & Err.Description & " >" & oMtgReq.Subject
End Sub


'--------------------------------------------------------------------------------------------------
' HELP_ProcessEmail - Checks each email against a list of filing rules or existing ticket
'--------------------------------------------------------------------------------------------------
Private Sub HELP_ProcessEmail(ByRef oMail As MailItem)
    
    Dim sClient         As String   '-- Used to store values for ticket creation
    Dim sTicketNum      As String
    Dim sTopic          As String
    Dim sInitials       As String
    Dim sSubject        As String
    Dim sSender         As String
    Dim sAssignee       As String
    Dim sDueDate        As String      '-- No Longer Used
    Dim iImportance     As Integer
        
    Dim blnFound        As Boolean  '-- Used for Filing rules
    Dim blnGoodCmd      As Boolean
    Dim sTestClient     As String
    Dim sTestStr        As String
    Dim iLenTestClient  As Integer
    Dim iLenTestStr     As Integer
    
    Dim idx             As Integer  '-- Used in loops
    
    Dim blnNewTicket    As Boolean  '-- Used to create new Tickets
    
    Dim oTicket         As TaskItem '-- Ticket
    Dim oEmail          As MailItem '-- for emails missing tickets
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Initialize the key variables that are used to create the Ticket
    sClient = ""
    sTicketNum = ""
    sTopic = ""
    sInitials = ""
    sAssignee = ""
    sDueDate = ""   '-- No longer used
    sSender = oMail.SenderName      '-- Get the Sender's name
    sSubject = Clean_Up_Subject(oMail.Subject)      '-- Clean up the mail Subject

    '----------------------------------------------------------------------------------------------
    '-- Test 1: If Critical Alert, then Play Sound
    '----------------------------------------------------------------------------------------------
    '-- Reset importance to Normal even though the incoming email is set to high
    oMail.Importance = olImportanceNormal
    
    If InStr(1, oMail.Subject, ALERT_Tag, vbTextCompare) > 0 Then
        oMail.Importance = olImportanceHigh

    ElseIf InStr(1, sSender, TALK_VOICEMAIL, vbTextCompare) > 0 Then
        oMail.Importance = olImportanceHigh
        sClient = TALK_GOOG
    
    End If
    
    '----------------------------------------------------------------------------------------------
    '-- Test 2: Check to see if the message has a ticket header |abbr|num|topic
    '--                                                        0| 1  | 2 |  3
    '----------------------------------------------------------------------------------------------
    Parse_Ticket_Header sSubject, sClient, sTicketNum, sTopic
    
    '----------------------------------------------------------------------------------------------
    '-- Test 3: If no header was found, match the mail against the filing rules to get the Client
    '----------------------------------------------------------------------------------------------
    If sClient = "" Then
    
        sTicketNum = ""     '-- Clear out the Ticket Number

        '-- These flags are used to control the testing for different conditions
        blnFound = False        '-- set to TRUE upon a match in the table
        blnGoodCmd = False

        '-- Loop through all the Filing Rules table until match Found
        '-- Filing Rules table is initialize during the INIT subroutine
        idx = 0
        While m_iRowsRules > -1 And Not blnFound And idx <= m_iRowsRules

            sTestClient = m_asFilingRules(idx, FILERULE_CLIENT)     '-- Client Abbrev
            sTestStr = m_asFilingRules(idx, FILERULE_MATCH_TEXT)    '-- Value to Match
            iLenTestClient = Len(sTestClient)
            iLenTestStr = Len(sTestStr)

            '-- Does the Sender's Email Address match the Search string?
            If InStr(1, oMail.SenderEmailAddress, sTestStr, vbTextCompare) > 0 Then
                blnFound = True

            '-- Does the subject start with a Command string? (this handles +CCCC)
            ElseIf StrComp(sTestStr, Left(sSubject, iLenTestStr), vbTextCompare) = 0 Then

                blnGoodCmd = Parse_Ticket_Command(sSubject, sClient, sTopic, sAssignee, sDueDate, iImportance)

                If blnGoodCmd Then
                    sSubject = sTopic
                    oMail.Subject = sSubject
                    oMail.Importance = iImportance
                    blnFound = True
                End If

            '-- Was this sent from TechBldrs?
            ElseIf InStr(1, oMail.SenderEmailAddress, TB1, vbTextCompare) > 0 _
                Or InStr(1, oMail.SenderEmailAddress, TB2, vbTextCompare) > 0 Then

                '-- Who was it sent to?
                If InStr(1, oMail.Recipients.Item(1).Address, sTestStr, vbTextCompare) > 0 Then
                    blnFound = True
                End If
            End If

            '-- If found, then the client is the one that matched the rule or was set in Parse_Ticket_Command
            If blnFound Then
                If sClient = "" Then
                    sClient = sTestClient
                End If

            '-- If not found, then goto the next rule
            Else
                idx = idx + 1
            End If

            '-- Allow the System to do other things
            DoEvents
        Wend
    End If
    
    '--------------------------------------------------------------------------------
    '-- Client, but no Ticket Number -> Get next number & New Ticket
    '--------------------------------------------------------------------------------
    '-- Goes to the Crap Catcher
    If sClient = "" Then
        sClient = NOCLIENT_CLIENT
        sTicketNum = NOCLIENT_TICKETNUM
        sTopic = sSubject
    End If

    Set oTicket = Nothing
    
    '-- Get a new ticket number
    If sClient <> "" And sTicketNum <> "" Then
        Set oTicket = CmFindMatchingTask(m_fldrTickets.Items, sClient, sTicketNum)
    End If

    '--------------------------------------------------------------------------------
    '-- Find an existing ticket or process a new ticket
    '--------------------------------------------------------------------------------
    '-- If no ticket was found, then create a new ticket
    If oTicket Is Nothing Then
        
        '-- Get a new ticket number; also prevents replying to old tickets generating dup tickets
        sTicketNum = NextTicketNum(sClient)
       
        '-- Default the Topic with the subject, add ticket header to the subject
        sTopic = sSubject
        sSubject = TKTDELIM & sClient & TKTDELIM & sTicketNum & TKTDELIM & " {" & sSender & "} " & sSubject '-- Added Sender to the Subject
    
        '-- Create a new ticket and move it to the Tickets folder
        Help_MakeTicket_Task oMail, sClient, sTicketNum, sTopic, sSubject, sSender, sAssignee, MSGCLS_Ticket, m_fldrTickets
                                             
        '-- Send RECEIPT EMAIL for new ticket '-- NOT YET IMPLEMENTED
   
    Else
        '-- Set all the flags and status of this ticket
        '-- Call a sub in another module so that we don't have to move the Const and variables to this module
        Update_Ticket oTicket, oMail
    End If
     
    '-- Change the incoming mail to a Ticket Mail and move it to the Mail folder
    Set oEmail = Help_MakeTicket_Mail(oMail, sClient, sTicketNum, sTopic, sSubject, MSGCLS_Mail, m_fldrMail)
   
    '-- Distribute the mail to each connected mailbox; New ticket once moved will be nothing
    If Not oTicket Is Nothing Then
        HELP_DistributeMail oTicket.UserProperties(TKT_ASSIGNEE), oEmail
    End If

    Set oTicket = Nothing
    Set oEmail = Nothing
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:HELP_ProcessEmail [" & Err.Number & "] " & Err.Description & " >" & oMail.Subject
    If ERR_RESUME Then Resume Next
End Sub


'----------------------------------------------------------------------------------------
'-- HELP_MakeTicket_Init - Sets up the module level variables, folders and filing rules
'----------------------------------------------------------------------------------------
Private Function HELP_MakeTicket_Init() As Boolean
   
    HELP_MakeTicket_Init = False

    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Create Speech Object
'    Set m_oSpeech = CreateObject("sapi.spvoice")
    
    '-- Initialize all the public global vaiables
    If CmInitPublicGlobals Then
    
        '------------------------------------------------------------------------
        '-- Load the Filing Rules into an array
        '------------------------------------------------------------------------
        '-- Col 0: Client Abbrev, Col 1: Value to match, Col 2 : Company Name
        m_iRowsRules = CmTableLoadfromBody(m_fldrLists.Items(LIST_FILING_RULES), m_asFilingRules)
        
        '-- Set the flag so that we don't do this again
        HELP_MakeTicket_Init = True
        
    End If
    

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:HELP_MakeTicket_Init [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' Clean_Up_Subject - Remove all the unneeded characters from a subject string
'--------------------------------------------------------------------------------------------------
Private Function Clean_Up_Subject(ByRef sTemp As String) As String
    Dim sTemp2 As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    sTemp2 = Trim(sTemp)
    sTemp2 = Replace(sTemp2, "FW: ", "")
    sTemp2 = Replace(sTemp2, "FW:", "")
    sTemp2 = Replace(sTemp2, "fw: ", "")
    sTemp2 = Replace(sTemp2, "fw:", "")
    sTemp2 = Replace(sTemp2, "Fw: ", "")
    sTemp2 = Replace(sTemp2, "Fw:", "")

    sTemp2 = Replace(sTemp2, "RE: ", "")
    sTemp2 = Replace(sTemp2, "RE:", "")
    sTemp2 = Replace(sTemp2, "re: ", "")
    sTemp2 = Replace(sTemp2, "re:", "")
    sTemp2 = Replace(sTemp2, "Re: ", "")
    sTemp2 = Replace(sTemp2, "Re:", "")

    sTemp2 = Replace(sTemp2, "Updated: ", "")
    sTemp2 = Replace(sTemp2, "Updated:", "")
    sTemp2 = Replace(sTemp2, "Copy: ", "")
    sTemp2 = Replace(sTemp2, "Copy:", "")
    sTemp2 = Replace(sTemp2, "  ", " ")
    sTemp2 = Replace(sTemp2, "  ", " ")
    sTemp2 = Replace(sTemp2, "  ", " ")
    sTemp2 = Replace(sTemp2, "  ", " ")
    Clean_Up_Subject = Trim(sTemp2)

    Exit Function
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:Clean_Up_Subject [" & Err.Number & "] " & Err.Description & " >" & sTemp
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' Parse_Ticket_Header - A properly formed ticket header is separated using the Ticket Delimiter
'                        |abbr|num|topic
'                       0| 1  | 2 |  3
'--------------------------------------------------------------------------------------------------
Private Function Parse_Ticket_Header(ByRef sSubject As String, _
                                ByRef sClient As String, _
                                ByRef sTicketNum As String, _
                                ByRef sTopic As String) As Boolean

    Dim arrStr()    As String   '-- Used to split the Subject using ticket delimiter, into Client, Ticket Num & Topic
    Dim idx         As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Parse_Ticket_Header = False
    
    '-- Clear the substrings before any tests
    sClient = ""
    sTicketNum = ""
    sTopic = ""

    '--0 If the first character is not the Ticket Delimiter, then exit out
    If (Left(sSubject, 1) <> TKTDELIM) Then Exit Function
    
    '--1 Split the subject into sections
    arrStr = Split(sSubject, TKTDELIM)
    
    '--2 If the ticket number is numeric, then proper header
    If IsNumeric(Trim(arrStr(2))) Then
        sClient = UCase(Trim(arrStr(1)))
        sTicketNum = Trim(arrStr(2))
        sTopic = Trim(Replace(arrStr(3), vbTab, ""))
        
        '-- Handle a subject that included the TKTDELIM in the text
        For idx = 4 To UBound(arrStr)
            sTopic = sTopic & TKTDELIM & Trim(Replace(arrStr(idx), vbTab, ""))
        Next
    End If
    
    Parse_Ticket_Header = True
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:Parse_Ticket_Header [" & Err.Number & "] " & Err.Description & " >" & sSubject
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' Parse_Ticket_Command - A properly formed ticket command returns TRUE
'--     syntax: +CCCC[{!.}ti.yymmdd]<space>Topic
'--             - "+CCCC" = create new ticket for client CCCC
'--             - <space> or "!" or "." ends client code
'--                         "!" indicates urgent status
'--             - ti is the Tech Initials to assign to the ticket (2 char)
'--             - YYMMDD as due date
'--------------------------------------------------------------------------------------------------
Private Function Parse_Ticket_Command(ByRef sSubject As String, _
                                      ByRef sClient As String, _
                                      ByRef sTopic As String, _
                                      ByRef sAssignee As String, _
                                      ByRef sDueDate As String, _
                                      ByRef iImportance As Integer) As Boolean
    Dim idx As Integer
    Dim sCommand As String
    Dim arrStr()    As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Parse_Ticket_Command = False
    
    sTopic = ""
    sAssignee = ""
    sDueDate = ""
    iImportance = olImportanceNormal
    
    '--1 Find the space the separate the Topic from the command
    sCommand = ""
    idx = InStr(1, sSubject, " ", vbTextCompare)
    If idx > 0 Then
        sTopic = Trim(Mid(sSubject, idx + 1, Len(sSubject) - idx))
        sCommand = Mid(sSubject, 1, idx - 1)
    End If
    If sCommand = "" Then sCommand = sSubject
    
    '--2 Search the command for Urgent "!"
    idx = InStr(1, sCommand, "!", vbTextCompare)
    If idx > 0 Then
        iImportance = olImportanceHigh
        sCommand = Replace(sCommand, "!", ".")
    End If
    
    '--3 Break up the command. If this is a properly formed command, then
    '--     array(0) = +CCCC
    '--     array(1) = tech Initials
    '--     array(2) = due date
   
    arrStr = Split(sCommand, ".")
    
    If UBound(arrStr) > 2 Then
        '--Bad Command
        Exit Function
    End If
    
    If UBound(arrStr) = 2 Then
        sDueDate = arrStr(2)
    End If
   
    If UBound(arrStr) >= 1 Then
    
        '-- KLUDGE to work with Kasya's groups.
        If Len(arrStr(1)) = 2 Then sAssignee = arrStr(1)
        
        '-- KLUDGE - If the 1st char is numeric, then it came from Kaseya Group name
        If IsNumeric(Left(sAssignee, 1)) Then sAssignee = ""
    End If
    
    '-- This will get replaced with the client from the filing rule
    sClient = UCase(Mid(arrStr(0), 2, Len(arrStr(0)) - 1))
    
    Parse_Ticket_Command = True
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:Parse_Ticket_Command [" & Err.Number & "] " & Err.Description & " >" & sSubject
    If ERR_RESUME Then Resume Next
End Function


'--------------------------------------------------------------------------------------------------
' ProcessAllMail - Used to manually process all the emails of a folder.
'
'   NOTE: This is not part of the Rule's script.
'--------------------------------------------------------------------------------------------------
Public Sub Help_ProcessAllMail()
    Dim fldr    As Folder 'MAPIFolder
    Dim cItems  As Items
    Dim oItem   As Object 'MailItem
    Dim idx     As Integer
    Dim iCnt    As Integer
           
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
           
    '-- Get the folder of emails to process
    m_blnMakeTicket_Init = False
    Set fldr = Application.GetNamespace("MAPI").PickFolder
    If fldr Is Nothing Then Exit Sub
    
    '-- Sort the emails by descending order of the time recieved
    Set cItems = fldr.Items
    cItems.Sort "[ReceivedTime]", True
    
    '-- Call the HELP_MakeTicket function with each email.  This simulates emails arriving
    '-- in the folder.
    '-- Loop in reverse order, because emails that match the filter are deleted.
    iCnt = cItems.Count
    For idx = iCnt To 1 Step -1
        Set oItem = cItems.Item(idx)
'        If oItem.Class = olMail Then
        HELP_MakeTicket oItem
'        End If
        DoEvents
    Next
 
    Set fldr = Nothing
    Set cItems = Nothing
    Set oItem = Nothing
    
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:Help_ProcessAllMail [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Sub

'----------------------------------------------------------------------------------------
' NextTicketNum - Looks for an item that matches the Client, increments the number
'                 in the BillingInformation field
'
'   NOTE: This is designed for the Ticket Numbers folder, and it must remain as an
'         external record because it Ticket Number can be generated from multiple
'         routines and forms.
'----------------------------------------------------------------------------------------
Private Function NextTicketNum(sClient As String) As String
    Dim oTicketNum      As TaskItem
    Dim sLastTicketNum  As String
    Dim sNextTicketNum  As String
 
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
 
    NextTicketNum = ""
 
    'Find the Ticket number matching the Subject field = to sClient
    Set oTicketNum = Nothing
    On Error Resume Next
    Set oTicketNum = m_fldrTicketNum.Items(sClient)
    If Not ERR_IGNORE Then On Error GoTo 0          '--$$$ Ignore Runtime Errors
    If ERR_HANDLER Then On Error GoTo ERRORHANDLER  '--$$$ Use ErrorHandler
    
    If Not (oTicketNum Is Nothing) Then
    
        '-- Get the number
        sLastTicketNum = oTicketNum.BillingInformation
        
        '-- Exclude first 11 ticket numbers
        If (sLastTicketNum Mod 10000 < 11) Then
            sLastTicketNum = 10
        End If
        
        '-- Add 1 and pad with leading "0" to 4 chars
        sNextTicketNum = Right("000" & sLastTicketNum + 1, 4)
        
        '-- Save the new number
        oTicketNum.BillingInformation = sNextTicketNum
        oTicketNum.Save
 
        '-- Return the new number
        NextTicketNum = sNextTicketNum
    End If
    
    Set oTicketNum = Nothing
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP01_MakeTicket_from_Rule:NextTicketNum [" & Err.Number & "] " & Err.Description & " >" & sClient
    If ERR_RESUME Then Resume Next
End Function

