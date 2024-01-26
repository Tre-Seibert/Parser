Attribute VB_Name = "HELP00_PUBLIC_1U"
'############################################################################################################
'## HELP00_Public - VBA module that holds Public Constants and Variables common to many Ticket Builder     1t
'############################################################################################################

Option Explicit

'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                                                            ByVal uFlags As Long) As Long

Public Const DOMAIN = "techbldrs.com"   '-- Used to support multiple Public Folders

Public Const ERR_IGNORE = False         '-- Turns on or off the ERR_IGNORE Run-time errors
Public Const ERR_HANDLER = True          '-- Turns on using the ErrorHandler Jump point
Public Const ERR_RESUME = True          '-- Turns on whether to continue to the line after the error point, or stop

'--------------------------------------------------------------------------------------------------
'-- Folders that hold the data for the Ticket System
Public Const PF = "Public Folders"
Public Const APF = "All Public Folders"
Public Const FAV = "Favorites"
Public Const FLDR_COMPANY = "TechBldrs Inc"                     '-- Top-level Parent folder
Public Const FLDR_TICKET = "TB Tickets"                         '-- +- Ticket Folder (Tasks)
Public Const FLDR_BACKUPTICKET = "TB Backup Tickets"            '-- +- Ticket Backup Folder (Tasks) - used to recover system randomly deleted tickets
Public Const FLDR_TICKET_MAIL = "TB Mail"                       '--     +- Email Folder (Mail) for Tickets
Public Const FLDR_TICKET_TIME = "TB Time"                       '--     +- Time Folder (Appointments) for Tickets
Public Const FLDR_TICKET_PART = "TB Part"                       '--     +- Parts Folder (Tasks) for Tickets
Public Const FLDR_DATA = "_#DATA#"                              '-- +- Folder with the the following subfolders:
Public Const FLDR_LISTS = "Lists"                               '--     +- Folder of lists for the Ticket Form and Help Filing
Public Const FLDR_TEMPLATES = "Templates"                       '--     +- Folder of email templates
'Public Const FLDR_DATA_FILING_RULES = "Filing Rules"            '--     +- Folder with Filing Rules
Public Const FLDR_DATA_TICKETNUM = "Clients & Ticket Numbers"   '--     +- Folder with Last Ticket Numbers
'Public Const FLDR_DATA_TECH_INITALS = "Technician Initials"     '--     +- Folder with Technician Initals
Public Const FLDR_CALENDAR = "TB Calendar"                      '-- +- Group Dispatch Calendar
Public Const FLDR_CLIENTS = "TB Clients"                        '-- +- Group Client * Company Folder (Contacts)
Public Const FLDR_DEL_TICKETS = "TB DEL Tickets"   '-- Subfolder of TB Tickets where DELETED tickets are moved to
Public Const FLDR_BACKUP_TICKETS = "TB Backup Tickets"   '-- Subfolder of TB Tickets where DELETED tickets are moved to
Public Const FLDR_TICKET_ARCHIVE = "TB Tickets Archive"     '-- +- Ticket Folder (Archive)(Tasks)


'-- Tables needed for Help Ticket Processing
Public Const LIST_FILING_RULES = "Help.FilingRules"             '-- Item in the Company\Data\Lists folder
Public Const FILERULE_CLIENT = 0                                '--  Column of the Filing Rule Table
Public Const FILERULE_MATCH_TEXT = 1                            '--  Column of the Filing Rule Table
Public Const FILERULE_COMPANY_NAME = 2                          '--  Column of the Filing Rule Table
Public Const LIST_ASSIGNEES = ".Assignee"                       '-- Item in the Company\Data\Lists folder
Public Const ASSIGNEE_INITIALS = 0
Public Const ASSIGNEE_NAME = 1
Public Const ASSIGNEE_EMAIL = 2
Public Const ASSIGNEE_TEXTMSG = 3
Public Const ASSIGNEE_CREWHU = 4

'-- TechBldrs possible domains
Public Const TB1 = "@tecmanage.com"
Public Const TB2 = "@techbldrs.com"
Public Const TB3 = "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/"
Public Const ADMIN = "administrator"

'-- Alerts to check for in the Subject of the Mail
Public Const ALERT_Tag = "?c?"                          '-- Alerts from Kaseya
Public Const ALERT_Backup = "Backup Failed!"
Public Const ALERT_Backup2 = "Check Backup"

'-- Google Voice Voicemails
Public Const TALK_VOICEMAIL = "Google Voice"
Public Const TALK_GOOG = "GOOG"

'-- No Client Defaults
Public Const NOCLIENT_CLIENT = "<none>"
Public Const NOCLIENT_TICKETNUM = "0000"
Public Const NOCLIENT_TOPIC = "*** Need to File ****"
 
'-- Forms used in Ticket Builder
Public Const MSGCLS_Note = "IPM.Note"
Public Const MSGCLS_MtgRequest = "IPM.Schedule.Meeting.Request"
Public Const MSGCLS_Ticket = "IPM.Task.TB_Ticket"
Public Const MSGCLS_Mail = "IPM.Note.TB_Mail"
Public Const MSGCLS_Time = "IPM.Appointment.TB_Time"
Public Const MSGCLS_Part = "IPM.Task.TB_Part"
Public Const MSGCLS_Reply = "IPM.Note.TB_Reply"

'-- Common
Public Const TKTDELIM = "|"                                     '-- Delimiter used in the Ticket Header
Public Const Tkt_DELETED_Substr = "-2DEL-"         '-- Set in TICKET_MergeSelectedTickets when tickets are merged
Public Const RESPONSE_DAYS = 2                                  '-- Days to respond to a Ticket
Public Const NODATE = #1/1/4501#                                '-- This is a null date

'-- These are User Defined fields and for consistency are also the name of the form object.
'-- USED IN 2 PLACES: (A) Help System VBA code to create the Ticket task item (B) TB_Ticket form VBS
Public Const TKT_ACTION = ".Action"
Public Const TKT_ASSIGNEE = ".Assignee"
Public Const TKT_CAUSE1 = ".Cause1"
Public Const TKT_CLIENT = ".Client"
Public Const TKT_DATE_CONTRACT_EXPIRE = ".DateContractExpire"
Public Const TKT_DATE_CREATED = ".DateCreated"
Public Const TKT_DATE_LAST_ACTIVITY = ".DateLastActivity"
Public Const TKT_HRS_ACTUAL_TOTAL = ".HrsActualTotal"
Public Const TKT_HRS_BILLABLE_TOTAL = ".HrsBillableTotal"
Public Const TKT_HRS_COMPLETION = ".HrsCompletion"
Public Const TKT_HRS_DURATION = ".HrsDuration"
Public Const TKT_HRS_GRATIS_TOTAL = ".HrsGratisTotal"
Public Const TKT_INVOICE_NOTES = ".InvoiceNotes"
Public Const TKT_INVOICE_NUM = ".InvoiceNum"
Public Const TKT_JOB = ".Job"
Public Const TKT_LOG = ".Log"
Public Const TKT_MACHINE_NAME = ".MachineName"
Public Const TKT_MACHINE_SUPPORT = ".MachineSupport"
Public Const TKT_MAIL_TEMPLATE = ".MailTemplate"
Public Const TKT_REASON = ".Reason"
Public Const TKT_REQUESTOR = ".Requestor"
Public Const TKT_STATUS = ".Status"
Public Const TKT_TECHNAME = ".Tech"
Public Const TKT_TICKETMONTH = ".TicketMonth"
Public Const TKT_TICKETNUM = ".TicketNum"
Public Const TKT_TICKETYEAR = ".TicketYear"
Public Const TKT_TOPIC = ".Topic"
Public Const TKT_USER = ".User"

Public Const TKT_STATUS_NEW = "New"
Public Const TKT_STATUS_TO_CLIENT = "Email Sent"
Public Const TKT_STATUS_FROM_CLIENT = "Client Replied"

Public Const TKT_CAT0_URGENT = "0 Urgent"
Public Const TKT_CAT1_HIGH = "1 High"
Public Const TKT_CAT1_REOPENED = "1 Re-Opened"
Public Const TKT_CAT2_NORMAL = "2 Normal"
Public Const TKT_CAT3_FOLLOWUP = "3 Follow Up"
Public Const TKT_CAT4_BACKUP = "4 Backup"
Public Const TKT_CAT5_ONSITE = "5 On-Site"
Public Const TKT_CAT6_PROJECT = "6 Project"
Public Const TKT_CAT7_ORDERED = "7 Ordered"
Public Const TKT_CAT8_TIME = "8 Time"
Public Const TKT_CAT9_REVIEW = "9 REVIEW"

Public Const TKT_REASON_SUPPORT = "Support"
Public Const TKT_REASON_BILLABLE = "Billable"
Public Const TKT_REASON_RESOLVED = "Resolved"
Public Const TKT_REASON_INTERNALPROJECT = "InternalProject"
Public Const TKT_REASON_ADMIN = "Admin"
Public Const TKT_REASON_ALERT = "Alert"

'-- These are User Defined fields and for consistency are also the name of the form object.
'-- USED IN 2 PLACES: (A) Help System VBA code to create Mail mail item, (B) TB_MAIL form VBS
Public Const MAIL_APPROVAL = ".Approval"
Public Const MAIL_CLIENT = ".Client"
Public Const MAIL_DATE_CREATED = ".DateCreated"
Public Const MAIL_TICKETNUM = ".TicketNum"
Public Const MAIL_TOPIC = ".Topic"

Public Const PART_TOPIC = ".Topic"

'-- These are User Defined fields and for consistency are also the name of the form object.
'-- COPY TO (A) HELP00_PUBLIC (B) TKT_FORM_TB_Ticket (C) TKT_FORM_TB_Time
Public Const TIME_BILLABLE = ".Billable"           '-- Interactive-only Flag
Public Const TIME_GRATIS = ".Gratis"               '-- Interactive-only Flag
Public Const TIME_QUOTED = ".Quoted"               '-- Interactive-only Flag
Public Const TIME_REVIEWED = ".REVIEWED"           '-- NOT USED: Interactive-only Flag
Public Const TIME_INVOICE_NUM = ".InvoiceNum"      '-- Filled in during Invoice creation via Outlook macro
Public Const TIME_JOB = ".Job"                     '-- Filled in during Invoice creation via Outlook macro
Public Const TIME_INVOICE_DESC = ".InvoiceDesc"    '-- Used during Invoice creation via Outlook macro
Public Const TIME_BILLEND = ".BillEnd"
Public Const TIME_BILLHOURS = ".BillHours"
Public Const TIME_BILLSTART = ".BillStart"
Public Const TIME_CLIENT = ".Client"
Public Const TIME_DATE_CREATED = ".DateCreated"
Public Const TIME_HOURS = ".Hours"
Public Const TIME_TECH = ".Tech"
Public Const TIME_TICKETNUM = ".TicketNum"
Public Const TIME_TOPIC = ".Topic"
Public Const TIME_WORKDATE = ".WorkDate"

'-- Heartbeat
Public m_ihour             As Integer  '-- Saves the current hour
Public m_iday              As Integer  '-- Saves the current day
Public m_blnInitialized    As Boolean  '-- Used in HeartBeat_SendTextMsg to load the m_asPhoneTextMsg array
Public m_asPhoneTextMsg()  As String   '-- Holds the PHone Text Msg addresses - Needed becuase the distribution list does not work

'-- Global ticketing system folders set in Help_MakeTicket_Init
Public m_NS                 As NameSpace
Public m_APF                As Folder   '-- All Public Folders in case there are more than 1 Public Folders attached
Public m_fldrCompany        As Folder   '-- Top level folder company
Public m_fldrTickets        As Folder   '-- Folder for Tickets (Tasks)
Public m_fldrMail           As Folder   '-- Folder for Emails (Mail)
Public m_fldrTime           As Folder   '-- Folder for Time (Appointment)
Public m_fldrPart           As Folder   '-- Folder for Parts (Task)
Public m_fldrBackupTickets  As Folder   '-- Folder for Backup Tickets (Tasks)
Public m_fldrLists          As Folder   '-- Folder of lists (e.g. Filing Rules, Assignee)
Public m_fldrTicketNum      As Folder   '-- Folder with Ticket Numbers
Public m_fldrCalendar       As Folder   '-- Folder for the Dispatch Calendar - used to determine Help Desk
Public m_fldrInbox          As Folder   '-- Current Inbox - used to distribute ticketDim m_blnMakeTicket_Init    As Boolean          '-- Flag to prevent reinitializing the variable

'-- Array to hold the Assignee. Array goes from 0 to N
Public m_asAssignees()      As String
Public m_iRowsAssignees     As Integer

Public m_dteErrorStart      As Date
Public m_iErrorCount        As Integer
Public m_sErrorMsg          As String


'--------------------------------------------------------------------------------------------------
'-- Routine to display the error message
'--------------------------------------------------------------------------------------------------
Public Sub HandleError(Optional ByRef sMsgBody As String = "Error Unknown")
    '-- Shell command calling location of executable
    'Dim sGmailTextProgram As String
    'sGmailTextProgram = "C:\Users\help\Desktop\GmailTextAlert.exe" '-- Location of Gmail Text Program Executable
    'Call Shell("""" & sGmailTextProgram & """ """ & sMsgHead & "*" & sMsgBody & """", vbNormalFocus)

    Const ERRHEAD = "Ticket Error: "
    Const ERRMAX = 10
    
    Dim blnShutdown As Boolean
    Dim oMsg As MailItem
    
    '-- Save the error message to check for repeating messages
    If m_sErrorMsg <> sMsgBody Then
        m_sErrorMsg = sMsgBody
        m_dteErrorStart = Now()
        m_iErrorCount = 1
        blnShutdown = False
    Else
        m_iErrorCount = m_iErrorCount + 1
        
        If (m_iErrorCount > ERRMAX) And (DateDiff("s", Now(), m_dteErrorStart) < 2) Then
            sMsgBody = "Error Max Reached: " & ERRMAX & " Outlook Shutdown " & sMsgBody
            blnShutdown = True
        End If
    End If
    
    Debug.Print ERRHEAD & sMsgBody
    
    '-- Send a message to Joe
    Set oMsg = Application.CreateItem(olMailItem)
    oMsg.Subject = ERRHEAD & sMsgBody
    oMsg.Body = ERRHEAD & sMsgBody
    oMsg.Recipients.Add "6109370900@mms.att.net"
    oMsg.Recipients.Add "jawe@techbldrs.com"
    oMsg.Send

    If blnShutdown Then
        '-- Shutdown Outlook
        Application.Quit
    End If

End Sub


'----------------------------------------------------------------------------------------
'-- Init_Public_Globals - Sets up the Public Global level variables, folders and filing rules
'----------------------------------------------------------------------------------------
Public Function Init_Public_Globals() As Boolean
    Dim fldrData    As Folder
    
    Init_Public_Globals = False
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- Get the NameSpace where all the Outlook Folders live
    Set m_NS = Application.GetNamespace("MAPI")
    
    '-- Get the correct Public Folder and favor the Favorites Folder
 '   Set m_APF = GetAllPublicFolders(m_APF, DOMAIN)
 
    Set m_APF = m_NS.Folders(PF & " - " & m_NS.GetDefaultFolder(olFolderInbox).Parent).Folders(APF)
    
    
    Set m_fldrCompany = m_APF.Folders(FLDR_COMPANY)
    
    '------------------------------------------------------------------------
    '-- Point to the folders that will hold all the Tickets, Emails and Time
    '------------------------------------------------------------------------
    Set m_fldrTickets = m_fldrCompany.Folders(FLDR_TICKET)      '-- Task folder
    Set m_fldrMail = m_fldrTickets.Folders(FLDR_TICKET_MAIL)    '-- Mail folder
    Set m_fldrTime = m_fldrTickets.Folders(FLDR_TICKET_TIME)    '-- Appointment folder
    Set m_fldrPart = m_fldrTickets.Folders(FLDR_TICKET_PART)    '-- Task folder
    Set m_fldrBackupTickets = m_fldrTickets.Folders(FLDR_BACKUPTICKET)    '-- Task folder
    
    '------------------------------------------------------------------------
    '-- Point to the Data Folders for Filing Rules, Last Ticket Num, Technician Initials
    '------------------------------------------------------------------------
    Set fldrData = m_fldrCompany.Folders(FLDR_DATA)
    Set m_fldrLists = fldrData.Folders(FLDR_LISTS)
    Set m_fldrTicketNum = fldrData.Folders(FLDR_DATA_TICKETNUM)
    
    '------------------------------------------------------------------------
    '-- Find the Group Calendar, Client Contacts, & current users Inbox
    '------------------------------------------------------------------------
    Set m_fldrCalendar = m_fldrCompany.Folders(FLDR_CALENDAR)
    Set m_fldrInbox = m_NS.GetDefaultFolder(olFolderInbox)
    
    '-- Load the Assignees Names and initials table
    '-- Col 0: Initials, Col 1: Name, Col 2: Email, Col 3: Text Msg Email, Col 4: CrewHu Id
    m_iRowsAssignees = TableLoad_from_Body(m_fldrLists.Items(LIST_ASSIGNEES), m_asAssignees)
    
    '-- Set the flag so that we don't do this again
    Init_Public_Globals = True
    
    '-- Clean up
    Set fldrData = Nothing

    Exit Function
ERRORHANDLER:
    HandleError "In Mod:Rtn HELP00_PUBLIC:Init_Public_Globals [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function



'--------------------------------------------------------------------------------------------------
'-- TableLoad_from_Body - Loads a table array from the body of an item with each row separated by
'--                       carriage return, and each column separated by TKTDELIM.
'--                         Blank rows and rows starting with signal apostrophy (') are ignored.
'--                         All cells have Tab characters removed, and spaces trimmed.
'--  Returns: Number of non-blank rows in the table
'--------------------------------------------------------------------------------------------------
Public Function TableLoad_from_Body(ByRef oItem As Object, _
                                    ByRef asArray() As String) As Integer

    Dim sList       As String
    Dim arrStr()      As String 'Array
    Dim arrStrCols()  As String 'Array
    Dim iRows       As Integer
    Dim iCols       As Integer
    Dim idx         As Integer
    Dim jdx         As Integer
    Dim iNonBlankRows   As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    '-- If there is no source for the list, then get out.
    If oItem Is Nothing Then Exit Function
    
    '-- 1. Take the contents of body and split it up into an array rows
    sList = Trim(oItem.Body)
    sList = Replace(sList, vbLf, "")
    arrStr = Split(sList, vbCr)
    
    '-- Initialize variables
    iRows = UBound(arrStr)
    iCols = -1
    iNonBlankRows = -1
    
    '-- Load each 2D table into an array
    For idx = 0 To iRows

        '-- If the row is non-blank and does not start with a single apostrophy ('), then add it
        If Trim(arrStr(idx)) <> "" Then
            If Mid(arrStr(idx), 1, 1) <> "'" Then
        
                '-- Keep a counter of non-blank rows that are loaded
                iNonBlankRows = iNonBlankRows + 1
                
                '-- 2. Each string needs to be split in columns
                arrStr(idx) = Replace(arrStr(idx), vbTab, " ")
                arrStrCols = Split(arrStr(idx), TKTDELIM)
                
                '-- Resize the array to fit the largest amount of data
                '-- *** Multiple REDIM can only expand the last dimension (columns). ie: you can't change the rows ***
                If iCols < UBound(arrStrCols) Then
                    iCols = UBound(arrStrCols)
                    '-- This can only be done once for a 2D Array
                    ReDim Preserve asArray(iRows, iCols)    '-- This is where the array gets dimensioned.
                End If
                
                '-- Load each element stripping all tabs and trimming spaces
                For jdx = 0 To UBound(arrStrCols)
                    asArray(iNonBlankRows, jdx) = Trim(Replace(arrStrCols(jdx), vbTab, ""))
                Next
            End If
        End If
    Next

    '-- Returning -1 means nothing got loaded
    TableLoad_from_Body = iNonBlankRows
    
    Exit Function
ERRORHANDLER:
    HandleError "In Mod:Rtn HELP00_PUBLIC:TableLoad_from_Body [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function

'--------------------------------------------------------------------------------------------------
'-- TableFind - Searches a Table array, in a column, for a matching search string, and returns the
'--             cooresponding value from another column.
'--------------------------------------------------------------------------------------------------
Public Function TableFind(ByRef asArray() As String, _
                          ByRef iSearchCol As Integer, _
                          ByRef sSearch As String, _
                          ByRef iReturnCol As Integer) As String

    Dim idx As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    TableFind = ""
    
    '-- Loop through the array looking for the search string
    For idx = 0 To UBound(asArray)
    
        '-- If a match is found, then return the element from the cooresponding column
        If StrComp(asArray(idx, iSearchCol), sSearch, vbTextCompare) = 0 Then
            TableFind = asArray(idx, iReturnCol)
            Exit For
        End If
    Next
    
    Exit Function
ERRORHANDLER:
    HandleError "In Mod:Rtn HELP00_PUBLIC:TableFind [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function

 
'------------------------------------------------------------------------------------------------------------
'-- GetAllPublicFolders
'------------------------------------------------------------------------------------------------------------
Public Function GetAllPublicFolders(ByRef fAPF As Folder, ByRef sDomain As String) As Folder
    Dim fTop As Folder
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set GetAllPublicFolders = fAPF
    If fAPF Is Nothing Then
        For Each fTop In Application.GetNamespace("MAPI").Folders
            If Mid(fTop.Name, 1, Len(PF)) = PF Then
                If Right(fTop.Name, Len(sDomain)) = DOMAIN Then
                    Exit For
                End If
            End If
        Next
        Set GetAllPublicFolders = fTop.Folders(APF)
    End If
    Set fTop = Nothing
    
    Exit Function
ERRORHANDLER:
    HandleError "In Mod:Rtn HELP00_PUBLIC:GetAllPublicFolders [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function

'############################################################################################################
'## HELP00_Public - VBA module that holds Public Constants and Variables common to many Ticket Builder
'############################################################################################################
