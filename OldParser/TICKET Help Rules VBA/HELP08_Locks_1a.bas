Attribute VB_Name = "HELP08_Locks_1a"
'############################################################################################################
'## HELP08_Locks - VBA Code to handle Ticket Locks for the Ticket Parser                                   1a
'##
'##   1) Since there is no locking of items in Exchange, these functions along with routine in the Ticket
'##      form handle the locking when a Ticket with open by a user.
'##
'##   2) There VBS Code in TB Ticket form (TKT_FOMR_TB_Ticket_xx.bas) that create the Lock when the Ticket
'##      is opened.  They are Lock_Create, Lock_Find and Lock_Clear
'##
'##   3) Locks are created in folder Public Folders \ All Public Folders \ TB Locks
'##
'##   4) All Locks are deleted during the TKTAdmin00_Optimize process when it runs TKTAdmin02_CleanLocks
'##
'############################################################################################################

Option Explicit

'-- *** TICKET_00_COMMON = Module must be loaded to use Public Constants and Variables ***


'----------------------------------------------------------------------------------------------------------
'-- Lock_Find - Find lock if it exist in the TB Locks folder
'--     Subject = Client|TicketNum
'--     BillingInformation = User
'--     ReminderTime = DateLastActivity
'--     Role = Ticket Status
'----------------------------------------------------------------------------------------------------------
Public Function Lock_Find(ByRef cLocks As Items, ByRef sClient As String, ByRef sTicketNum As String) As TaskItem
    
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
    CmHandleError "HELP08_Locks:Lock_Find [" & Err.Number & "] " & Err.Description & " >" & Lock_Find.Subject
    If ERR_RESUME Then Resume Next
End Function


'----------------------------------------------------------------------------------------------------------
'-- Lock_RemoveOld - For Locks that don't get removed, remove any Lock older than 90 minutes
'----------------------------------------------------------------------------------------------------------
Public Sub Lock_RemoveOld()
    Dim cLocks                  As Items
    Dim oLock                   As TaskItem
    Dim asClient_TicketNum()    As String
    Dim oTicket                 As TaskItem
    Dim idx                     As Integer
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set m_fldrLocks = CmGetPubFavFolder(m_fldrCompany, FLDR_LOCKS)        '-- Task Folder
    Set m_fldrTickets = CmGetPubFavFolder(m_fldrCompany, FLDR_TICKET)      '-- Task folder

    Set cLocks = m_fldrLocks.Items

    For idx = cLocks.Count To 1 Step -1
    
        Set oLock = cLocks.Item(idx)
        
        '-- Delete Locks with no User
        If oLock.Subject = "" Then
            oLock.Delete

        '-- If it is after 6:00pm
        ElseIf Time > TimeValue("18:00:00") Then
        
            '-- If the Lock has been idle, then Update the ticket and delete the Lock
            If DateAdd("n", 90, oLock.LastModificationTime) < Now Then
            
                '-- Split the Subject into Client and TicketNum
                asClient_TicketNum = Split(oLock.Subject, "|")
                
                '-- Find the Ticket matching the Client and TicketNum
                '--     asClient_TicketNum(0) = Client
                '--     asClient_TicketNum(1) = TicketNum
                Set oTicket = CmFindMatchingTask(m_fldrTickets.Items, asClient_TicketNum(0), asClient_TicketNum(1))
        
                '-- Ticket Not Found, then DELETE Lock
                If oTicket Is Nothing Then
                    oLock.Delete
                
                '-- Ticket Found matching an Old Lock, then Update the Ticket and delete the Lock
                Else
                    '-- Update the Ticket
                    If oLock.ReminderTime <> NODATE And oTicket.UserProperties(TKT_DATE_LAST_ACTIVITY) < oLock.ReminderTime Then
                        oTicket.UserProperties(TKT_DATE_LAST_ACTIVITY) = oLock.ReminderTime
                        
                        If oLock.Role <> "" Then
                            oTicket.UserProperties(TKT_STATUS) = oLock.Role
                        End If
                    End If
                    
                    oTicket.Save
                    
                    oLock.Delete
                End If
            End If
        End If
                
    Next

    Set cLocks = Nothing
    Set oLock = Nothing
    Set oTicket = Nothing
    Exit Sub
ERRORHANDLER:
    CmHandleError "HELP08_Locks:Lock_Find [" & Err.Number & "] " & Err.Description & " >" & oLock.Subject
    If ERR_RESUME Then Resume Next

End Sub
