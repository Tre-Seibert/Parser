Attribute VB_Name = "HELP07_ReturnTicketList_1m"
'############################################################################################################
'## HELP07_ReturnTicketList - VBA Code to return a list of tickets                                         1m
'############################################################################################################

Option Explicit

'-- *** TICKET_00_COMMON = Module must be loaded to use Public Constants and Variables ***

'--------------------------------------------------------------------------------------------------
'-- HELP_ReturnTicketList - Emails that start with a "?" in the subject trigger that it is an
'--                         alert, or a request for a list of tickets per client or assignee.
'--                         ?  - requests OPEN tickets
'--                         ?? - requests ALL tickets
'--                            - following the ? can be client abbreviation or assignee initials
'-- Returns TRUE if a message is sent
'--------------------------------------------------------------------------------------------------
Public Function HELP_ReturnTicketList(ByRef oMail As Object) As Boolean
    Dim sSubject    As String
    Dim blnRtnAll   As Boolean
    Dim sTemp       As String
    Dim sFilter     As String
    Dim cTickets    As Items
    Dim oMsg        As MailItem
    Dim idx         As Integer
    Dim oTicket     As TaskItem
    Dim sStatus     As String
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    HELP_ReturnTicketList = False
    
    '-- If the subject has an Alert in it, then get out
    sSubject = oMail.Subject
    If InStr(1, sSubject, ALERT_Tag, vbTextCompare) > 0 Then Exit Function
    
    '-- Double question mark means return all
    If Left(sSubject, 2) = "??" Then
        sTemp = Mid(sSubject, 3, Len(sSubject) - 2)
        blnRtnAll = True
    '-- Single question mark means return only open tickets, not 9 Review, Future, Quoted
    ElseIf Left(sSubject, 1) = "?" Then
        sTemp = Mid(sSubject, 2, Len(sSubject) - 1)
        blnRtnAll = False
    End If
    sTemp = Trim(sTemp)
    
    '-- Build the filter for Client or Assignee
    If Len(sTemp) > 2 Then
        sFilter = "[.Client] = '" & sTemp & "'"
    Else
        sFilter = "[.Assignee] = '" & sTemp & "'"
    End If
    
    '-- Get the Tickets
    Set cTickets = m_fldrTickets.Items.Restrict(sFilter)
    cTickets.Sort "[Subject]"
    
    '-- Create the message to return
    Set oMsg = Application.CreateItem(olMailItem)
    oMsg.Recipients.Add oMail.SenderName
    oMsg.Subject = sTemp & " Tickets"
    oMsg.BodyFormat = olFormatRichText
    
    '-- Create the body of the message
    For idx = 1 To cTickets.Count
        Set oTicket = cTickets.Item(idx)
        
        '-- Either include only the open tickets, or everything
        sStatus = Left(oTicket.Categories, 1)
        If (("1" <= sStatus And sStatus <= "8") And Not oTicket.Complete) Or blnRtnAll Then
            oMsg.Body = Trim(oTicket.Subject) & " [" & oTicket.Categories & "] (" & oTicket.UserProperties(".Assignee") & ")" _
                        & vbCrLf & oMsg.Body
        End If
    Next

    '-- Send the email
    oMsg.sEnd
    
    oMail.Delete
    HELP_ReturnTicketList = True

    Set cTickets = Nothing
    Set oMsg = Nothing
    Set oTicket = Nothing
    
    Exit Function
ERRORHANDLER:
    CmHandleError "HELP07_ReturnTicketList:HELP_ReturnTicketList [" & Err.Number & "] " & Err.Description & " >" & sSubject
    
    If ERR_RESUME Then Resume Next
End Function

