Attribute VB_Name = "HELP03_MakeTicket_Mail_1s"
'############################################################################################################
'## HELP02_MakeTicket_Mail - VBA Code to create a Mail mail/note item                                      1s
'############################################################################################################

Option Explicit

'-- *** HELP00_Public = Module must be loaded to use Public Constants and Variables ***

'--------------------------------------------------------------------------------------------------
' Help_MakeTicket_Mail - Converts the standard incoming email to a HELP Ticket Email and moves it
'                        to the target mail folder.
'--------------------------------------------------------------------------------------------------
Public Function Help_MakeTicket_Mail(ByRef oMail As MailItem, _
                                     ByRef sClient As String, _
                                     ByRef sTicketNum As String, _
                                     ByRef sTopic As String, _
                                     ByRef sSubject As String, _
                                     ByRef sMsgClass As String, _
                                     ByRef fTarget As Folder) As MailItem
    
    '--$$$ 3 Conditions: 1-Default no Error processing, 2-Ignore Errors, 3-Handle Errors
    If ERR_IGNORE Then
        On Error Resume Next
    ElseIf ERR_HANDLER Then
        On Error GoTo ERRORHANDLER
    End If
    
    Set Help_MakeTicket_Mail = Nothing
    
    If sClient <> "" And sTicketNum <> "" Then
      '--------------------------------------------------------------------------------
      '-- Create a Ticket
      '--------------------------------------------------------------------------------
      '-- Change the Message Class so the custom fields will show and the correct forms
      '-- are use for Replies, Forwards and Posts
      oMail.MessageClass = sMsgClass  '-- maybe ineffective!!!
      
      '-- Add new fields
      oMail.UserProperties.Add MAIL_APPROVAL, olYesNo
      oMail.UserProperties.Add MAIL_CLIENT, olText
      oMail.UserProperties.Add MAIL_TICKETNUM, olText
      oMail.UserProperties.Add MAIL_TOPIC, olText
      
      oMail.UserProperties(MAIL_APPROVAL) = False
      oMail.UserProperties(MAIL_CLIENT) = sClient
      oMail.UserProperties(MAIL_TICKETNUM) = sTicketNum
      oMail.UserProperties(MAIL_TOPIC) = sTopic
      
      '-- Reset importance to Normal even though the incoming email is set to high
      oMail.Importance = olImportanceNormal
      
      '--------------------------------------------------------------------------------
      '-- This field is used to group the emails into like Topics.
      oMail.BillingInformation = ""
      If sClient <> "" Then
          oMail.BillingInformation = TKTDELIM & sClient & TKTDELIM & sTicketNum & TKTDELIM & " "
      End If
      oMail.BillingInformation = oMail.BillingInformation & sTopic
      
      '--------------------------------------------------------------------------------
      oMail.Subject = sSubject
      
      '-- KLUDGE: Trying to prevent conflicts from occuring in Outlook 2010 Public Folder TB Mail
      oMail.UnRead = False
      
      '--------------------------------------------------------------------------------
      Set Help_MakeTicket_Mail = oMail
      
      '--------------------------------------------------------------------------------
    
      oMail.MessageClass = sMsgClass  '-- Outlook Quirk: Requires message class AGAIN only for Mail
      oMail.Move fTarget
    End If
    

    Exit Function
ERRORHANDLER:
    CmHandleError "In Mod:Rtn HELP03_MakeTicket_Mail:Help_MakeTicket_Mail [" & Err.Number & "] " & Err.Description
    If ERR_RESUME Then Resume Next
End Function
