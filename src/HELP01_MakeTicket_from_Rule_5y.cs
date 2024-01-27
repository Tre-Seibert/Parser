// **************************************************
//                   Imports
// **************************************************
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;

// '## TICKET_00_COMMON - Common Public Constants, Variables & Routines for the Ticket System


//'--------------------------------------------------------------------------------------------------
//' These C# functions are designed for the HELP email account when it receives new emails.  Each
//' email is tested against a list of "Filing Rules", to find the proper Client and Ticket Number.
//' If the Ticket exists, then IPM.Note.TB_Mail form is filled out.  If not, then both the
//' IPM.Task.TB_Ticket and IPM.Note.TB_Mail forms are created.
//'
//' HELP_MakeTicket           - Triggered when a new email is recievd in the HELP inbox
//' HELP_ProcessEmail         - Processes mail that comes in
//' HELP_ProcessTime          - Processes time that gets submitted
//' HELP_MakeTicket_Init      - Initializes all the module level variables to do once per Outlook session.
//'
//' The "Filing Rules" are managed in a Outlook folder or Lists.
//'   Client  - Client abbreviation (4 to 8 characters)
//'   Subject - Value to look for in the Sender Address and Subject
//'   Company - Company to assign to the Task (used only for description)
//'
//' The Clients and Ticket Numbers are managed in a separate Outlook Task folder.  Each taskitem is a
//' Client record with the following fields populated...
//'   Client (Subject)                    - Client abbreviation (4 to 8 characters)
//'   Ticket Number (Billing Information) - 4 Digit number used for the Ticket Number
//'   Help Email (Contact)                - The HELP email to send the reply so that it gets recorded
//'                                         back into the Ticket system
//'--------------------------------------------------------------------------------------------------

// Module Level Constants
// '-- Errors for rejecting a time entry 
const string ERR_BAD_TICKET_HEADER = "No Ticket Header found.\r\nCorrect the Subject and add |<Client>|<Ticket#>|";
const string ERR_LONG_DURATION = "The duration exceeds 24 hours.\r\nCorrect Start & End dates and times";
const string ERR_NO_TIME_DESC = "There is no detailed description for the time entry.\r\nEnter the work that was performed.";
const string ERR_NO_TICKET = "There is no Ticket matching the ticket header.";
const string ERR_NO_TIME_TOPIC = "Topic field must be filled in.";

// '-- Module Level Variables because everytime the rule gets triggered, these do not have to be reinitialized.
bool m_blnMakeTicket_Init = false;

// '-- Array to hold the Filing Rules. Array goes from 0 to N
string[,] m_asFilingRules;
int m_iRowsRules;

// '-- Array to hold rules for flags
string[,] m_asFlagRules;
int m_iRowsFlagRules;
int m_iColsFlagRules;

// '-- Array to hold rules for alerts
string[,] m_asAlertRules;
int m_iRowsAlertRules;

// '-- Array to hold spam rules
string[,] m_asSpamRules;
int m_iRowsSpamRules;

namespace Parser.src
{
    class Parser
    {

        // Make these members static
        static bool m_blnMakeTicket_Init = false;
        static Folder m_fldrTickets;
        static Folder m_fldrLocks;


        static void Main(string[] args)
        {
            // Run the monitoring loop on a separate thread
            Thread monitoringThread = new Thread(MonitorOutlook);
            monitoringThread.Start();

            // Keep the application running
            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }

        static void MonitorOutlook()
        {
            // Initialize Outlook Application
            Application outlookApp = new Application();

            // Get Inbox folder
            MAPIFolder inbox = outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // Infinite loop to continuously monitor emails
            while (true)
            {
                foreach (object item in inbox.Items)
                {
                    if (item is MailItem)
                    {
                        // Process each email using your logic
                        MailItem email = (MailItem)item;

                        // Add your email processing logic here
                        HELP_MakeTicket(email);

                        // For demonstration purposes, just print the subject
                        Console.WriteLine($"New Emails: {email.Subject}");
                    }
                }
                // Sleep for a while before checking for new emails again
                Thread.Sleep(TimeSpan.FromMinutes(1));
            }
        }

        // HELP_MakeTicket - It checks each new mail and decides what to process.
        public static void HELP_MakeTicket(MailItem oItem)
        {
            MailItem oMail = null;
            MeetingItem oMtgReq = null;
            bool blnRtnList = false;

            try
            {
                // Initialize all the module level variables only the first time
                if (!m_blnMakeTicket_Init)
                {
                    m_blnMakeTicket_Init = HELP_MakeTicket_Init();
                }

                // Email entries - Make sure the email is a type we can process
                if (oItem is MailItem && (string.Equals(oItem.MessageClass, MSGCLS_Note, StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(oItem.MessageClass, MSGCLS_Mail, StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(oItem.MessageClass, MSGCLS_Reply, StringComparison.OrdinalIgnoreCase)))
                {
                    oMail = oItem;

                    // If the trigger email starts with a ?, return the list of tickets
                    blnRtnList = false;
                    if (oMail.Subject.StartsWith("?"))
                    {
                        blnRtnList = HELP_ReturnTicketList(oMail);
                    }

                    // If a list of tickets was not sent (e.g., Alert email), then process the email
                    if (!blnRtnList)
                    {
                        // Process HELP Emails
                        HELP_ProcessEmail(oMail);
                    }
                }
                // Time entries
                else if (string.Equals((oItem as MeetingItem).MessageClass, MSGCLS_MtgRequest, StringComparison.OrdinalIgnoreCase))
                {
                    oMtgReq = (MeetingItem)oItem;
                    // Accept Time emails
                    HELP_ProcessTime(oMtgReq);
                }

                // Do the HeartBeat processing
                Lock_RemoveOld();
                // HeartBeat(m_fldrTickets); // MOVED TO separate EXE nightly batch run
            }
            catch (System.Exception ex)
            {
                TICKET_00_COMMON TICKET_00_COMMON = new TICKET_00_COMMON();
                TICKET_00_COMMON.CmHandleError("HELP01_MakeTicket_from_Rule:HELP_MakeTicket", $"{ex.Message} >{(oMail != null ? oMail.Subject : oMtgReq != null ? oMtgReq.Subject : string.Empty)}");
                if (TICKET_00_COMMON.ERR_RESUME)
                {
                    // Resume Next
                }
            }
            finally
            {
                Marshal.ReleaseComObject(oMail);
                Marshal.ReleaseComObject(oMtgReq);
            }
        }


    }
}
