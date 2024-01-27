using System;
using System.Collections.Generic;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

public class TICKET_00_COMMON
{
    //##################################################################################################
    //## TICKET_00_COMMON - Common Public Constants, Variables & Routines for the Ticket System       2t
    //##################################################################################################

    // Global Constants used in the Ticket System
    public const bool ERR_IGNORE = false;    // (FALSE) Turns on or off the ERR_IGNORE Run-time errors
    public const bool ERR_HANDLER = true;    // (TRUE) Turns on using the ErrorHandler Jump point
    public const bool ERR_RESUME = true;     // (TRUE) Turns on whether to continue to the line after the error point, or stop

    // Folders that hold the data for the Ticket System
    public const string FLDR_COMPANY = "TechBldrs Inc";                // Top-level Parent folder
    public const string FLDR_ARCHIVE_TOP = "TB Ticket Archives";       // Top-level Parent folder for all archives
    public const string FLDR_TICKET = "TB Tickets";                    // +- Ticket Folder (Tasks)
    public const string FLDR_BACKUPTICKET = "TB Backup Tickets";       // +- Ticket Backup Folder (Tasks) - used to recover system randomly deleted tickets
    public const string FLDR_TICKET_MAIL = "TB Mail";                  //     +- Email Folder (Mail) for Tickets
    public const string FLDR_TICKET_TIME = "TB Time";                  //     +- Time Folder (Appointments) for Tickets
    public const string FLDR_TICKET_PART = "TB Part";                  //     +- Parts Folder (Tasks) for Tickets
    public const string FLDR_LOCKS = "TB Locks";                       // +- Folder with locked tickets
    public const string FLDR_DATA = "_#DATA#";                        // +- Folder with the following subfolders:
    public const string FLDR_LISTS = "Lists";                         //     +- Folder of lists for the Ticket Form and Help Filing
    public const string FLDR_TEMPLATES = "Templates";                 //     +- Folder of email templates
    public const string FLDR_DATA_TICKETNUM = "Clients & Ticket Numbers"; //     +- Folder with Last Ticket Numbers

    // Tables needed for Help Ticket Processing
    public const string LIST_FILING_RULES = "Help.FilingRules";               // Item in the Company\Data\Lists folder
    public const int FILERULE_CLIENT = 0;                                      // Column of the Filing Rule Table
    public const int FILERULE_MATCH_TEXT = 1;                                  // Column of the Filing Rule Table
    public const int FILERULE_COMPANY_NAME = 2;                               // Column of the Filing Rule Table
    public const string LIST_ASSIGNEES = ".Assignee";                         // Item in the Company\Data\Lists folder
    public const int ASSIGNEE_INITIALS = 0;
    public const int ASSIGNEE_NAME = 1;
    public const int ASSIGNEE_EMAIL = 2;

    // Table need for Help Ticket Processing (for the Flag Rules, Alerts, etc)
    public const string LIST_SPAM_RULES = "Help.SpamRules";
    public const int SPAM_PHRASES = 0;

    public const string LIST_FLAG_RULES = "Help.FlagRules";
    public const int FLAG_RULE_SYMBOL = 0;

    public const string LIST_ALERT_RULES = "Help.AlertRules";
    public const int ALERT_RULE_INDICATOR = 0;

    public const int RULE_CATEGORY = 1;
    public const int RULE_STATUS = 2;
    public const int RULE_REASON = 3;
    public const int RULE_ASSIGNEE = 4;
    public const int RULE_CAUSE = 5;

    // TechBldrs possible domains
    public const string TB1 = "@tecmanage.com";
    public const string TB2 = "@techbldrs.com";
    public const string TB3 = "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/";

    // Alerts to check for in the Subject of the Mail
    public const string ALERT_Tag = "?c?";                   // Alerts from Kaseya
    public const string ALERT_Backup = "Backup Failed!";
    public const string ALERT_Backup2 = "Check Backup";

    // Google Voice Voicemails
    public const string TALK_VOICEMAIL = "Google Voice";
    public const string TALK_GOOG = "GOOG";

    // No Client Defaults
    public const string NOCLIENT_CLIENT = "<none>";
    public const string NOCLIENT_TICKETNUM = "0000";
    //public const string NOCLIENT_TOPIC = "*** Need to File ****";

    // Forms used in Ticket Builder
    public const string MSGCLS_Note = "IPM.Note";
    public const string MSGCLS_MtgRequest = "IPM.Schedule.Meeting.Request";
    public const string MSGCLS_Ticket = "IPM.Task.TB_Ticket";
    public const string MSGCLS_Mail = "IPM.Note.TB_Mail";
    public const string MSGCLS_Time = "IPM.Appointment.TB_Time";
    public const string MSGCLS_Part = "IPM.Task.TB_Part";
    public const string MSGCLS_Reply = "IPM.Note.TB_Reply";

    // Common
    public const string TKTDELIM = "|";                                     // Delimiter used in the Ticket Header
    public const string Tkt_DELETED_Substr = "-2DEL-";       // Set in TICKET_MergeSelectedTickets when tickets are merged
    public const int RESPONSE_DAYS = 2;                                  // Days to respond to a Ticket
    public static DateTime NODATE = new DateTime(4501, 1, 1);           // This is a null date

    // These are User Defined fields and for consistency are also the name of the form object.
    // USED IN: (A) Ticket Parser - Help System VBA code to create the Ticket task item (B) Ticket Admin tools (C) copied into TB_Ticket forms VBS
    // BUILT-IN Fields used with the Tickets

    public const string TKT_ASSIGNEE = ".Assignee";
    public const string TKT_CAUSE1 = ".Cause1";
    public const string TKT_CLIENT = ".Client";
    public const string TKT_DATE_CREATED = ".DateCreated";
    public const string TKT_DATE_LAST_ACTIVITY = ".DateLastActivity";
    public const string TKT_DATE_MODIFIED = ".DateModified";
    public const string TKT_HRS_ACTUAL_TOTAL = ".HrsActualTotal";
    public const string TKT_HRS_BILLABLE_TOTAL = ".HrsBillableTotal";
    public const string TKT_HRS_DURATION = ".HrsDuration";
    public const string TKT_HRS_GRATIS_TOTAL = ".HrsGratisTotal";
    public const string TKT_HRS_FIRST_TOUCH = ".HrsFirstTouch";
    public const string TKT_INVOICE_NUM = ".InvoiceNum";
    public const string TKT_JOB = ".Job";
    public const string TKT_LOG = ".Log";
    public const string TKT_MACHINE_NAME = ".MachineName";
    public const string TKT_MACHINE_SUPPORT = ".MachineSupport";
    public const string TKT_MAIL_TEMPLATE = ".MailTemplate";
    public const string TKT_PROJECT = ".Project";
    public const string TKT_REASON = ".Reason";
    public const string TKT_REQUESTOR = ".Requestor";
    public const string TKT_STATUS = ".Status";
    public const string TKT_TECHNAME = ".Tech";
    public const string TKT_TICKETMONTH = ".TicketMonth";
    public const string TKT_TICKETNUM = ".TicketNum";
    public const string TKT_TICKETYEAR = ".TicketYear";
    public const string TKT_TOPIC = ".Topic";
    public const string TKT_UNIQUEID = ".UniqueID";

    public const string TKT_STATUS_NEW = "New";
    public const string TKT_STATUS_TO_CLIENT = "Email Sent";
    public const string TKT_STATUS_FROM_CLIENT = "Client Replied";

    public const string TKT_CAT0_URGENT = "0 Urgent";
    public const string TKT_CAT1_HIGH = "1 High";
    public const string TKT_CAT1_REOPENED = "1 Re-Opened";
    public const string TKT_CAT2_NORMAL = "2 Normal";
    public const string TKT_CAT3_FOLLOWUP = "3 Follow Up";
    public const string TKT_CAT4_BACKUP = "4 Backup";
    public const string TKT_CAT5_ONSITE = "5 On-Site";
    public const string TKT_CAT6_PROJECT = "6 Project";
    public const string TKT_CAT7_ORDERED = "7 Ordered";
    public const string TKT_CAT8_TIME = "8 Time";
    public const string TKT_CAT9_REVIEW = "9 REVIEW";
    public const string TKT_CAT_QUOTED = "Quoted";

    public const string TKT_REASON_SUPPORT = "Support";
    public const string TKT_REASON_BILLABLE = "Billable";
    public const string TKT_REASON_RESOLVED = "Resolved";
    public const string TKT_REASON_INTERNALPROJECT = "InternalProject";
    public const string TKT_REASON_ADMIN = "Admin";
    public const string TKT_REASON_ALERT = "Alert";

    public const string TKT_ACTION_QUOTED = "Quoted";

    // These are User Defined fields and for consistency are also the name of the form object.
    // USED IN: (A) Ticket Parser - Help System VBA code to create the Ticket task item (B) Ticket Admin tools (C) copied into TB_Ticket forms VBS
    public const string MAIL_APPROVAL = ".Approval";
    public const string MAIL_CLIENT = ".Client";
    public const string MAIL_DATE_CREATED = ".DateCreated";
    public const string MAIL_TICKETNUM = ".TicketNum";
    public const string MAIL_TOPIC = ".Topic";

    public const string MAILTOPIC_Quoted = "TechBldrs Quote for";
    public const string MAILTOPIC_ToQuote = "TO QUOTE";

    public const string PART_TOPIC = ".Topic";

    // These are User Defined fields and for consistency are also the name of the form object.
    // USED IN: (A) Ticket Parser - Help System VBA code to create the Ticket task item (B) Ticket Admin tools (C) copied into TB_Ticket forms VBS
    public const string TIME_BILLABLE = ".Billable";              // Interactive-only Flag
    public const string TIME_BILLEND = ".BillEnd";
    public const string TIME_BILLHOURS = ".BillHours";
    public const string TIME_BILLSTART = ".BillStart";
    public const string TIME_CLIENT = ".Client";
    public const string TIME_DATE_CREATED = ".DateCreated";
    public const string TIME_GRATIS = ".Gratis";                  // Interactive-only Flag
    public const string TIME_HOURS = ".Hours";
    public const string TIME_INVOICE_DESC = ".InvoiceDesc";        // Used during Invoice creation via Outlook macro
    public const string TIME_INVOICE_NUM = ".InvoiceNum";          // Filled in during Invoice creation via Outlook macro
    public const string TIME_JOB = ".Job";                         // Filled in during Invoice creation via Outlook macro
    public const string TIME_QUOTED = ".Quoted";                   // Interactive-only Flag
    public const string TIME_REVIEWED = ".REVIEWED";               // NOT USED: Interactive-only Flag
    public const string TIME_TECH = ".Tech";
    public const string TIME_TICKETNUM = ".TicketNum";
    public const string TIME_TOPIC = ".Topic";
    public const string TIME_UniqueID = ".UniqueID";
    public const string TIME_WORKDATE = ".WorkDate";

    //**************************************************************************************************
    // Global Variables used in the Ticket System
    //**************************************************************************************************

    // Global ticketing system folders set in Help_MakeTicket_Init
    // Global ticketing system folders set in Help_MakeTicket_Init
    public static Outlook.MAPIFolder m_fldrCompany;        // Top level folder company
    public static Outlook.MAPIFolder m_fldrTickets;        // Folder for Tickets (Tasks)
    public static Outlook.MAPIFolder m_fldrMail;           // Folder for Emails (Mail)
    public static Outlook.MAPIFolder m_fldrTime;           // Folder for Time (Appointment)
    public static Outlook.MAPIFolder m_fldrPart;           // Folder for Parts (Task)
    public static Outlook.MAPIFolder m_fldrBackupTickets;  // Folder for Backup Tickets (Tasks)
    public static Outlook.MAPIFolder m_fldrLists;          // Folder of lists (e.g., Filing Rules, Assignee)
    public static Outlook.MAPIFolder m_fldrTicketNum;      // Folder with Ticket Numbers
    public static Outlook.MAPIFolder m_fldrCalendar;       // Folder for the Dispatch Calendar - used to determine Help Desk
    public static Outlook.MAPIFolder m_fldrLocks;          // Folder for open tickets

    // Array to hold the Assignee. Array goes from 0 to N
    public static string[] m_asAssignees;
    public static int m_iRowsAssignees;

    public static DateTime m_dteErrorStart;
    public static int m_iErrorCount;
    public static string m_sErrorMsg;

	// Constants
	private const string ERRHEAD = " TktErr: ";
	private const int ERRMAX = 10;

	public static void CmHandleError(string sMsgBody = "Error Unknown") {

		// Save the error message to check for repeating messages
		bool blnShutdown;
		Outlook.MailItem oMsg;

		if (m_sErrorMsg != sMsgBody)
		{
			m_sErrorMsg = sMsgBody;
			m_dteErrorStart = DateTime.Now;
			m_iErrorCount = 1;
			blnShutdown = true; //#### False 2022-01-17
		}
		else
		{
			m_iErrorCount++;

			if (m_iErrorCount > ERRMAX && (DateTime.Now - m_dteErrorStart).TotalSeconds < 2)
			{
				sMsgBody = "Error Max Reached: " + ERRMAX + " Outlook Shutdown " + sMsgBody;
				blnShutdown = true;
			}
			else
			{
				blnShutdown = false;
			}
		}

		if (m_sErrorMsg.Contains("Network problems are preventing ") ||
			m_sErrorMsg.Contains("Object variable or With block variable not "))
		{
			blnShutdown = true;
		}

		Debug.Print(DateTime.Now + ERRHEAD + sMsgBody);

		// Send a message to Joe
		oMsg = (Outlook.MailItem) new Outlook.Application().CreateItem(Outlook.OlItemType.olMailItem)
;
		if (oMsg != null) {
			oMsg.Subject = ERRHEAD + sMsgBody;
			oMsg.Body = ERRHEAD + sMsgBody;
			oMsg.Recipients.Add("Help@techBldrs.com");
			oMsg.Recipients.Add("jawe@techbldrs.com");
			oMsg.Send();
		}

		if (blnShutdown)
		{
			// Shutdown Outlook
			new Outlook.Application().Quit();
		}
	}
}
