using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parser
{
    class Program
    {
        static void Main(string[] args) {
            // Run the monitoring loop on a separate thread
            Thread monitoringThread = new Thread(MonitorOutlook);
            monitoringThread.Start();

            // Keep the application running
            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }

        static void MonitorOutlook() {
            // Initialize Outlook Application
            Outlook.Application outlookApp = new Outlook.Application();

            // Get Inbox folder
            Outlook.MAPIFolder inbox = outlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            // Infinite loop to continuously monitor emails
            while (true) {
                foreach (object item in inbox.Items) {
                    if (item is Outlook.MailItem) {
                        // Process each email using your logic
                        Outlook.MailItem email = (Outlook.MailItem)item;

                        // Add your email processing logic here

                        // For demonstration purposes, just print the subject
                        Console.WriteLine($"New Email: {email.Subject}");
                    }
                }
                // Sleep for a while before checking for new emails again
                Thread.Sleep(TimeSpan.FromMinutes(1));
            }
        }
    }
}