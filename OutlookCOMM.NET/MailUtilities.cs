using System;
using System.Diagnostics;
using System.Net.Mail;
using OutlookCOMM.Core;

namespace OutlookCOMM.NET
{
    /// <summary>
    /// MailUtilities class
    /// </summary>
    public class MailUtilities : MailUtilitiesBase
    {
        /// <summary>
        /// Constructor which initializes a MailUtilities object with passed information.
        /// <see cref="MailUtilitiesBase.MailUtilitiesBase(string, string, string, string, string, string)"/>
        /// </summary>
        public MailUtilities(string from, string to, string cc, string bcc, string subject, string body) : base(from, to, cc, bcc, subject, body) { }

        /// <summary>
        /// Method which creates an EML file.
        /// <see cref="MailUtilitiesBase.SaveEML"/>
        /// </summary>
        public override bool SaveEML()
        {
            try
            {
                PrepareTempFolder();

                MailMessage message = new MailMessage();
                PrepareMessage(ref message);
                PrepareAttachments(ref message);

                using (SmtpClient smtpClient = new SmtpClient { DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory, PickupDirectoryLocation = TempPath })
                {
                    // Save the created message to the disk instead of calling the default mail client
                    smtpClient.Send(message);
                }
                message.Dispose();

                // Open the finalized EML file with the default mail client set to open *.eml files
                Process.Start(FinalizeEML());
            }
            catch (Exception ex)
            {
                // Write the error into the Windows registries
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(ex.Message + Environment.NewLine + ex.StackTrace, EventLogEntryType.Error);
                }

                return false;
            }

            return true;
        }
    }
}
