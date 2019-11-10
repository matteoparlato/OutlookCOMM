using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using OutlookCOMM.Core;

namespace OutlookCOMM.COM
{
    [ComVisible(true)]
    [Guid("C79C6ABA-10F6-4DEA-B9AE-69DDB62C5881")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("MailUtilities")]  
    public class MailUtilities : MailUtilitiesBase
    {
        /// <summary>
        /// Constructor which initializes a MailUtilities object with passed information.
        /// <see cref="MailUtilitiesBase(string, string, string, string, string, string, string)"/>
        /// </summary>
        public MailUtilities(string from, string to, string cc, string bcc, string subject, string body) : base(from, to, cc, bcc, subject, body) { }

        /// <summary>
        /// Constructor with no parameters required for COM initialization.
        /// </summary>
        public MailUtilities()
        {
            //
        } 

        /// <summary>
        /// 
        /// <see cref="MailUtilitiesBase(string, string, string, string, string, string, string)"/>
        /// </summary>
        /// <returns>An instance of MailUtilities</returns>
        public MailUtilities CreateInstance(string from, string to, string cc, string bcc, string subject, string body)
        {
            return new MailUtilities(from, to, cc, bcc, subject, body);
        }

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

                SmtpClient smtpClient = new SmtpClient
                {
                    DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
                    PickupDirectoryLocation = TempPath
                };

                // Save the created message to the disk instead of calling the default mail client
                smtpClient.Send(message);

                message.Dispose();

                FinalizeEML();

                // Open the EML file with the default mail client set to open *.eml files
                Process.Start(Path.Combine(TempPath, "MailToSend.eml"));
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
