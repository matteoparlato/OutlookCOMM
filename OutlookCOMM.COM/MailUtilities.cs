using System;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;

namespace OutlookCOMM.COM
{
    [Guid("69C85C8D-DBEB-4d85-83A7-7E5077AD11BA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOutlookCOMM
    {
        /// <summary>
        /// Method which creates an EML file with passed information.
        /// </summary>
        /// <param name="from">The sender address of the mail</param>
        /// <param name="to">The receiver address of the mail </param>
        /// <param name="subject">The subject of the mail to send</param>
        /// <param name="body">The body of the mail to send</param>
        /// <param name="attachmentPath">The path of the attachment to add to the mail</param>
        /// <param name="unsent">Set the X-Unsent property of the EML</param>
        [DispId(1)] bool SaveEML(string from, string to, string cc, string bcc, string subject, string body, string attachmentPath, bool unsent, bool useOutlookAccount);
    }
    
    [Guid("0C216A19-E1B7-4b05-86D3-4C516BDDC041")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MailUtilities")]  
    public class MailUtilities:IOutlookCOMM
    {
        /// <summary>
        /// Method which creates an EML file with passed information.
        /// </summary>
        /// <param name="from">The sender address of the mail</param>
        /// <param name="to">The receiver address of the mail </param>
        /// <param name="subject">The subject of the mail to send</param>
        /// <param name="body">The body of the mail to send</param>
        /// <param name="attachmentPath">The path of the attachment to add to the mail</param>
        /// <param name="unsent">Set the X-Unsent property of the EML</param>
        public bool SaveEML(string from, string to, string cc, string bcc, string subject, string body, string attachmentPath, bool unsent, bool useOutlookAccount)
        {
            string tempFolderPath = Path.Combine(Path.GetTempPath(), "OutlookCOMM");

            try
            {
                if (Directory.Exists(tempFolderPath))
                {
                    // Delete previously created files and folders
                    Directory.Delete(tempFolderPath, true);
                }
                // Create temporary folder where to place the eml file
                Directory.CreateDirectory(tempFolderPath);

                MailMessage message = new MailMessage();

                // When set unsent = true the EML will be opened in a *COMPOSE* window (NOT a display one)
                if (unsent)
                {
                    message.Headers.Add("X-Unsent", "1");
                }

                // Add mail addresses
                if (!string.IsNullOrEmpty(from))
                    message.From = new MailAddress(from);
                if (useOutlookAccount)
                    message.From = new MailAddress("example@example.com");
                if (!string.IsNullOrEmpty(to))
                    message.To.Add(to);
                if (!string.IsNullOrEmpty(cc))
                    message.CC.Add(cc);
                if (!string.IsNullOrEmpty(bcc))
                    message.Bcc.Add(bcc);

                // Add subject and body to 
                message.Subject = subject;
                message.SubjectEncoding = Encoding.UTF8;
                message.Body = body;
                message.BodyEncoding = Encoding.UTF8;

                // Add the attachment to the EML (encoded as base64 string)
                if (!string.IsNullOrEmpty(attachmentPath))
                {
                    if (File.Exists(attachmentPath))
                    {
                        // Avoid System.IO.File.Delete exception
                        string newAttachmentPath = Path.Combine(tempFolderPath, Path.GetFileName(attachmentPath));
                        File.Copy(attachmentPath, newAttachmentPath);
                        message.Attachments.Add(new Attachment(newAttachmentPath));
                    }
                    else
                    {
                        throw new FileNotFoundException("The attachment was not found in specified folder.", attachmentPath);
                    }
                }

                // Always use HTML body format
                message.IsBodyHtml = true;

                SmtpClient smtpClient = new SmtpClient
                {
                    DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
                    PickupDirectoryLocation = tempFolderPath
                };
                // Save the created message to the disk instead of calling the default mail client
                smtpClient.Send(message);

                // Remove X-Sender and From information so that the EML file will be opened in a *COMPOSE* window (NOT a display one)
                using (StreamReader reader = new StreamReader(new DirectoryInfo(tempFolderPath).GetFiles().ToArray()[0].FullName))
                using (StreamWriter writer = new StreamWriter(Path.Combine(tempFolderPath, "MailToSend.eml")))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (useOutlookAccount)
                        {
                            if (line.StartsWith("X-Sender:") || line.StartsWith("From:"))
                                continue;
                        }

                        writer.WriteLine(line);
                    }
                }

                // Open the eml file with the default mail client set to open *.eml files
                Process.Start(Path.Combine(tempFolderPath, "MailToSend.eml"));
            }
            catch (Exception ex)
            {
                // Write the error into the Windows registries - interactive mode not supported in RTC
                using (EventLog log = new EventLog("Application"))
                {
                    log.Source = "Application";
                    log.WriteEntry(ex.Message + Environment.NewLine + ex.StackTrace, EventLogEntryType.Error);
                }

                return false;
            }

            return true;
        }
    }
}