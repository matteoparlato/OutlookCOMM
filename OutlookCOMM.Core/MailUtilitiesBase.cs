using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace OutlookCOMM.Core
{
    public abstract class MailUtilitiesBase : IMailUtilities
    {
        public string From { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string AttachmentPath { get; set; }
        public char Delimiter { get; set; } = ';';
        public bool Unsent { get; set; } = true;
        public bool UseOutlookAccount { get; set; } = true;

        internal readonly string TempPath = Path.Combine(Path.GetTempPath(), "OutlookCOMM");

        /// <summary>
        /// Constructor which initializes a MailUtilitiesBase object with passed information.
        /// </summary>
        /// <param name="from">The sender address of the mail</param>
        /// <param name="to">The receiver address (or addresses) of the mail</param>
        /// <param name="cc">The CC address (or addresses) of the mail</param>
        /// <param name="bcc">The BCC address (or addresses) of the mail</param>
        /// <param name="subject">The subject of the mail to send</param>
        /// <param name="body">The body of the mail to send</param>
        /// <param name="attachmentPath">The path of the file to add as an attachment to the mail</param>
        protected MailUtilitiesBase(string from, string to, string cc, string bcc, string subject, string body, string attachmentPath)
        {
            From = from;
            To = to;
            CC = cc;
            BCC = bcc;
            Subject = subject;
            Body = body;
            AttachmentPath = attachmentPath;
        }

        /// <summary>
        /// Method which creates an EML file.
        /// </summary>
        /// <returns></returns>
        public abstract bool SaveEML();

        /// <summary>
        /// Method which prepares the temp folder where the EML file will be stored.
        /// </summary>
        internal void PrepareTempFolder()
        {
            if (Directory.Exists(TempPath))
            {
                // Delete previously created files and folders to avoid possible conflicts
                Directory.Delete(TempPath, true);
            }

            Directory.CreateDirectory(TempPath);
        }

        /// <summary>
        /// Method which initializes passed MailMessage object.
        /// </summary>
        /// <param name="message">The MailMessage to initialize</param>
        internal void PrepareMessage(ref MailMessage message)
        {
            if (Unsent)
            {
                // Allow to open the EML in a *COMPOSE* window
                message.Headers.Add("X-Unsent", "1");
            }

            if (!string.IsNullOrEmpty(From))
            {
                message.From = new MailAddress(From);
            }

            if (UseOutlookAccount)
            {
                // Avoid System.InvalidOperationException
                message.From = new MailAddress("example@example.com");
            }

            if (!string.IsNullOrEmpty(To))
            {
                foreach (string toAddress in Helpers.SplitAddressesByDelimiter(To, Delimiter))
                {
                    message.To.Add(toAddress);
                }
            }

            if (!string.IsNullOrEmpty(CC))
            {
                foreach (string toAddress in Helpers.SplitAddressesByDelimiter(CC, Delimiter))
                {
                    message.CC.Add(toAddress);
                }
            }

            if (!string.IsNullOrEmpty(BCC))
            {
                foreach (string toAddress in Helpers.SplitAddressesByDelimiter(BCC, Delimiter))
                {
                    message.Bcc.Add(toAddress);
                }
            }

            message.Subject = Subject;
            message.SubjectEncoding = Encoding.UTF8;

            message.Body = Body;
            message.BodyEncoding = Encoding.UTF8;

            message.IsBodyHtml = true;
        }

        /// <summary>
        /// Method which adds an attachment to passed MailMessage.
        /// </summary>
        /// <param name="message">The MailMessage where to add the attachment</param>
        internal void PrepareAttachments(ref MailMessage message)
        {
            // Add the attachment to the EML (encoded as base64 string)
            if (!string.IsNullOrEmpty(AttachmentPath))
            {
                if (File.Exists(AttachmentPath))
                {
                    // Avoid System.IO.File.Delete exception by making a copy of the original file
                    string copiedAttachmentPath = Path.Combine(TempPath, Path.GetFileName(AttachmentPath));
                    File.Copy(AttachmentPath, copiedAttachmentPath);
                    message.Attachments.Add(new Attachment(copiedAttachmentPath));
                }
                else
                {
                    throw new FileNotFoundException("The attachment was not found in specified folder.", AttachmentPath);
                }
            }
        }

        /// <summary>
        /// Method which finalizes the created EML file.
        /// </summary>
        internal void FinalizeEML()
        {
            // Remove X-Sender and From headers so that the EML file will be opened in a *COMPOSE* window
            using (StreamReader reader = new StreamReader(new DirectoryInfo(TempPath).GetFiles().ToArray()[0].FullName))
            using (StreamWriter writer = new StreamWriter(Path.Combine(TempPath, "MailToSend.eml")))
            {
                string line;
                if (!UseOutlookAccount)
                {
                    if ((line = reader.ReadToEnd()) != null)
                    {
                        writer.Write(line);
                    }

                    return;
                }

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("X-Sender:") || line.StartsWith("From:"))
                    {
                        continue;
                    }

                    writer.WriteLine(line);
                }
            }
        }
    }
}
