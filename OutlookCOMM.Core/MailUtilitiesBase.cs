using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookCOMM.Core
{
    /// <summary>
    /// MailUtilitiesBase abstract class
    /// </summary>
    public abstract class MailUtilitiesBase : IMailUtilities
    {
        /// <summary>
        /// The string which contains the from address of the mail.
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// The string which contains the recipients of the mail.
        /// </summary>
        public string To { get; set; }

        /// <summary>
        /// The string which contains the carbon copy (CC) recipients of the mail.
        /// </summary>
        public string CC { get; set; }

        /// <summary>
        /// The string which contains the blind carbon copy (BCC) recipients of the mail.
        /// </summary>
        public string BCC { get; set; }

        /// <summary>
        /// The string which contains the subject line of the mail.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// The string which contains the body text of the mail.
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// The attachment collection used to store data attached to the mail.
        /// </summary>
        /// <value>
        /// The key contains the path of the file added as an attachment while the value contains
        /// the text to use as file name for the attachment (can be empty).
        /// </value>
        [ComVisible(false)]
        public List<KeyValuePair<string, string>> Attachments { get; } = new List<KeyValuePair<string, string>>();

        /// <summary>
        /// The string which contains the delimiter to use when splitting the addresses.
        /// </summary>
        /// <value>
        /// The default value is ';'.
        /// </value>
        public char Delimiter { get; set; } = ';';

        /// <summary>
        /// The boolean which determine whether is necessary to add the X-Unsent header to the mail.
        /// </summary>
        /// <value>
        /// The default value is true.
        /// </value>
        public bool Unsent { get; set; } = true;

        /// <summary>
        /// The boolean which determine whether is necessary to use the From address configured in Outlook. 
        /// </summary>
        /// <value>
        /// The default value is true.
        /// </value>
        public bool UseOutlookAccount { get; set; } = true;

        internal readonly string TempPath = Path.Combine(Path.GetTempPath(), "OutlookCOMM");

        /// <summary>
        /// Constructor with no parameters required for COM initialization.
        /// </summary>
        public MailUtilitiesBase()
        {
            //
        }

        /// <summary>
        /// Constructor which initializes a MailUtilitiesBase object with passed information.
        /// </summary>
        /// <param name="from">The from address of the mail</param>
        /// <param name="to">The recipients address (or addresses) of the mail</param>
        /// <param name="cc">The CC recipients of the mail</param>
        /// <param name="bcc">The BCC recipients of the mail</param>
        /// <param name="subject">The subject line of the mail</param>
        /// <param name="body">The body text of the mail</param>
        protected MailUtilitiesBase(string from, string to, string cc, string bcc, string subject, string body)
        {
            From = from;
            To = to;
            CC = cc;
            BCC = bcc;
            Subject = subject;
            Body = body;
        }

        /// <summary>
        /// Method which creates and opens an EML file.
        /// </summary>
        /// <returns>The result of the operation</returns>
        public abstract bool SaveEML();

        /// <summary>
        /// Method which prepares the temp folder where the EML file will be stored.
        /// </summary>
        /// <remarks>
        /// Recursively deletes previously created temporary files and folders to avoid possible conflicts when
        /// creating the new EML file.
        /// </remarks>
        internal void PrepareTempFolder()
        {
            if (Directory.Exists(TempPath))
            {
                Directory.Delete(TempPath, true);
            }

            Directory.CreateDirectory(TempPath);
        }

        /// <summary>
        /// Method which initializes passed MailMessage object with the information specified in MailUtilitiesBase
        /// properties.
        /// </summary>
        /// <param name="message">The MailMessage to initialize</param>
        internal void PrepareMessage(ref MailMessage message)
        {
            if (Unsent)
            {
                // Allows to open the EML in a compose window
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
                foreach (string ccAddress in Helpers.SplitAddressesByDelimiter(CC, Delimiter))
                {
                    message.CC.Add(ccAddress);
                }
            }

            if (!string.IsNullOrEmpty(BCC))
            {
                foreach (string bccAddress in Helpers.SplitAddressesByDelimiter(BCC, Delimiter))
                {
                    message.Bcc.Add(bccAddress);
                }
            }

            message.Subject = Subject;
            message.SubjectEncoding = Encoding.UTF8;

            message.Body = Body;
            message.BodyEncoding = Encoding.UTF8;

            message.IsBodyHtml = true;
        }

        /// <summary>
        /// Method which initializes passed MailMessage Attachments property with the attachments added to the
        /// Attachments collection.
        /// </summary>
        /// <param name="message">The MailMessage to initialize</param>
        internal void PrepareAttachments(ref MailMessage message)
        {
            foreach(KeyValuePair<string,string> attachment in Attachments)
            {
                if (!string.IsNullOrEmpty(attachment.Key))
                {
                    if (File.Exists(attachment.Key))
                    {
                        // Avoid System.IO.File.Delete exception by making a copy of the original file
                        string copiedAttachmentPath = Path.Combine(TempPath, string.IsNullOrEmpty(attachment.Value) ? Path.GetFileName(attachment.Key) : attachment.Value);
                        File.Copy(attachment.Key, copiedAttachmentPath, true);

                        // Add the attachment to the EML (encoded as base64 string)
                        message.Attachments.Add(new Attachment(copiedAttachmentPath));
                    }
                }
            }
        }

        /// <summary>
        /// Method which allows to add a file to the Attachments collection.
        /// </summary>
        /// <param name="filePath">The path of the file to add</param>
        /// <param name="fileName">The text to use as file name</param>
        public void AddAttachment(string filePath, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = Path.GetFileName(filePath);
            }

            Attachments.Add(new KeyValuePair<string,string>(filePath, fileName));
        }

        /// <summary>
        /// Method which finalizes the created EML file.
        /// </summary>
        /// <remarks>
        /// Creates a copy of the created EML file with a known file name (SMTPClient.Send uses random GUIDs as file names
        /// when saving the EML to the disk). If UseOutlookAccount is true then X-Sender and From headers will be removed
        /// so that the mail account defined in Outlook is used as From address.
        /// </remarks>
        /// <returns>The path of finalized EML file</returns>
        internal string FinalizeEML()
        {
            string finalizedEMLFilePath = Path.Combine(TempPath, "MailToSend.eml");

            using (StreamReader reader = new StreamReader(new DirectoryInfo(TempPath).GetFiles().Where(file => file.Extension.Equals(".eml")).ToArray()[0].FullName))
            using (StreamWriter writer = new StreamWriter(finalizedEMLFilePath))
            {
                string line;
                if (!UseOutlookAccount)
                {
                    if ((line = reader.ReadToEnd()) != null)
                    {
                        writer.Write(line);
                    }
                }
                else
                {
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.StartsWith("X-Sender:") || line.StartsWith("From:"))
                        {
                            continue;
                        }

                        writer.WriteLine(line);
                    }
                }

                return finalizedEMLFilePath;
            }
        }
    }
}
