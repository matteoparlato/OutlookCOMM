using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OutlookCOMM.Core
{
    /// <summary>
    /// IMailUtilities interface
    /// </summary>
    public interface IMailUtilities
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
        [ComVisible(false)]
        public List<KeyValuePair<string,string>> Attachments { get; }

        /// <summary>
        /// The string which contains the delimiter to use when splitting the addresses.
        /// </summary>
        public char Delimiter { get; set; } 

        /// <summary>
        /// The boolean which determine whether is necessary to add the X-Unsent header to the mail.
        /// </summary>
        public bool Unsent { get; set; }

        /// <summary>
        /// The boolean which determine whether is necessary to use the From address configured in Outlook. 
        /// </summary>
        public bool UseOutlookAccount { get; set; }

        /// <summary>
        /// Method which creates and opens an EML file.
        /// </summary>
        /// <returns>The result of the operation</returns>
        public bool SaveEML();
    }
}
