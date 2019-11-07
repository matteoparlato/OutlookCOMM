using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OutlookCOMM.Core
{
    public interface IMailUtilities
    {
        public string From { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        [ComVisible(false)]
        public List<KeyValuePair<string,string>> Attachments { get; }
        public char Delimiter { get; set; } 
        public bool Unsent { get; set; }
        public bool UseOutlookAccount { get; set; }

        public bool SaveEML();
    }
}
