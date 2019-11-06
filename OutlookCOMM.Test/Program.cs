using OutlookCOMM.NET;

namespace OutlookCOMM.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            MailUtilities mailUtilities = new MailUtilities("example@example.com", "example@example.com", "example@example.com", "example@example.com", "Subject", "Body");
            mailUtilities.AddAttachment(@"C:\temp\testFileA.txt", "");
            mailUtilities.AddAttachment(@"C:\temp\testFileB.txt", "CustomFileName.test");
            mailUtilities.SaveEML();
        }
    }
}
