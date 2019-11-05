using OutlookCOMM.NET;

namespace OutlookCOMM.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            MailUtilities mailUtilities = new MailUtilities("example@example.com", "example@example.com", "example@example.com", "example@example.com", "Subject", "Body", @"C:\temp\test.txt");
            mailUtilities.SaveEML();
        }
    }
}
