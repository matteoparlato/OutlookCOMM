using OutlookCOMM.NET;

namespace OutlookCOMM.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            MailUtilities.SaveEML("example@example.com", "example@example.com", "example@example.com", "example@example.com", "Subject", "Body", @"C:\temp\test.txt", true, false);
        }
    }
}
