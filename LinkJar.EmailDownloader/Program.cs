using LinkJar.GraphService;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace LinkJar.EmailDownloader
{
    public class Program
    {
        private static async Task Main(string[] args)
        {
            //Sample Usage

            string inbox = "bablypm@0sx5k.onmicrosoft.com";
            string fromDate = "04/22/2022";
            DateTime dateTime = Convert.ToDateTime(DateTime.ParseExact(fromDate, "MM/dd/yyyy", CultureInfo.InvariantCulture));

            Console.WriteLine($"Getting emails for {inbox}...");

            EmailService service = new EmailService();
            List<EmailMessage> emails = await service.GetEmails(inbox, dateTime);
           //List<EmailMessage> emails = await service.GetMyEmails(dateTime);

            foreach (var email in emails)
            {
                Console.WriteLine($"Subject: {email.Subject} : ReceivedDate: {email.ReceivedDate} ");
                if (email.Attachments.Count > 0)
                {
                    Console.WriteLine("Downloading attchments..");
                    foreach (var attachment in email.Attachments)
                    {
                        await service.GetAttachment(attachment);

                        Console.WriteLine($"Downloaded {attachment.Name}");
                    }
                }
            }
            Console.WriteLine("Completed...");
            Console.ReadLine();
        } 
    }
}