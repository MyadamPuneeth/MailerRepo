using MimeKit;
using MailKit.Net.Smtp;
using OfficeOpenXml;


namespace Mailer.Services
{
    public class MailerService : IMailerService
    {
        public bool SendEmailFile(IFormFile uploadedFile)
        {

            if (uploadedFile == null || uploadedFile.Length == 0)
            {
                return false;
            }

            var tempFilePath = Path.GetTempFileName();
            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                uploadedFile.CopyTo(stream);
            }

            string senderEmail = "puneethmyadam123@gmail.com";
            string senderPassword = "rssn osld uley imwl"; 
            string smtpServer = "smtp.gmail.com";
            int smtpPort = 587;

            using (var package = new ExcelPackage(new FileInfo(tempFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; 
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Map headers to column indices
                Dictionary<string, int> columnMapping = new Dictionary<string, int>();
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Text.Trim(); 
                    if (!string.IsNullOrEmpty(header))
                    {
                        columnMapping[header.ToLower()] = col; 
                    }
                }

                // Ensure required columns exist
                if (!columnMapping.ContainsKey("email") ||
                    !columnMapping.ContainsKey("first name") ||
                    !columnMapping.ContainsKey("manager email"))
                {
                    Console.WriteLine("Required columns (Email, First Name, Manager Email) are missing in the Excel file.");
                    return false;
                }

                // Process each row starting from row 2 to skip header
                for (int row = 2; row <= rowCount; row++)
                {
                    string recipientEmail = worksheet.Cells[row, columnMapping["email"]].Text;
                    string recipientName = worksheet.Cells[row, columnMapping["first name"]].Text;
                    string managerEmail = worksheet.Cells[row, columnMapping["manager email"]].Text;

                    // Generate Email Body
                    string subject = "Personalized Greetings";
                    string body = $@"
                    Hi {recipientName},

                    This is a personalized email for you

                    Best Regards,
                    [Your Name]";

                    // Send Email (using MailKit)
                    SendEmail(smtpServer, smtpPort, senderEmail, senderPassword, recipientEmail, managerEmail, subject, body);
                    Console.WriteLine($"Email sent to {recipientName} at {recipientEmail}, CC to manager {managerEmail}");
                }
            }

            return true;
        }

        public bool SendEmailText(string email)
        {
            string senderEmail = "puneethmyadam123@gmail.com";
            string recipientEmail = email;
            string smtpServer = "smtp.gmail.com";
            int smtpPort = 587;
            string smtpPassword = "rssn osld uley imwl";

            try
            {
                // Create a new emailer message
                var emailer = new MimeMessage();
                emailer.From.Add(new MailboxAddress("Puneeth", senderEmail));
                emailer.To.Add(new MailboxAddress("Recipient", recipientEmail));
                emailer.Subject = "Test Email Sending";
                emailer.Body = new TextPart(MimeKit.Text.TextFormat.Plain)
                {
                    Text = "Hello from the land of C#!"
                };

                // Send emailer
                using (var smtp = new SmtpClient())
                {
                    smtp.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                    smtp.Authenticate(senderEmail, smtpPassword);
                    smtp.Send(emailer);
                    smtp.Disconnect(true);

                    Console.WriteLine("Email sent successfully!");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending emailer: {ex.Message}");
                return false;
            }
        }

        static void SendEmail(string smtpServer, int smtpPort, string senderEmail, string senderPassword, string recipientEmail, string ccEmail, string subject, string body)
        {


            try
            {
                var email = new MimeMessage();
                email.From.Add(new MailboxAddress("Puneeth", senderEmail));
                email.To.Add(new MailboxAddress("recipient", recipientEmail));
                email.Subject = subject;
                email.Body = new TextPart(MimeKit.Text.TextFormat.Plain)
                {
                    Text = body
                };

                // Add CC email if provided
                if (!string.IsNullOrWhiteSpace(ccEmail))
                {
                    email.Cc.Add(new MailboxAddress("Manager", ccEmail));
                }

                // Send the email using MailKit SMTP client
                using (var smtp = new SmtpClient())
                {
                    smtp.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                    smtp.Authenticate(senderEmail, senderPassword);
                    smtp.Send(email);
                    smtp.Disconnect(true);
                    Console.WriteLine("Email sent successfully to " + recipientEmail);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }

    }
}
