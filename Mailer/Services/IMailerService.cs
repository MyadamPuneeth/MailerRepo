namespace Mailer.Services
{
    public interface IMailerService
    {
        public bool SendEmailText(string email);
        public bool SendEmailFile(IFormFile uploadedFile);
    }
}
