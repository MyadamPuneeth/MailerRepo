namespace Mailer.Models
{
    public class EmailViewModel
    {
        public string? UserEmail { get; set; }
        public IFormFile? ExcelSheet { get; set; }
        public string? FileLink { get; set; }
    }
}
