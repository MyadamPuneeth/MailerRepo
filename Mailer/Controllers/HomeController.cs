using Mailer.Models;
using Mailer.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace Mailer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly MailerService Mailman;
        public HomeController(ILogger<HomeController> logger, MailerService mailman)
        {
            _logger = logger;
            Mailman = mailman;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult MailerHomePage()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> MailerHomePage(EmailViewModel model)
        {
            if (model.UserEmail != null)
            {
                if (Mailman.SendEmailText(model.UserEmail))
                {
                    return RedirectToAction("EmailSentPage");
                }
                return RedirectToAction("EmailNotSentPage");
            }

            if (model.ExcelSheet != null)
            {
                if (Mailman.SendEmailFile(model.ExcelSheet))
                {
                    return RedirectToAction("EmailSentPage");
                }
                return RedirectToAction("EmailNotSentPage");
            }
            else
            {
                return RedirectToAction("EmailNotSentPage");
            }
        }

        public IActionResult EmailSentPage()
        {
            return View();
        }

        public IActionResult EmailNotSentPage()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
