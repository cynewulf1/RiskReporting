using Microsoft.AspNetCore.Mvc;

namespace RiskReporting.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index() => View();
    }
}