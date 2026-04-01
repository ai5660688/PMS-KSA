using Microsoft.AspNetCore.Mvc;

namespace PMS.Controllers
{
    public class ErrorController : Controller
    {
        [Route("Home/Error")]
        public IActionResult Error()
        {
            return View();
        }
    }
}