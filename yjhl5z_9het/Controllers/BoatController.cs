using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using yjhl5z_9het.Models;

namespace yjhl5z_9het.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class BoatController : ControllerBase
    {
        [HttpGet]
        [Route("questions/all")]
        public IActionResult MindegyHogyHivjak()
        {
            HajosContext context = new HajosContext();
            var kérdések = from x in context.Questions select x.Question1;

            return Ok(kérdések);
        }
    }
}
