using Microsoft.AspNetCore.Mvc;

namespace Sankhya.Rupaantran.ExcelWebAddIn.WebAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class RupaantranController : ControllerBase
    {
        [HttpGet("convert")]
        public IActionResult Convert(
            [FromQuery] decimal amount,        
            [FromQuery] string type = "Lakhs", // Lakhs, Crores, Millions, Billions
            [FromQuery] string format = "N2")  // N0, N1, N2
        {
            try
            {
                string result = type.ToLower() switch
                {
                    "lakhs" => SankhyaRupaantran.ToLakhsString(amount, format),
                    "crores" => SankhyaRupaantran.ToCroresString(amount, format),
                    "millions" => SankhyaRupaantran.ToMillionsString(amount, format),
                    "billions" => SankhyaRupaantran.ToBillionsString(amount, format),
                    _ => SankhyaRupaantran.ToLakhsString(amount, format)
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }
    }
}
