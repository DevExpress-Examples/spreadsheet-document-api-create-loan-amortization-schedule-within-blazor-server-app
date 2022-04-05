using BlazorApp_SpreadsheetDocumentAPI;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;

namespace BlazorApp_SpreadsheetAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase
    {
        readonly DocumentService documentService;

        public ExportController(DocumentService documentService)
        {
            this.documentService = documentService;
        }

        [HttpGet]
        [Route("[action]")]
        public async Task<IActionResult> Xlsx([FromQuery] double loanAmount, 
            [FromQuery] int periodInYears, [FromQuery] double interestRate, 
            [FromQuery] DateTime loanStartDate)
        {
            var document = await documentService.GetXlsxDocumentAsync(loanAmount, periodInYears, 
                interestRate, loanStartDate);
            return File(document, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                "output.xlsx");
        }

        [HttpGet]
        [Route("[action]")]
        public async Task<IActionResult> Pdf([FromQuery] double loanAmount, 
            [FromQuery] int periodInYears, [FromQuery] double interestRate, 
            [FromQuery] DateTime loanStartDate)
        {
            var document = await documentService.GetPdfDocumentAsync(loanAmount, periodInYears, 
                interestRate, loanStartDate);
            return File(document, "application/pdf", "output.pdf");
        }
    }
}
