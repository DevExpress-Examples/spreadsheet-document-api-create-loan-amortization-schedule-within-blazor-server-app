using System;
using System.IO;
using System.Threading.Tasks;
using DevExpress.Spreadsheet;

namespace BlazorApp_SpreadsheetAPI.Code
{
    public class DocumentService
    {
        public async Task<byte[]> GetXlsxDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan)
        {
            // Generate a workbook 
            // that contains an amortization schedule.
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            // Save the document as XLSX.
            return await workbook.SaveDocumentAsync(DocumentFormat.Xlsx);
        }

        public async Task<byte[]> GetPdfDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan)
        {
            // Generate a workbook 
            // that contains an amortization schedule.
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            // Export the document to HTML.
            using var ms = new MemoryStream();
            await workbook.ExportToPdfAsync(ms);
            return ms.ToArray();
        }

        public async Task<byte[]> GetHtmlDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime startDateOfLoan)
        {
            // Generate a workbook 
            // that contains an amortization schedule.
            using var workbook = await GenerateDocumentAsync(loanAmount, periodInYears, interestRate, startDateOfLoan);
            // Export the document to HTML.
            using var ms = new MemoryStream();
            await workbook.ExportToHtmlAsync(ms, workbook.Worksheets[0]);
            return ms.ToArray();
        }
        async Task<Workbook> GenerateDocumentAsync(double loanAmount, int periodInYears, double interestRate, DateTime loanStartDate)
        {
            var workbook = new Workbook();
            // Load document template.
            await workbook.LoadDocumentAsync("Data/LoanAmortizationScheduleTemplate.xltx");
            // Generate a loan amortization schedule
            // based on the template and loan information.
            new LoanAmortizationScheduleGenerator(workbook)
                .GenerateDocument(loanAmount, periodInYears, interestRate, loanStartDate);
            return workbook;
        }
    }
}
