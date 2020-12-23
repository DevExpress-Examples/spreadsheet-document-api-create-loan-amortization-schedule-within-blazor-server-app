<!-- default file list -->
*Files to look at*:

* [DocumentGenerator.cs](./CS/BlazorApp_SpreadsheetAPI/Code/DocumentGenerator.cs)

* [DocumentService.cs](./CS/BlazorApp_SpreadsheetAPI/Code/DocumentService.cs)

* [ExportController.cs](./CS/BlazorApp_SpreadsheetAPI/Controllers/ExportController.cs)

* [Index.razor](./CS/BlazorApp_SpreadsheetAPI/Pages/Index.razor)
<!-- default file list end -->

# Spreadsheet Document API - How to Create a Loan Amortization Schedule within Your .NET 5 Blazor Server App

This example demonstrates how to create a Blazor Server application that targets .NET 5 and leverages the capabilities of the [DevExpress Spreadsheet Document API](https://docs.devexpress.com/OfficeFileAPI/14912/spreadsheet-document-api) to build a loan amortization schedule.

The application allows users to enter loan information (loan amount, repayment period in years, annual interest rate, and start date). Once data is entered, the Spreadsheet immediately recalculates loan payments and updates data on the application page. Users can export the result to XLSX or PDF as needed.

![Spreadsheet - Final App](./images/spreadsheet-api-blazor-final-app.png)

To run this application, you need to install or restore the following NuGet packages:

* [DevExpress.Document.Processor](https://nuget.devexpress.com/packages/DevExpress.Document.Processor/) - Contains the DevExpress Office File API components.

* [DevExpress.Blazor](https://nuget.devexpress.com/packages/DevExpress.Blazor/) - Contains all DevExpress Blazor UI components.

