<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/323888358/20.2.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T960167)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
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

<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-create-loan-amortization-schedule-within-blazor-server-app&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-create-loan-amortization-schedule-within-blazor-server-app&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
