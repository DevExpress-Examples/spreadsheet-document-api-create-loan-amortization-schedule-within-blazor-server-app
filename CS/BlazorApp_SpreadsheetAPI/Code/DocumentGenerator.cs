using System;
using System.Drawing;
using DevExpress.Spreadsheet;

namespace BlazorApp_SpreadsheetAPI.Code
{
    public class LoanAmortizationScheduleGenerator
    {
        IWorkbook workbook;

        public LoanAmortizationScheduleGenerator(IWorkbook workbook)
        {
            this.workbook = workbook;
        }
        Worksheet Sheet { get { return workbook.Worksheets[0]; } }
        DateTime LoanStartDate { 
            get { return Sheet["E8"].Value.DateTimeValue; } 
            set { Sheet["E8"].Value = value; } 
        }
        double LoanAmount { 
            get { return Sheet["E4"].Value.NumericValue; }
            set { Sheet["E4"].Value = value; }
        }
        double InterestRate { 
            get { return Sheet["E5"].Value.NumericValue; } 
            set { Sheet["E5"].Value = value; } 
        }
        int PeriodInYears { 
            get { return (int)Sheet["E6"].Value.NumericValue; } 
            set { Sheet["E6"].Value = value; } 
        }
        int ActualNumberOfPayments { 
            get { return (int)Math.Round(Sheet["I6"].Value.NumericValue); } 
        }
        int ScheduledNumberOfPayments { 
            get { return (int)Math.Round(Sheet["I5"].Value.NumericValue); } 
        }
        string ActualLastRow { 
            get { return (11 + ActualNumberOfPayments).ToString(); } 
        }

        public void GenerateDocument(double loanAmount, int periodInYears, 
            double interestRate, DateTime loanStartDate)
        {
            workbook.BeginUpdate();
            try
            {
                ClearData();
                LoanAmount = loanAmount;
                InterestRate = interestRate;
                PeriodInYears = periodInYears;
                LoanStartDate = loanStartDate;
                GenerateLoanAmortizationTable();
                ApplyFormatting();
                SpecifyPrintOptions();
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        void ClearData()
        {
            // Clear all data
            // except for rows 1 through 11. 
            var range = Sheet.GetDataRange().Exclude(Sheet["1:11"]);
            if (range != null)
                range.Clear();
            Sheet["I4"].ClearContents();
            Sheet["I6:I8"].ClearContents();
            workbook.DefinedNames.Clear();
        }

        void GenerateLoanAmortizationTable()
        {
            CreateDefinedNames();
            // Calculate loan payment amount.
            Sheet["I4"].FormulaInvariant = 
                "=PMT(Interest_Rate_Per_Month,Scheduled_Number_Payments,-Loan_Amount)";
            // Calculate the scheduled number of payments.
            Sheet["I5"].FormulaInvariant = "=Loan_Years*Number_of_Payments_Per_Year";
            // Calculate the actual number of payments.
            Sheet["I6"].FormulaInvariant = "=ROUNDUP(Actual_Number_of_Payments,0)";
            // Recalculate all formulas in the document.
            workbook.Calculate();
            // Calculate the total amount of early payments.
            Sheet["I7"].FormulaInvariant = "=SUM(F12:F" + ActualLastRow + ")";
            // Calculate the total interest paid.
            Sheet["I8"].FormulaInvariant = "=SUM($I$12:$I$" + ActualLastRow + ")";

            if (ScheduledNumberOfPayments == 0)
                return;

            // Populate the "Payment Number" column.
            for (int i = 0; i < ActualNumberOfPayments; i++)
                Sheet["B" + (i + 12).ToString()].Value = i + 1;

            // Populate the "Payment Date" column.
            Sheet["C12:C" + ActualLastRow].FormulaInvariant = 
                "=DATE(YEAR(Loan_Start),MONTH(Loan_Start)+(B12)*12/Number_of_Payments_Per_Year,DAY(Loan_Start))";

            // Calculate the beginning balance for each period.
            Sheet["D12"].Formula = "=Loan_Amount";
            if (ScheduledNumberOfPayments > 1)
                Sheet["D13:D" + ActualLastRow].Formula = "=J12";

            // Populate the "Scheduled Payment" column.
            Sheet["E12:E" + ActualLastRow].FormulaInvariant = 
                "=IF(D12>0,IF(Scheduled_Payment<D12, Scheduled_Payment, D12),0)";
            // Populate the "Extra Payment" column.
            Sheet["F12:F" + ActualLastRow].FormulaInvariant = 
                "=IF(Extra_Payments<>0, IF(Scheduled_Payment<D12, G12-E12, 0), 0)";
            // Calculate total payment amount.
            Sheet["G12:G" + ActualLastRow].FormulaInvariant = "=H12+I12";
            // Calculate the principal part of the payment. 
            Sheet["H12:H" + ActualLastRow].FormulaInvariant = 
                "=IF(J12>0,PPMT(Interest_Rate_Per_Month,B12,Actual_Number_of_Payments,-Loan_Amount),D12)";
            // Calculate interest payments for each period.
            Sheet["I12:I" + ActualLastRow].FormulaInvariant = 
                "=IF(D12>0,IPMT(Interest_Rate_Per_Month,B12,Actual_Number_of_Payments,-Loan_Amount),0)";
            // Calculate the remaining balance for each period.
            Sheet["J12:J" + ActualLastRow].FormulaInvariant = 
                "=IF(D12-PPMT(Interest_Rate_Per_Month,B12,Actual_Number_of_Payments,-Loan_Amount)>0," +
                "D12-PPMT(Interest_Rate_Per_Month,B12,Actual_Number_of_Payments,-Loan_Amount),0)";
            // Calculate the cumulative interest paid on the loan.
            Sheet["K12:K" + ActualLastRow].FormulaInvariant = "=SUM($I$12:$I12)";
            // Recalculate all formulas in the document. 
            workbook.Calculate();
        }

        void CreateDefinedNames()
        {
            string sheetName = "'" + Sheet.Name + "'";
            char separator = workbook.Options.Culture.TextInfo.ListSeparator[0];

            // Define names for cell ranges and functions
            // used in payment amount calculation.
            DefinedNameCollection definedNames = workbook.DefinedNames;
            definedNames.Add("Loan_Amount", sheetName + "!$E$4");
            definedNames.Add("Interest_Rate", sheetName + "!$E$5");
            definedNames.Add("Loan_Years", sheetName + "!$E$6");
            definedNames.Add("Number_of_Payments_Per_Year", sheetName + "!$E$7");
            definedNames.Add("Loan_Start", sheetName + "!$E$8");
            definedNames.Add("Extra_Payments", sheetName + "!$E$9");
            definedNames.Add("Scheduled_Payment", sheetName + "!$I$4");
            definedNames.Add("Scheduled_Number_Payments", sheetName + "!$I$5");
            definedNames.Add("Interest_Rate_Per_Month", 
                "=Interest_Rate/Number_of_Payments_Per_Year");
            definedNames.Add("Actual_Number_of_Payments", 
                "=NPER(Interest_Rate_Per_Month" + separator + " " + 
                sheetName + "!$I$4+Extra_Payments" + separator + " -Loan_Amount)");
        }

        void ApplyFormatting()
        {
            // Format the amortization table.
            CellRange range;
            // Change the color of even rows in the table.
            for (int i = 1; i < ActualNumberOfPayments; i += 2)
            {
                range = Sheet.Range.FromLTRB(1, 11 + i, 10, 11 + i);
                range.Fill.BackgroundColor = Color.FromArgb(217, 217, 217);
            }

            range = Sheet["B11:K" + ActualLastRow];
            Formatting formatting = range.BeginUpdateFormatting();
            try
            {
                // Display vertical inside borders within the table.
                formatting.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin;
                formatting.Borders.InsideVerticalBorders.Color = Color.White;
                // Center text vertically within table cells.
                formatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            }
            finally
            {
                range.EndUpdateFormatting(formatting);
            }

            Sheet["B12:C" + ActualLastRow].Alignment.Horizontal = 
                SpreadsheetHorizontalAlignment.Right;
            // Apply the number format to table columns.
            Sheet["C11:C" + ActualLastRow].NumberFormat = "m/d/yyyy";
            Sheet["D11:K" + ActualLastRow].NumberFormat = 
                "_(\\$* #,##0.00_);_(\\$ (#,##0.00);_(\\$* \" - \"??_);_(@_)";
        }

        void SpecifyPrintOptions()
        {
            // Specify print settings
            // for the worksheet.
            Sheet.SetPrintRange(Sheet.GetDataRange());
            Sheet.PrintOptions.FitToPage = true;
            Sheet.PrintOptions.FitToWidth = 1;
            Sheet.PrintOptions.FitToHeight = 0;
        }
    }
}
