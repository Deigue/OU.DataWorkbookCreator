using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OU.DataWorkbookCreator.Utilities;
using OU.DataWorkbookCreator.Repository;


namespace OU.DataWorkbookCreator
{
    class Program: CreateSpreadsheetWorkbook
    {
        private string CreateTemplateCopy(string[] output)
        {
            string outputDirectory = output[0];
            string outputFileName = output[1];
            string inputFullPath=null;

            if(output.Length==3)
                inputFullPath = output[2];

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);
            string outputFullPath = Path.Combine(outputDirectory, outputFileName);
            if (File.Exists(inputFullPath))
                System.IO.File.Copy(inputFullPath, outputFullPath, true);
            else
                WorkbookCreator(outputFullPath);
                
            return outputFullPath;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("_______________Workbook Creator_______________\n");

            //Class members
            var workbookCreator = new Program();
            var spreadsheetHelper = new SpreadsheetHelper();
            var testRepository = new TestRepository();
            var sheetFiller = new FillSheets();

            string fileName = workbookCreator.CreateTemplateCopy(args);

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fileName, true))
            {
                var workbookPart = spreadsheet.WorkbookPart;

                //Sheet IDs:
                var summary = spreadsheetHelper.GetWorkSheetPart(workbookPart, "Summary");
                var footings = spreadsheetHelper.GetWorkSheetPart(workbookPart, "Footings");
                var detailReport = spreadsheetHelper.GetWorkSheetPart(workbookPart, "Detail Report");
                var fxrates = spreadsheetHelper.GetWorkSheetPart(workbookPart, "FX Rates");

                //Populate workbook with Repo Data:
                sheetFiller.InsertIntoSummary(summary, testRepository.SummaryList());
                sheetFiller.InsertIntoFootings(footings, testRepository.FootingsList());
                sheetFiller.InsertIntoDetailReport(detailReport, testRepository.DetailReportList());
                sheetFiller.InsertIntoFXRates(fxrates, testRepository.FXRatesListToday(), testRepository.FXRatesListYesterday());

                //Force Calculations before save:
                spreadsheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadsheet.WorkbookPart.Workbook.CalculationProperties.CalculationOnSave = true;
                spreadsheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                spreadsheet.WorkbookPart.Workbook.Save();
                
                Console.WriteLine(String.Format("Workbook has been created in {0}{1}", args[0], args[1]));
                Console.Write("Press Any Key to Continue... ");
                Console.ReadKey();
            }                  
        }
    }
}
