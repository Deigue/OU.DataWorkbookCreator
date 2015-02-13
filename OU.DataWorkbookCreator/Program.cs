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
           
            string fileName = workbookCreator.CreateTemplateCopy(args);
            

            Console.ReadKey();
        }
    }
}
