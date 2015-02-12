using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OU.DataWorkbookCreator
{
    public class Program
    {
        private string CreateTemplateCopy(string[] output)
        {
            string outputDirectory = output[0];
            string outputFileName = output[1];
            string inputPath = @"../../Template/Template.xlsx";
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);
            string outputFullPath = Path.Combine(outputDirectory, outputFileName);
            if(File.Exists(inputPath))
                System.IO.File.Copy(inputPath, outputFullPath, true);
            return outputFullPath;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("_______________Workbook Creator_______________\n");

            //Class members
            var workbookCreator = new Program();
            Spread

        }
    }
}
