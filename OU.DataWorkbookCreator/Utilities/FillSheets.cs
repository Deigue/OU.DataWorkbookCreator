using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OU.DataWorkbookCreator.Models;

//Font Format ID's:
//0: Default Font
//1: Size 14 Style Verdana Bold Type Text
//2: Size 10 Style Verdana Bold Type Text
//3: Size 10 Style Verdana Bold Center Alignment Type Text
//5: Size 10 Style Verdana Center Alignment Type Text
//6: Size 10 Style Verdana Center Alignment Type Number
//23: Size 10 Style Verdana Bold Type Number

namespace OU.DataWorkbookCreator.Utilities
{
    class FillSheets: SpreadsheetHelper
    {
        //Inserts Total for current desk in Footings Sheet:
        private void FootingsDeskTotal(WorksheetPart worksheetPart, uint correction, uint startRowIndex, uint currentRowIndex)
        {
            uint endRowIndex = currentRowIndex - 1;
            string totalRowG = string.Format("SUM(G{0}:G{1}", startRowIndex, endRowIndex);
            string totalRowL = string.Format("SUM(L{0}:L{1}", startRowIndex, endRowIndex);
            string totalRowM = string.Format("SUM(M{0}:M{1}", startRowIndex, endRowIndex);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'G', 1, 23, totalRowG);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'L', 1, 23, totalRowL);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'M', 1, 23, totalRowM);
        }

        //Inserts Total for current desk in Detail Report Sheet:
        private void DetailReportDeskTotal(WorksheetPart worksheetPart, uint correction, uint startRowIndex, uint currentRowIndex)
        {
            uint endRowIndex = currentRowIndex - 1;
            string totalRowH = string.Format("SUM(H{0}:H{1}", startRowIndex, endRowIndex);
            string totalRowK = string.Format("SUM(K{0}:K{1}", startRowIndex, endRowIndex);
            string totalRowL = string.Format("SUM(L{0}:L{1}", startRowIndex, endRowIndex);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'H', 1, 23, totalRowH);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'K', 1, 23, totalRowK);
            UpdateCellFormula(worksheetPart, correction, currentRowIndex, 'L', 1, 23, totalRowL);
        }

        //Inserts New Desk and OrgUnit into worksheets:
        private void InsertNewDesk(WorksheetPart worksheetPart, uint correction, string currentDesk, string orgUnit, uint rowIndex)
        {
            UpdateCell(worksheetPart, correction, rowIndex, 'A', 2, 2, currentDesk);
            UpdateCell(worksheetPart, correction, rowIndex, 'B', 2, 2, orgUnit);
        }

        public void InsertIntoSummary(WorksheetPart worksheetPart, List<SummaryModel> summaryList)
        {
            int existingRows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Count();
            var sortedSummaryList = summaryList.OrderBy(x => x.OrgUnit).ThenBy(x => x.Amount).ToList();
            uint rowIndex = 1;
            uint correction = Convert.ToUInt32(rowIndex - 1 - existingRows); //Correction Factor

            foreach (SummaryModel summaryModel in sortedSummaryList)
            {
                var orgUnit = Convert.ToString(summaryModel.OrgUnit);
                var amount = Convert.ToString(summaryModel.Amount);
                UpdateCell(worksheetPart, correction, rowIndex, 'A', 1, 5, orgUnit);
                UpdateCell(worksheetPart, correction, rowIndex, 'B', 1, 6, amount);
                rowIndex++;
            }
        }

        public void InsertIntoFootings(WorksheetPart worksheetPart, List<FootingsModel> footingsList)
        {
            int existingRows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Count();
            var sortedFootingsList = footingsList.OrderBy(x => x.Desk).ThenBy(x => x.OrgUnit).ThenBy(x=> x.CCY).ToList();
            bool listEntry = true;
            int deskCounter =1;
            string previousDesk = null;
            uint rowIndex = 1;
            uint correction = Convert.ToUInt32(rowIndex - 1 - existingRows); //Correction Factor
            uint deskStartRowIndex = rowIndex;

            foreach (FootingsModel footingsModel in sortedFootingsList)
            {
                var currentDesk = footingsModel.Desk;
                var orgUnit = footingsModel.OrgUnit;
                var CCY = footingsModel.CCY;
                var CGLM = Convert.ToString(footingsModel.CGLM);
                var CAdjustment = Convert.ToString(footingsModel.CAdjustment);
                var T1GLM = Convert.ToString(footingsModel.T1GLM);
                var T1Adjustment = Convert.ToString(footingsModel.T1Adjustment);

                if ((currentDesk != previousDesk) && (!listEntry))
                {
                    FootingsDeskTotal(worksheetPart, correction, deskStartRowIndex, rowIndex);
                    rowIndex+=2;
                    InsertNewDesk(worksheetPart, correction, currentDesk, orgUnit, rowIndex);
                    deskStartRowIndex = rowIndex;
                }
                else if (listEntry)                
                    InsertNewDesk(worksheetPart, correction, currentDesk, orgUnit, rowIndex);
                

                var ColumnF = String.Format("SUM(D{0}:E{1})", rowIndex, rowIndex);
                var ColumnG = String.Format("F{0}/VLOOKUP(C{1},'FX Rates' !A$3:B$200,2)", rowIndex, rowIndex);
                var ColumnK = String.Format("SUM(I{0}:J{1})", rowIndex, rowIndex);
                var ColumnL = String.Format("K{0}/VLOOKUP(C{1},'FX Rates' !A$3:B$200,2)",rowIndex,rowIndex);
                var ColumnM = String.Format("K{0}/VLOOKUP(C{1},'FX Rates' !D$3:E$200,2)", rowIndex, rowIndex);
                var ColumnO = String.Format("L{0}-M{1}", rowIndex, rowIndex);
                var ColumnP = String.Format("G{0}-L{1}", rowIndex, rowIndex);
                var ColumnQ = String.Format("G{0}-M{1}", rowIndex, rowIndex);
                var ColumnR = String.Format("Q{0}/G{1}", rowIndex, rowIndex);
                
                UpdateCell(worksheetPart, correction, rowIndex, 'C', 2, 2, CCY);
                UpdateCell(worksheetPart, correction, rowIndex, 'D', 1, 6, CGLM);
                UpdateCell(worksheetPart, correction, rowIndex, 'E', 1, 6, CAdjustment);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'F', 1, 6, ColumnF);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'G', 1, 6, ColumnG);
                UpdateCell(worksheetPart, correction, rowIndex, 'I', 1, 6, T1GLM);
                UpdateCell(worksheetPart, correction, rowIndex, 'J', 1, 6, T1Adjustment);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'K', 1, 6, ColumnK);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'L', 1, 6, ColumnL);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'M', 1, 6, ColumnM);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'O', 1, 6, ColumnO);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'P', 1, 6, ColumnP);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'Q', 1, 6, ColumnQ);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'R', 1, 6, ColumnR);

                rowIndex++;

                //Final Row
                if (deskCounter == sortedFootingsList.Count())
                {
                    FootingsDeskTotal(worksheetPart, correction, deskStartRowIndex, rowIndex);
                    rowIndex++;
                }
            }
        }

        public void InsertIntoDetailReport(WorksheetPart worksheetPart, List<DetailReportModel> detailReportList)
        {
            int existingRows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Count();
            var sortedDetailReportList = detailReportList.OrderBy(x => x.Desk).ThenBy(x => x.OrgUnit).ThenBy(x => x.CCY).ThenBy(x => x.Account).ThenBy(x => x.Product).ThenBy(x => x.CustGroup).ToList();
            bool listEntry = true;
            int deskCounter = 1;
            string previousDesk = null;
            uint rowIndex = 1;
            uint correction = Convert.ToUInt32(rowIndex - 1 - existingRows); //Correction Factor
            uint deskStartRowIndex = rowIndex;

            foreach (DetailReportModel detailReportModel in sortedDetailReportList)
            {
                var currentDesk = detailReportModel.Desk;
                var orgUnit = detailReportModel.OrgUnit;
                var CCY = detailReportModel.CCY;
                var account = detailReportModel.Account;
                var product = detailReportModel.Product;
                var custGroup = detailReportModel.CustGroup;
                var CGLM = Convert.ToString(detailReportModel.CGLM);
                var T1GLM = Convert.ToString(detailReportModel.T1GLM);

                if ((currentDesk != previousDesk) && (!listEntry))
                {
                    DetailReportDeskTotal(worksheetPart, correction, deskStartRowIndex, rowIndex);
                    rowIndex += 2;

                    InsertNewDesk(worksheetPart, correction, currentDesk, orgUnit, rowIndex);
                    deskStartRowIndex = rowIndex;
                }
                else if (listEntry)
                    InsertNewDesk(worksheetPart, correction, currentDesk, orgUnit, rowIndex);

                var ColumnH = String.Format("G{0}/VLOOKUP(C{1},'FX Rates' !A$3:B$200,2", rowIndex, rowIndex);
                var ColumnK = String.Format("J{0}/VLOOKUP(C{1},'FX Rates' !A$3:B$200,2", rowIndex, rowIndex);
                var ColumnL = String.Format("J{0}/VLOOKUP(C{1},'FX Rates' !D$3:E$200,2", rowIndex, rowIndex);
                var ColumnN = String.Format("K{0}-L{1}", rowIndex, rowIndex);
                var ColumnO = String.Format("H{0}-K{1}", rowIndex, rowIndex);
                var ColumnP = String.Format("H{0}-L{1}", rowIndex, rowIndex);
                var ColumnQ = String.Format("P{0}/H{1}", rowIndex, rowIndex);

                UpdateCell(worksheetPart, correction, rowIndex, 'C', 2, 2, CCY);
                UpdateCell(worksheetPart, correction, rowIndex, 'D', 2, 2, account);
                UpdateCell(worksheetPart, correction, rowIndex, 'E', 2, 2, product);
                UpdateCell(worksheetPart, correction, rowIndex, 'F', 2, 2, custGroup);
                UpdateCell(worksheetPart, correction, rowIndex, 'G', 1, 6, CGLM);
                UpdateCell(worksheetPart, correction, rowIndex, 'J', 1, 6, T1GLM);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'H', 1, 6, ColumnH);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'K', 1, 6, ColumnK);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'L', 1, 6, ColumnL);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'N', 1, 6, ColumnN);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'O', 1, 6, ColumnO);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'P', 1, 6, ColumnP);
                UpdateCellFormula(worksheetPart, correction, rowIndex, 'Q', 1, 6, ColumnQ);

                rowIndex++;

                //Final Row
                if (deskCounter == sortedDetailReportList.Count())
                {
                    DetailReportDeskTotal(worksheetPart, correction, deskStartRowIndex, rowIndex);
                    rowIndex++;
                }

                previousDesk = currentDesk;
                listEntry = false;
                deskCounter++;
            }
        }

        public void InsertIntoFXRates(WorksheetPart worksheetPart, List<FXRatesModel> fxratesTodayList, List<FXRatesModel> fxratesYesterdayList)
        {
            int existingRows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Count();
            var sortedTodayList = fxratesTodayList.OrderBy(x => x.CCY).ToList();
            var sortedYesterdayList = fxratesYesterdayList.OrderBy(x => x.CCY).ToList();
            uint rowIndex = 1;
            uint correction = Convert.ToUInt32(rowIndex - 1 - existingRows);

            foreach (FXRatesModel fxratesModel in sortedTodayList)
            {
                var CCY = fxratesModel.CCY;
                var rate = fxratesModel.Rate.ToString();

                UpdateCell(worksheetPart, correction, rowIndex, 'A', 2, 0, CCY);
                UpdateCell(worksheetPart, correction, rowIndex, 'B', 1, 6, rate);
                rowIndex++;
            }

            rowIndex = 3; //Reset Index

            foreach (FXRatesModel fxratesModel in sortedYesterdayList)
            {
                var CCY = fxratesModel.CCY;
                var rate = fxratesModel.Rate.ToString();

                UpdateCell(worksheetPart, correction, rowIndex, 'D', 2, 0, CCY);
                UpdateCell(worksheetPart, correction, rowIndex, 'E', 1, 6, rate);
                rowIndex++;
            }
        }
    }
}
