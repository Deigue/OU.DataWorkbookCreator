using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OU.DataWorkbookCreator.Models;
using OU.DataWorkbookCreator.Mappings;

namespace OU.DataWorkbookCreator.Repository
{
    class TestRepository : OUDataMapping
    {
        
        public static List<SummaryModel> SummaryList()
        {
            var summaryList = new List<SummaryModel>();
            
            
            summaryList.Add(MapToSummaryModel(12,65346));
            summaryList.Add(MapToSummaryModel(200, 646));
            summaryList.Add(MapToSummaryModel(2, 653246));
            summaryList.Add(MapToSummaryModel(31, 45));
            summaryList.Add(MapToSummaryModel(31, 46));
            summaryList.Add(MapToSummaryModel(31, 4526));
            summaryList.Add(MapToSummaryModel(33331, 46));
            summaryList.Add(MapToSummaryModel(1, 426));
            summaryList.Add(MapToSummaryModel(1, 4));
            summaryList.Add(MapToSummaryModel(3351, 32446));
            return summaryList;
        }

        static List<FootingsModel> FootingsList()
        {
            var footingsList = new List<FootingsModel>();
            footingsList.Add(MapToFootingsModel("TEMPORARY COMPANY", "22323", "JPY", 3415135.325, 0.01, 6464321.34, 0.02));
            footingsList.Add(MapToFootingsModel("TEMPORARY COMPANY", "22323", "CHF", 3457375.325, 0.03, 64321.34, 0.0));
            footingsList.Add(MapToFootingsModel("TEMPORARY COMPANY", "22323", "USD", 74352435.325, 0.03, 321.34, 0.0));
            footingsList.Add(MapToFootingsModel("PLC", "22323", "JPY", 3415135.325, 0.01, 435425.32, 0.07));
            footingsList.Add(MapToFootingsModel("ABC", "22323", "HKD", 3415135.325, 0.01, 1235613.5, 0.0));
            footingsList.Add(MapToFootingsModel("ABC", "22323", "CAD", 3415135.325, 0.01, 754221.34, 0.01));
            footingsList.Add(MapToFootingsModel("PLC", "22323", "JPY", 3415135.325, 0.01, 6464321.34, 0.0));
            footingsList.Add(MapToFootingsModel("NEWCOMPANY", "22323", "CAD", 3415135.325, 0.01, 646421.34, 0.01));
            return footingsList;
        }

        static List<DetailReportModel> DetailReportList()
        {
            var detailReportList = new List<DetailReportModel>();
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "EUR", "1234152", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "JPY", "1234152", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "EUR", "462456", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "JPY", "1234152", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "EUR", "1234152", "0100", "A0002", 12345.54, 9231245.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "AUD", "1234152", "0001", "A0000", 2412543.2, 31245879.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "EUR", "5324324", "0400", "A0000", 12345.54, 31245233.23));
            detailReportList.Add(MapToDetailReportModel("DESK A", "11111", "EUR", "1234152", "0001", "SS924", 999999.23, 31245.23));
            detailReportList.Add(MapToDetailReportModel("APPLE", "22233", "EUR", "1234152", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("SAMSUNG", "35463", "EUR", "1234152", "0001", "A0000", 12345.54, 12245.23));
            detailReportList.Add(MapToDetailReportModel("APPLE", "12342", "EUR", "1234152", "0911", "A0000", 12345.54, 5534.23));
            detailReportList.Add(MapToDetailReportModel("SAMSUNG", "35435", "EUR", "1234152", "0001", "A0000", 12345.54, 31245.23));
            detailReportList.Add(MapToDetailReportModel("SAMSUNG", "11111", "CAD", "1234152", "0982", "A0000", 12345.54, 31245.23));
            return detailReportList;
        }

        static List<FXRatesModel> FXRatesListToday()
        {
            var fxRatesListToday = new List<FXRatesModel>();
            fxRatesListToday.Add(MapToFXRatesModel("BRL",1.81D));
            fxRatesListToday.Add(MapToFXRatesModel("CZK",19.82D));
            fxRatesListToday.Add(MapToFXRatesModel("CAD",1.00D));
            fxRatesListToday.Add(MapToFXRatesModel("CHF",0.93D));
            fxRatesListToday.Add(MapToFXRatesModel("CLP",500.86D));
            fxRatesListToday.Add(MapToFXRatesModel("CNH",6.20D));
            fxRatesListToday.Add(MapToFXRatesModel("CNY",6.18D));
            fxRatesListToday.Add(MapToFXRatesModel("COP",1844.67D));
            fxRatesListToday.Add(MapToFXRatesModel("AUD",0.96D));
            return fxRatesListToday;         
        }

        static List<FXRatesModel> FXRatesListYesterday()
        {
            var fxRatesListYesterday = new List<FXRatesModel>();
            fxRatesListYesterday.Add(MapToFXRatesModel("CLP", 504.27D));
            fxRatesListYesterday.Add(MapToFXRatesModel("BRL", 1.81D));
            fxRatesListYesterday.Add(MapToFXRatesModel("AUD", 0.95D));
            fxRatesListYesterday.Add(MapToFXRatesModel("CHF", 0.93D));
            fxRatesListYesterday.Add(MapToFXRatesModel("CNH", 6.24D));
            fxRatesListYesterday.Add(MapToFXRatesModel("CNY", 6.22D));
            fxRatesListYesterday.Add(MapToFXRatesModel("COP", 1862.81D));
            fxRatesListYesterday.Add(MapToFXRatesModel("CZK", 19.74D));
            fxRatesListYesterday.Add(MapToFXRatesModel("CAD", 1.00D));
            return fxRatesListYesterday;
        }
    }
}
