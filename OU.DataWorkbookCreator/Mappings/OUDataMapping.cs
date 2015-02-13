using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OU.DataWorkbookCreator.Models;

namespace OU.DataWorkbookCreator.Mappings
{
    public class OUDataMapping
    {
        public SummaryModel MapToSummaryModel(double orgUnit, double amount)
        {
            var summaryModel = new SummaryModel();
            summaryModel.OrgUnit = orgUnit;
            summaryModel.Amount = amount;
            return summaryModel;
        }

        public FootingsModel MapToFootingsModel(string desk, string orgUnit, string CCY, double CGLM, double CAdjustment, double T1GLM, double T1Ajustment)
        {
            var footingsModel = new FootingsModel();
            footingsModel.Desk = desk;
            footingsModel.OrgUnit = orgUnit;
            footingsModel.CCY = CCY;
            footingsModel.CGLM = CGLM;
            footingsModel.CAdjustment = CAdjustment;
            footingsModel.T1GLM = T1GLM;
            footingsModel.T1Adjustment = T1Ajustment;
            return footingsModel;
        }

        public DetailReportModel MapToDetailReportModel(string desk, string orgUnit, string CCY, string account, string product, string custGroup, double CGLM, double T1GLM)
        {
            var detailReportModel = new DetailReportModel();
            detailReportModel.Desk = desk;
            detailReportModel.OrgUnit = orgUnit;
            detailReportModel.CCY = CCY;
            detailReportModel.Account = account;
            detailReportModel.CustGroup = custGroup;
            detailReportModel.CGLM = CGLM;
            detailReportModel.T1GLM = T1GLM;
            return detailReportModel;
        }

        public FXRatesModel MapToFXRatesModel(string CCY, double rate)
        {
            var fxRatesModel = new FXRatesModel();
            fxRatesModel.CCY = CCY;
            fxRatesModel.Rate = rate;
            return fxRatesModel;
        }
    }
}
