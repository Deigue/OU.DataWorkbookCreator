using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OU.DataWorkbookCreator.Models;
using OU.DataWorkbookCreator.Mappings;

namespace OU.DataWorkbookCreator.Repository
{
    static class TestRepository: OUDataMapping
    {
        
        public List<SummaryModel> SummaryList()
        {
            var summaryList = new List<SummaryModel>();
            summaryList.Add(MapToSummaryModel(12,65346));
            summaryList.Add(MapToSummaryModel(200, 646));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));
            summaryList.Add(MapToSummaryModel(12, 65346));



            return summaryList;
        }
    }
}
