using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OU.DataWorkbookCreator.Models
{
    public class DetailReportModel
    {
        public string Desk { get; set; }
        public string OrgUnit { get; set; }
        public string CCY { get; set; }
        public string Account { get; set; }
        public string Product { get; set; }
        public string CustGroup { get; set; }
        public double CGLM { get; set; }
        public double T1GLM { get; set; }
    }
}
