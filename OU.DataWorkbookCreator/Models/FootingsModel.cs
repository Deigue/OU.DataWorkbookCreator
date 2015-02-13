using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OU.DataWorkbookCreator.Models
{
    public class FootingsModel
    {
        public string Desk { get; set; }
        public string OrgUnit { get; set; }
        public string CCY { get; set; }
        public double CGLM { get; set; }
        public double CAdjustment { get; set; }
        public double T1GLM { get; set; }
        public double T1Adjustment { get; set; }
    }
}
