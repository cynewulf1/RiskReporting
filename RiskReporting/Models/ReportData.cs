using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RiskReporting.Models
{
    public class ReportData
    {
        public string ReportId { get; set; }
        public List<ReportDataItem> DataItems { get; set; }
    }
}
