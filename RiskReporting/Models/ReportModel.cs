using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RiskReporting.Models
{
    public class ReportModel
    {
        public IFormFile ExcelFile { get; set; }
        public IFormFile DataFile { get; set; }
    }
}
