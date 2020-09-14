using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace RiskReporting.Models
{
    public interface IReportBuilder
    {
        string GenerateReport(IFormFile file, IList<ReportData> reports);
    }
}