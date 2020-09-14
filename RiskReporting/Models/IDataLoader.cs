using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace RiskReporting.Models
{
    public interface IDataLoader
    {
        IList<ReportData> LoadData(IFormFile dataFile);
    }
}