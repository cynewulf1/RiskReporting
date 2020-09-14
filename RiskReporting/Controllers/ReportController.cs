using System;
using Microsoft.AspNetCore.Mvc;
using RiskReporting.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;


namespace RiskReporting.Controllers
{
    public class ReportController : Controller
    {
        ILogger _logger;
        private IConfiguration _configuration;

        public ReportController(IConfiguration configuration, ILogger<ReportController> logger)
        {
            _logger = logger;
            _logger.LogInformation("ReportController instantiated");

            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult Index(ReportModel model)
        {
            try
            {
                // Load the XML
                _logger.LogDebug("Loading XML");
                var dataLoader = new XmlDataLoader(_configuration, _logger);
                var xml = dataLoader.LoadData(Request.Form.Files[1]);

                // Load the HTML and populate with the XML data (single pass)
                _logger.LogDebug("Converting Excel to HTML");
                var builder = new ExcelReportBuilder(_configuration, _logger);
                var html = builder.GenerateReport(Request.Form.Files[0], xml);

                _logger.LogDebug("Outputting HTML as Content");
                // Output the content as HTML
                return Content(html, "text/html; charset=UTF-8");
            }
            catch (Exception ex)
            {
                _logger.LogError($"{ex.Message}\r\n{ex.StackTrace}");
                return Content($"{ex.Message}\r\n{ex.StackTrace}");
            }
        }
    }
}