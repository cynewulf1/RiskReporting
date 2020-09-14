using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Xml;

namespace RiskReporting.Models
{
    public class XmlDataLoader : IDataLoader
    {
        ILogger _logger;
        private DatafileOptions datafileOptions = new DatafileOptions();

        public XmlDataLoader(IConfiguration configuration, ILogger logger)
        {
            try
            {
                _logger = logger;
                _logger.LogInformation("XmlDataLoader instantiated");

                // Bind the config data
                configuration.GetSection(DatafileOptions.Datafile).Bind(datafileOptions);

            }
            catch (Exception ex)
            {
                _logger.LogError($"{ex.Message}\r\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Load the XML data into reportsData
        /// </summary>
        public IList<ReportData> LoadData(IFormFile dataFile)
        {
            try
            {
                List<ReportData> reportsData = new List<ReportData>();

                if (dataFile != null && dataFile.Length > 0 && datafileOptions != null)
                {
                    // Load the XML
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(dataFile.OpenReadStream());

                    // Get the report nodes
                    XmlNodeList reports = xmlDoc.SelectNodes($"{datafileOptions.RootNode}/{datafileOptions.ItemNode}");

                    // For each report node...
                    foreach (XmlNode report in reports)
                    {
                        ReportData dataItem = new ReportData();
                        // Get the report ID
                        dataItem.ReportId = report.SelectSingleNode(datafileOptions.TabNode).InnerText;

                        // Get the report values
                        XmlNodeList items = report.SelectNodes(datafileOptions.ValueNode);

                        dataItem.DataItems = new List<ReportDataItem>();
                        foreach (XmlNode item in items)
                        {
                            // Add a data item to the collection
                            dataItem.DataItems.Add(new ReportDataItem()
                            {
                                Row = item[datafileOptions.ValueRow].InnerText,
                                Col = item[datafileOptions.ValueColumn].InnerText,
                                Val = item[datafileOptions.Value].InnerText
                            });
                        }
                        reportsData.Add(dataItem);
                    }
                }

                return reportsData;
            }
            catch (Exception ex)
            {
                _logger.LogError($"{ex.Message}\r\n{ex.StackTrace}");
                return null;
            }
        }
    }
}
