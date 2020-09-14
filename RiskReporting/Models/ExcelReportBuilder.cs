using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiskReporting.Models
{
    public class ExcelReportBuilder : IReportBuilder
    {
        ILogger _logger;
        private HtmlFormattingOptions htmlFormattingOptions = new HtmlFormattingOptions();

        public ExcelReportBuilder(IConfiguration configuration, ILogger logger)
        {
            try
            {
                _logger = logger;
                _logger.LogInformation("ExcelReportBuilder instantiated");

                // Bind the config data
                configuration.GetSection(HtmlFormattingOptions.HtmlFormatting).Bind(htmlFormattingOptions);
            }
            catch (Exception ex)
            {        
                _logger.LogError($"{ex.Message}\r\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Read the Excel file and produce the HTML report(s).
        /// </summary>
        /// <param name="file">The Excel file</param>
        /// <param name="reports">The report(s) data</param>
        /// <returns>HTML reports</returns>
        public string GenerateReport(IFormFile file, IList<ReportData> reports)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                // Append the html header data
                sb.Append(htmlFormattingOptions.HeaderHtml);

                // For each report (read from the XML)...
                foreach (var report in reports)
                {
                    Dictionary<int, int> columnReferences = new Dictionary<int, int>();
                    Dictionary<int, int> rowReferences = new Dictionary<int, int>();
                    int columnReferenceRow = int.MaxValue;

                    // Append header
                    sb.Append($"<hr><h1>{report.ReportId}</h1><hr>");

                    // Append the table start HTML
                    sb.Append(htmlFormattingOptions.TableStartHtml);

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage excelPackage = new ExcelPackage(file.OpenReadStream()))
                    {
                        ExcelWorkbook workbook = excelPackage.Workbook;
                        if (workbook != null)
                        {
                            // Load the worksheet for the relevant report ID (e.g. "F 20.04")
                            ExcelWorksheet worksheet = workbook.Worksheets[report.ReportId];
                            if (worksheet != null)
                            {
                                // Get a list of merged cells
                                var merged = worksheet.MergedCells;

                                // Find the limits of the worksheet
                                int maxRows = worksheet.Dimension.Rows;
                                int maxCols = worksheet.Dimension.Columns;

                                // For each row in the worksheet...
                                for (int y = 1; y <= maxRows; y++)
                                {
                                    sb.Append("<tr>");
                                    // For each column in the worksheet...
                                    for (int x = 1; x <= maxCols; x++)
                                    {
                                        // Get current cell
                                        var cell = worksheet.Cells[y, x];
                                        string cellValue = "";
                                        if (cell.Value != null) cellValue = cell.Value.ToString();

                                        // If we're beyond the column reference row then we need to start looking for row references
                                        if (y > columnReferenceRow)
                                        {
                                            bool rowReference = IsReferenceCell(report.DataItems, cellValue, false);
                                            if (rowReference) rowReferences.Add(y, int.Parse(cellValue));
                                        }
                                        else
                                        {
                                            // Check the cell to see if it matches a column value
                                            bool columnReference = IsReferenceCell(report.DataItems, cell.Value.ToString(), true);
                                            if (columnReference)
                                            {
                                                // Add to the column references collection
                                                columnReferences.Add(x, int.Parse(cell.Value.ToString()));

                                                // Set the column reference row
                                                columnReferenceRow = y;
                                            }
                                        }

                                        // Check if we're at both a row reference and a column reference
                                        if (columnReferences.Count > 0 && rowReferences.Count > 0)
                                        {
                                            if (columnReferences.ContainsKey(x) && rowReferences.ContainsKey(y))
                                                cellValue = report.DataItems.SingleOrDefault(d => int.Parse(d.Col) == columnReferences[x] && int.Parse(d.Row) == rowReferences[y]).Val;
                                        }

                                        // Get the current background colour and add the bgcolor tag accordingly
                                        var bg = cell.Style.Fill.BackgroundColor.Rgb;
                                        string bgtag = "";
                                        if (bg != null) bgtag = $" bgcolor='#{bg.ToString().Substring(2, bg.ToString().Length - 2)}'";

                                        // Add the bold/italic tags if required
                                        if (cell.Style.Font.Bold || cell.Style.Font.Italic)
                                        {
                                            bgtag += " style='";
                                            if (cell.Style.Font.Bold) bgtag += "font-weight:bold;";
                                            if (cell.Style.Font.Italic) bgtag += "font-style:italic;";
                                            bgtag += "'";
                                        }

                                        // If this is a merged cell but not the first cell in the merge then don't output anything (to avoid repeats)
                                        if (!IsFirstMergedCell(merged, y, x)) cellValue = "";

                                        // Append the html to the string builder
                                        sb.Append($"<td {bgtag}>{cellValue}</td>");
                                    }
                                    sb.Append("</tr>");
                                }
                            }
                        }
                    }

                    // Append the table end HTML
                    sb.Append(htmlFormattingOptions.TableEndHtml);
                }
                // Append the html footer data
                sb.Append(htmlFormattingOptions.FooterHtml);

                return sb.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError($"{ex.Message}\r\n{ex.StackTrace}");
                return null;
            }
        }

        bool IsReferenceCell(List<ReportDataItem> data, string cellValue, bool isColumn)
        {
            // Check if this is a reference cell from the XML data
            int.TryParse(cellValue, out int cv);
            if (cv != 0)
            {
                if (isColumn)
                {
                    if (data.Exists(x => x.Col == cv.ToString())) return true;
                }
                else
                {
                    if (data.Exists(x => x.Row == cv.ToString())) return true;
                }
            }
            return false;
        }

        bool IsFirstMergedCell(ExcelWorksheet.MergeCellsCollection merged, int row, int column)
        {
            // Check if this cell is the first in a merged cell group
            // Also returns true if it's not a merged cell (non-merged cells are always first in their own group of 1)

            string m = merged[row, column];
            if (m == null) return true;

            var addr = new ExcelAddress(m);
            return (row == addr.Start.Row && column == addr.Start.Column);
        }

        string GetSpan(ExcelWorksheet.MergeCellsCollection merged, int row, int column)
        {
            // Unused code. The row/colspan was producing strange results.
            string span = "";

            string m = merged[row, column];
            if (m == null) return "";

            var addr = new ExcelAddress(m);
            int rowspan = 1 + addr.End.Row - addr.Start.Row;
            int colspan = 1 + addr.End.Column - addr.Start.Column;
            if (rowspan > 1) span += $" rowspan={rowspan}";
            if (colspan > 1) span += $" colspan={colspan}";

            return span;
        }
    }
}
