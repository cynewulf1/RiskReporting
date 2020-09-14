using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RiskReporting.Models
{
    public class HtmlFormattingOptions
    {
        public const string HtmlFormatting = "HtmlFormatting";

        public string HeaderHtml { get; set; }
        public string FooterHtml { get; set; }
        public string TableStartHtml { get; set; }
        public string TableEndHtml { get; set; }
    }
}
