using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RiskReporting.Models
{
    public class DatafileOptions
    {
        public const string Datafile = "Datafile";

        public string RootNode { get; set; }
        public string ItemNode { get; set; }
        public string TabNode { get; set; }
        public string ValueNode { get; set; }
        public string ValueColumn { get; set; }
        public string ValueRow { get; set; }
        public string Value { get; set; }
    }
}
