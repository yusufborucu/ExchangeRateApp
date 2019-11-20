using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExchangeRateApp.Web.Models
{
    public class ChartDataset
    {
        public string label { get; set; }
        public string backgroundColor { get; set; }
        public string borderColor { get; set; }
        public List<string> data { get; set; }
        public bool fill { get; set; }
    }
}