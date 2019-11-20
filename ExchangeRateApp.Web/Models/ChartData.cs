using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExchangeRateApp.Web.Models
{
    public class ChartData
    {
        public List<string> labels { get; set; }
        public List<ChartDataset> datasets { get; set; }
    }
}