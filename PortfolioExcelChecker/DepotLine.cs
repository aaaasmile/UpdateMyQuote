using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortfolioExcelChecker
{
    public class DepotLine
    {
        public DateTime Date { get; set; }
        public string Isin { get; set; }
        public string Description { get; set; }
        public double NumOfPieces { get; set; }
        public double Course { get; set; }
        public double CurrentValue { get; set; }
        public double BuyValue { get; set; }
        public double BuyValueOnDepotState { get; set; }
        public double SellValue { get; set; }
        public int Row { get; set; }
    }

    public class QuoteLine
    {
        public string Isin { get; set; }
        public string Place { get; set; }
        public double Quote { get; set; }
        public string LastUpdate { get; set; }
        public string Change { get; set; }
        public string Url { get; set; }
    }
}
