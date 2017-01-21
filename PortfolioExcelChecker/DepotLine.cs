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
        public string Isn { get; set; }
        public string Description { get; set; }
        public double NumOfPieces { get; set; }
        public double Course { get; set; }
        public double CurrentValue { get; set; }
        public double BuyValue { get; set; }
        public int Row { get; set; }
    }
}
