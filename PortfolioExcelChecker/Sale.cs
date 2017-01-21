using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortfolioExcelChecker
{
    public enum SaleOperationEnum
    {
        Sell,
        Buy,
        Capital
    }

    public class Sale
    {
        public SaleOperationEnum SaleOperation { get; set; }
        public DateTime Date { get; set; }
        public string Isn { get; set; }
        public string Description { get; set; }
        public string Place { get; set; }
        public double NumPiecesInOrder { get; set; }
        public double NumPiecesExecuted { get; set; }
        public double Price { get; set; }

        public const string SELL_CAPTION = "Verkauf";
        public const string BUY_CAPTION = "Kauf";
        public const string CAPITAL = "Kapitalmaßnahme";
    }
}
