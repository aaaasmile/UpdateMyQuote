using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortfolioExcelChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            PortfolioExcel portFolio = new PortfolioExcel();
            portFolio.FillBuy();
        }
    }
}
