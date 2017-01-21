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
            try
            {
                portFolio.OpenExcel();
                portFolio.FillBuy();
                portFolio.SaveExcel();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}",ex);
            }
            finally
            {
                portFolio.CloseExcel();
            }
            Console.WriteLine("Excel file {0} updated", portFolio.GetExcelFileName());
            Console.ReadKey();
        }
    }
}
