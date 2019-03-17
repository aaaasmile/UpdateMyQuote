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
            String excelFileName = @"D:\Documents\easybank\Portfolio.xlsm";
            PortfolioExcel portFolio = new PortfolioExcel();
            try
            {
                QuoteUpdaterLauncher updater = new QuoteUpdaterLauncher();
                updater.TeminatedEvent += (x) => 
                {
                    portFolio.OpenExcel(excelFileName);
                    //portFolio.FillBuy();
                    portFolio.UpdateQuote();
                    //portFolio.SaveExcel();
                    portFolio.Activate();
                };

                //updater.CheckVersion();
                String getQuotePath = @"D:\Projects\GItHub\ruby_scratch\finanz_net\get_quote.rb";
                Console.WriteLine("I am using hard coded paths, for your programming skill this is not an issue. Isn't it?");
                updater.StartProcess(getQuotePath);
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}",ex);
            }
            finally
            {
                //portFolio.CloseExcel();
            }
            Console.WriteLine("Excel file {0} updated\nPress any key to continue", portFolio.GetExcelFileName());
            Console.ReadKey();
        }
    }
}
