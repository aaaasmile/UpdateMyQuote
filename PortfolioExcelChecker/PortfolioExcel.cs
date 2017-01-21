using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PortfolioExcelChecker
{
    class PortfolioExcel
    {
        internal void FillBuy()
        {
            var xlApp = new Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Open(@"D:\Documents\easybank\Portfolio.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheetSales = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            var xlWorkSheetDepot = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            object misValue = System.Reflection.Missing.Value;

            List<Sale> sales = ExtractSales(xlWorkSheetSales);
            List<DepotLine> depotLines = ExtractDepotLines(xlWorkSheetDepot);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheetSales);
            releaseObject(xlWorkSheetDepot);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private List<DepotLine> ExtractDepotLines(Excel.Worksheet xlWorkSheet)
        {
            List<DepotLine> depotLines = new List<DepotLine>();
            bool searchData = true;
            DepotLine currItem;
            for (int i = 4; i < 300; i++)
            {
                string nextCellLbl = string.Format("A{0}", i);
                if (xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 == null)
                {
                    if (searchData)
                        continue;
                    else
                        break;
                }
                if (searchData)
                {
                    if (xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 is double)
                    {
                        searchData = false;
                    }
                    else
                        continue;
                }
                currItem = new DepotLine();
                currItem.Date = DateTime.FromOADate(xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2);
                currItem.Isn = GetCellString("B", i, xlWorkSheet);
                currItem.Description = GetCellString("C", i, xlWorkSheet);
                currItem.NumOfPieces = GetCellInteger("D", i, xlWorkSheet);
                depotLines.Add(currItem);
            }

            return depotLines;
        }

        private static List<Sale> ExtractSales(Excel.Worksheet xlWorkSheet)
        {
            bool searchData = true;
            Sale currItem;
            List<Sale> saleList = new List<Sale>();
            for (int i = 4; i < 300; i++)
            {
                string nextCellLbl = string.Format("A{0}", i);
                if (xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 == null)
                {
                    if (searchData)
                        continue;
                    else
                        break;
                }
                if (searchData)
                {
                    if (xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 is double)
                    {
                        searchData = false;
                    }
                    else
                        continue;
                }
                currItem = new Sale();
                currItem.Date = DateTime.FromOADate(xlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2);
                var SaleType = GetCellString("B", i, xlWorkSheet);
                currItem.SaleOperation = SaleType.Equals(Sale.BUY_CAPTION) ? SaleOperationEnum.Buy : SaleOperationEnum.Sell;
                currItem.Isn = GetCellString("C", i, xlWorkSheet);
                currItem.Description = GetCellString("D", i, xlWorkSheet);
                currItem.NumPiecesInOrder = GetCellInteger("G", i, xlWorkSheet);
                currItem.NumPiecesExecuted = GetCellInteger("I", i, xlWorkSheet);
                currItem.Price = GetCellDouble("K", i, xlWorkSheet);
                saleList.Add(currItem);
            }
            return saleList;
        }

        private static string GetCellString(string column, int row, Excel.Worksheet xlWorkSheet)
        {
            string result = string.Empty;
            string range = string.Format("{0}{1}", column, row);
            if (xlWorkSheet.get_Range(range, range).Value2 != null)
                result = xlWorkSheet.get_Range(range, range).Value2.ToString();
            return result;
        }

        private static double GetCellDouble(string column, int row, Excel.Worksheet xlWorkSheet)
        {
            double result = 0.0;
            string range = string.Format("{0}{1}", column, row);
            if (xlWorkSheet.get_Range(range, range).Value2 != null)
            {
                double.TryParse(xlWorkSheet.get_Range(range, range).Value2.ToString(), out result);
            }
            return result;
        }

        private static int GetCellInteger(string column, int row, Excel.Worksheet xlWorkSheet)
        {
            int result = 0;
            string range = string.Format("{0}{1}", column, row);
            if (xlWorkSheet.get_Range(range, range).Value2 != null)
                int.TryParse(xlWorkSheet.get_Range(range, range).Value2.ToString(), out result);
            return result;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
