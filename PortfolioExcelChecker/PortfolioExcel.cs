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
            var xlWorkBook = xlApp.Workbooks.Open(GetExcelFileName(), 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheetSales = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            var xlWorkSheetDepot = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            object misValue = System.Reflection.Missing.Value;

            List<Sale> sales = ExtractSales(xlWorkSheetSales);
            List<DepotLine> depotLines = ExtractDepotLines(xlWorkSheetDepot);
            CalculateBuyValue(sales, depotLines);
            WriteBuyValue(depotLines, xlWorkSheetDepot);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheetSales);
            releaseObject(xlWorkSheetDepot);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        internal string GetExcelFileName()
        {
            return @"D:\Documents\easybank\Portfolio.xlsx";
        }

        private void WriteBuyValue(List<DepotLine> depotLines, Excel.Worksheet xlWorkSheetDepot)
        {
            foreach (var depotLine in depotLines)
            {
                xlWorkSheetDepot.Cells[depotLine.Row, "J"] = depotLine.BuyValue;
                xlWorkSheetDepot.Cells[depotLine.Row, "K"] = depotLine.SellValue;
                xlWorkSheetDepot.Cells[depotLine.Row, "L"] = depotLine.BuyValueOnDepotState;
            }
        }

        private void CalculateBuyValue(List<Sale> sales, List<DepotLine> depotLines)
        {
            foreach (var depotLine in depotLines)
            {
                string isn = depotLine.Isn;
                var lastCapital = sales.Where(x => x.Isn.Equals(isn) && x.SaleOperation == SaleOperationEnum.Capital)
                        .OrderBy(x => x.Date).LastOrDefault();

                var purchLines = sales.Where(x => x.Isn.Equals(isn) && x.SaleOperation == SaleOperationEnum.Buy).ToArray();
                double piecesBuyed = purchLines.Sum(x => x.NumPiecesExecuted);
                var sellLines = sales.Where(x => x.Isn.Equals(isn) && x.SaleOperation == SaleOperationEnum.Sell);
                var numSell = sellLines.Sum(x => x.NumPiecesExecuted);

                numSell = Math.Abs(numSell);
                if ((piecesBuyed - numSell) != depotLine.NumOfPieces)
                {
                    bool argErr = true;
                    if (lastCapital != null)
                    {
                        var recentPurch = sales.Where(x => x.Isn.Equals(isn) &&
                                x.SaleOperation == SaleOperationEnum.Buy && x.Date > lastCapital.Date).Sum(x => x.NumPiecesExecuted);

                        var sumRounded = Math.Round(recentPurch + lastCapital.NumPiecesExecuted - numSell, 2);
                        if (sumRounded == depotLine.NumOfPieces)
                            argErr = false;
                    }
                    if (argErr)
                        throw new ArgumentOutOfRangeException(string.Format("Not all items in depot are buyed: isn {0}", isn));
                }

                double buyValue = 0.0;
                double purchItems = 0.0;
                foreach (var purchLine in purchLines)
                {
                    buyValue += purchLine.NumPiecesExecuted * purchLine.Price;
                    purchItems += purchLine.NumPiecesExecuted;
                }
                depotLine.BuyValue = Math.Round(buyValue,2);
                depotLine.BuyValueOnDepotState = Math.Round(buyValue / purchItems * depotLine.NumOfPieces, 2);

                double sellValue = 0.0;
                foreach (var sellLine in sellLines)
                {
                    sellValue += Math.Abs(sellLine.NumPiecesExecuted * sellLine.Price);
                }
                depotLine.SellValue = Math.Round(sellValue, 2);
            }
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
                currItem.NumOfPieces = GetCellDouble("D", i, xlWorkSheet);
                currItem.Row = i;
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
                string saleType = GetCellString("B", i, xlWorkSheet);
                switch (saleType)
                {
                    case Sale.BUY_CAPTION:
                        currItem.SaleOperation = SaleOperationEnum.Buy;
                        break;
                    case Sale.SELL_CAPTION:
                        currItem.SaleOperation = SaleOperationEnum.Sell;
                        break;
                    case Sale.CAPITAL:
                        currItem.SaleOperation = SaleOperationEnum.Capital;
                        break;
                    default:
                        throw new ArgumentException(string.Format("Sale type {0} not recognized", saleType));
                }
                currItem.Isn = GetCellString("M", i, xlWorkSheet);
                currItem.Description = GetCellString("D", i, xlWorkSheet);
                currItem.NumPiecesInOrder = GetCellDouble("G", i, xlWorkSheet);
                currItem.NumPiecesExecuted = GetCellDouble("I", i, xlWorkSheet);
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
            catch (Exception)
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
