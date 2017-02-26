using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PortfolioExcelChecker
{
    class PortfolioExcel
    {
        private Excel.Application _xlApp;
        private Excel.Workbook _xlWorkBook;
        private Excel.Worksheet _xlWorkSheetSales;
        private Excel.Worksheet _xlWorkSheetDepot;
        private Excel.Worksheet _xlWorkSheetPortfolio;

        private Dictionary<string, QuoteLine> _dctQuote = new Dictionary<string, QuoteLine>();

        internal void Activate()
        {
            _xlApp.Visible = true;
        }

        internal void OpenExcel()
        {
            _xlApp = new Excel.Application();
            _xlWorkBook = _xlApp.Workbooks.Open(GetExcelFileName(), 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            _xlWorkSheetSales = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(2);
            _xlWorkSheetDepot = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(3);
            _xlWorkSheetPortfolio = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(1);
        }
        internal void SaveExcel()
        {
            if (_xlApp != null)
            {
                object misValue = System.Reflection.Missing.Value;

                _xlWorkBook.Close(true, misValue, misValue);
            }
        }

        internal void CloseExcel()
        {
            if (_xlApp != null)
            {
                _xlApp.Quit();

                releaseObject(_xlWorkSheetSales);
                releaseObject(_xlWorkSheetDepot);
                releaseObject(_xlWorkBook);
                releaseObject(_xlApp);
                _xlApp = null;
            }
        }

        internal void UpdateQuote()
        {
            string[] lines = System.IO.File.ReadAllLines("quote.csv");
            _dctQuote.Clear();
            foreach (var line in lines)
            {
                var columns = line.Split(';');
                var quote = new QuoteLine()
                {
                    Isin = columns[0],
                    Change = columns[4],
                    LastUpdate = columns[3],
                    Place = columns[1],
                    Quote = double.Parse(columns[2]),
                    Url = columns[5]
                };
                _dctQuote.Add(columns[0], quote);
            }
            for (int i = 4; i < 100; i++)
            {
                string isin = GetCellString("B", i, _xlWorkSheetPortfolio);
                if (string.IsNullOrEmpty(isin))
                    break;

                QuoteLine currentQuote;
                if(_dctQuote.TryGetValue(isin, out currentQuote))
                {
                    _xlWorkSheetPortfolio.Cells[i, "H"] = currentQuote.Quote;
                    _xlWorkSheetPortfolio.Cells[i, "I"] = currentQuote.LastUpdate;
                    _xlWorkSheetPortfolio.Cells[i, "AA"] = currentQuote.Url;
                }
                else
                {
                    Console.WriteLine("ISIN {0} from portfolio not found", isin);
                }
            }

            for (int i = 4; i < 100; i++)
            {
                string isin = GetCellString("A", i, _xlWorkSheetDepot);
                if (string.IsNullOrEmpty(isin))
                    break;

                QuoteLine currentQuote;
                if (_dctQuote.TryGetValue(isin, out currentQuote))
                {
                    _xlWorkSheetDepot.Cells[i, "E"] = currentQuote.Quote;
                    _xlWorkSheetDepot.Cells[i, "O"] = currentQuote.LastUpdate;
                    _xlWorkSheetDepot.Cells[i, "P"] = currentQuote.Url;
                }
                else
                {
                    Console.WriteLine("ISIN {0} from depot not found", isin);
                }
            }
        }

        internal void FillBuy()
        {
            List<Sale> sales = ExtractSales(_xlWorkSheetSales);
            List<DepotLine> depotLines = ExtractDepotLines(_xlWorkSheetDepot);
            CalculateBuyValue(sales, depotLines);
            WriteBuyValue(depotLines, _xlWorkSheetDepot);
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
                string isn = depotLine.Isin;
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
                depotLine.BuyValue = Math.Round(buyValue, 2);
                depotLine.BuyValueOnDepotState = Math.Round(buyValue / purchItems * depotLine.NumOfPieces, 2);

                double sellValue = 0.0;
                foreach (var sellLine in sellLines)
                {
                    sellValue += Math.Abs(sellLine.NumPiecesExecuted * sellLine.Price);
                }
                depotLine.SellValue = Math.Round(sellValue, 2);
            }
        }

        private List<DepotLine> ExtractDepotLines(Excel.Worksheet xDepotlWorkSheet)
        {
            List<DepotLine> depotLines = new List<DepotLine>();
            bool searchData = true;
            DepotLine currItem;
            for (int i = 4; i < 300; i++)
            {
                string nextCellLbl = string.Format("A{0}", i);
                if (xDepotlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 == null)
                {
                    if (searchData)
                        continue;
                    else
                        break;
                }
                if (searchData)
                {
                    if (xDepotlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2 is double)
                    {
                        searchData = false;
                    }
                    else
                        continue;
                }
                currItem = new DepotLine();
                currItem.Date = DateTime.FromOADate(xDepotlWorkSheet.get_Range(nextCellLbl, nextCellLbl).Value2);
                currItem.Isin = GetCellString("B", i, xDepotlWorkSheet);
                currItem.Description = GetCellString("C", i, xDepotlWorkSheet);
                currItem.NumOfPieces = GetCellDouble("D", i, xDepotlWorkSheet);
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
