using Microsoft.Office.Interop.Excel;
using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace AniParser
{
    internal static class ExcelTablesParser
    {
        static Excel.Application _excel = new Excel.Application();
        static Excel.Workbook _wb;
        static Excel.Worksheet _ws;

        static Excel.Range RangeCode;
        static Excel.Range RangeWorksStart;
        static Excel.Range RangeDimensionStart;
        static Excel.Range RangeWorksEnd;
        static Excel.Range RangeClassStart;
        static Excel.Range RangeClassCurrent;
        static string currentHeader;

        static StringBuilder sb;

        public static void Parse(string path)
        {
            if (path == "" || path == null)
            {
                Debug.WriteLine($"path is Empty");
                return;
            }
            try
            {
                OpenWorkBook(path);
                _excel.Visible = true;

                findRangeCode();
                findRangeWorksStart();
                findRangeWorksEnd();
                findDimension();
                findRangeClassStart();
                dataCollection();
                // _ws.get_Range("")
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"error: {ex.Message}");
            }
        }

        private static void OpenWorkBook(string path)
        {
            _wb = _excel.Workbooks.Open(path);
            _ws = (Excel.Worksheet)_wb.ActiveSheet;
            _ws.Range["A1"].Select();
        }

        private static void findRangeCode()
        {
            RangeCode = _ws.Cells.Find("Код");
            //Debug.WriteLine(getRangeInfo(RangeCode));
        }
        private static void findRangeWorksStart()
        {
            RangeWorksStart = RangeCode.Offset[1, 1];
            //Debug.WriteLine(getRangeInfo(RangeWorksStart));
        }
        private static void findRangeWorksEnd()
        {
            RangeWorksEnd = RangeWorksStart.get_End(XlDirection.xlDown);
            //Debug.WriteLine(getRangeInfo(RangeWorksEnd));
        }
        private static void findDimension()
        {
            RangeDimensionStart = RangeWorksStart.Offset[0, 1];
            //Debug.WriteLine(getRangeInfo(RangeDimensionStart));
        }
        private static void findRangeClassStart()
        {
            RangeClassStart = RangeDimensionStart.Offset[0, 1].Offset[-1, 0];
        }

        static void dataCollection()
        {
            sb = new StringBuilder();
            sb.AppendLine(RangeCode.Offset[-1, 0].Value); // забираем строку измерение перед таблицей
            collectClassByRows(RangeClassStart);
            Debug.WriteLine(sb.ToString());
        }

        private static void collectClassByRows(Range range) 
        {
            RangeClassCurrent = range;
            RangeClassCurrent.Select();
            if (RangeClassCurrent[1].Value == null) return;
            string razdel = RangeClassCurrent[1].Value.ToString("MM-d-yy");
            //Debug.WriteLine($"{razdel}");

            getClassHeadersInLine("");
            //Debug.WriteLine($"{currentHeader}");

            for (int i = RangeWorksStart.Row; i <= RangeWorksEnd.Row; i++)
            {
                sb.AppendLine($"{razdel} {currentHeader} {_ws.Cells[i,RangeCode.Column].Value} {_ws.Cells[i, RangeWorksStart.Column].Value} {_ws.Cells[i, RangeDimensionStart.Column].Value} {_ws.Cells[i, RangeClassCurrent.Column].Value}");
            }
            collectClassByRows(RangeClassCurrent.Offset[0, 1]);
        }

        private static void getClassHeadersInLine(string header = "")
        {
            int nextRow = _excel.ActiveCell.Offset[-1, 0].Row;
            _ws.Cells[nextRow, RangeClassCurrent.Column].Select();
            Range range = _excel.Selection as Excel.Range;

            if (_excel.Selection[1].Row == RangeCode[1].Row)   // close method
            {
                currentHeader = $"{range[1].Value} {header}";
            }
            else
            {
                getClassHeadersInLine($"{range[1].Value} {header}");
            }
        }
        private static string getRangeInfo( Range range )
        {
            return $"range: {range.Address} = {range[1].Value}, [{range.Row}, {range.Column}]";
        }
    }
}
