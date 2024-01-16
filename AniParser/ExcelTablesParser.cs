using AniParser.Entity;
using AniParser.Entity.TSN;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace AniParser
{
    internal static class ExcelTablesParser
    {
        static Excel.Application _excel = new Excel.Application();
        static Excel.Worksheet _ws;

        static Excel.Range RangeCode;
        static Excel.Range RangeWorksStart;
        static Excel.Range RangeDimensionStart;
        static Excel.Range RangeWorksEnd;
        static Excel.Range RangeClassStart;
        static Excel.Range RangeClassCurrent;
        static string currentHeader;

        static TSNCommonTable currentTSNCommonTable;
        static TSNCommonTableRecord currentTSNCommonTableRecord;

        public static void Init(Excel.Application excel, Excel.Worksheet ws)
        {
            _excel = excel;
            _ws = ws;
        }

        private static TSNCommonTable extractCommonTable(string rangeCodeAddress)
        {
            Range rangeCode = _ws.Range[rangeCodeAddress];
            RangeCode = rangeCode;
            return extractCommonTable(rangeCode);
        }

        public static TSNCommonTable extractCommonTable(Range rangeCode)
        {
            RangeCode = rangeCode;
            findRangeWorksStart();
            findRangeWorksEnd();
            findDimension();
            findRangeClassStart();
            return dataCollection();
        }

        private static void findRangeCode(string rangeName)
        {
            Range tmpRange = _ws.Range[rangeName];
            RangeCode = tmpRange.Find("Код");
            Debug.WriteLine(getRangeInfo(RangeCode));
        }
        private static void findRangeWorksStart()
        {
            RangeWorksStart = RangeCode.Offset[1, 1];
            //Debug.WriteLine(getRangeInfo(RangeWorksStart));
        }
        private static void findRangeWorksEnd()
        {
            RangeWorksEnd = RangeWorksStart.get_End(XlDirection.xlDown);
        }
        private static void findDimension()
        {
            RangeDimensionStart = RangeWorksStart.Offset[0, 1];
        }
        private static void findRangeClassStart()
        {
            RangeClassStart = RangeDimensionStart.Offset[0, 1].Offset[-1, 0];
        }

        static TSNCommonTable dataCollection()
        {
            currentTSNCommonTable = new TSNCommonTable(RangeCode.Offset[-2, 0].Value, RangeCode.Offset[-1, 0].Value);
            collectClassByRows(RangeClassStart);
            //Debug.WriteLine(JsonConvert.SerializeObject(currentTSNCommonTable));
            return currentTSNCommonTable;
        }

        private static void collectClassByRows(Range range) 
        {
            RangeClassCurrent = range;
            RangeClassCurrent.Select();
            if (RangeClassCurrent[1].Value == null) return;
            string recordType;
            try
            {
                recordType = RangeClassCurrent[1].Value.ToString("MM-d-yy");
            }
            catch (Exception e)
            {
                recordType = RangeClassCurrent[1].Value.ToString();
            }
            List<string> currentHeaderList = new List<string>();
            getClassHeadersInLine(currentHeaderList, "");
            currentHeaderList.Reverse();

            for (int i = RangeWorksStart.Row; i <= RangeWorksEnd.Row; i++)
            {
                currentTSNCommonTableRecord = new TSNCommonTableRecord();

                currentTSNCommonTableRecord.RecordType = recordType;
                currentTSNCommonTableRecord.Header = currentHeaderList;
                currentTSNCommonTableRecord.Code = $"{_ws.Cells[i, RangeCode.Column].Value}".Replace(Environment.NewLine, " ").Replace("\n", " ");
                currentTSNCommonTableRecord.Work = _ws.Cells[i, RangeWorksStart.Column].Value.ToString().Replace(Environment.NewLine, " ").Replace("\n", " ");
                currentTSNCommonTableRecord.Dimension = $"{_ws.Cells[i, RangeDimensionStart.Column].Value}";
                currentTSNCommonTableRecord.Value = $"{_ws.Cells[i, RangeClassCurrent.Column].Value}";

                currentTSNCommonTable.AddRecord(currentTSNCommonTableRecord);
            }
            collectClassByRows(RangeClassCurrent.Offset[0, 1]);
        }

        private static void getClassHeadersInLine(List<string> currentHeaderList, string header = "")
        {
            int nextRow = _excel.ActiveCell.Offset[-1, 0].Row;
            _ws.Cells[nextRow, RangeClassCurrent.Column].Select();
            Range range = _excel.Selection as Excel.Range;

            if (_excel.Selection[1].Row == RangeCode[1].Row)   // close method
            {
                currentHeaderList.Add(range[1].Value.ToString().Replace(Environment.NewLine, " ").Replace("\n", " "));
                currentHeader = $"{range[1].Value} {header}";
            }
            else
            {
                currentHeaderList.Add(range[1].Value.ToString().Replace(Environment.NewLine, " ").Replace("\n", " "));
                getClassHeadersInLine(currentHeaderList, $"{range[1].Value} {header}");
            }
        }
        private static string getRangeInfo( Range range )
        {
            return $"range: {range.Address} = {range[1].Value}, [{range.Row}, {range.Column}]";
        }
    }
}
