using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AniParser.Entity.TSN
{
    public class TSNCompilationExtractor
    {
        public static Excel.Application _excel = new Excel.Application();
        private static Excel.Workbook _wb;
        public static Excel.Worksheet _ws;
        public Compilation compilation;

        private Branch currentBranch;
        private Section currentSection;
        public void Parse( string path)
        {
            _ws = OpenWorkBook(path);
            ExcelTablesParser.Init(_excel, _ws);
            Compilation compilation = new Compilation();
            getContentName(compilation);
            getContentRange(out (Range, Range) fromToRange);
            string fromToAddress = $"{fromToRange.Item1.Address}:{fromToRange.Item2.Address}";
            Range contentRange = _ws.Range[fromToAddress];
            for (int i = fromToRange.Item1.Row+1; i < fromToRange.Item2.Row; i++)
            {
                string address = "A" + i ;
                string strokeValue = $"{_ws.Range[address].Value}".Trim();
                Match match = Regex.Match(strokeValue, TSNRegexPatterns.NumberAtTheEndOfTheLinePattern, RegexOptions.IgnoreCase);

                if (strokeValue.StartsWith("Отдел"))
                {
                    currentSection = null;
                    currentBranch = compilation.AddBranch(new Branch());
                    currentBranch.Address = address;
                    
                    if (match != null)
                    {
                        currentBranch.branch_name = strokeValue.Remove(match.Index, match.Length);
                    }
                    else
                    {
                        currentBranch.branch_name = strokeValue;
                    }
                }
                else if (strokeValue.StartsWith("Раздел"))
                {
                    currentSection = currentBranch.AddSection(new Section());
                    currentSection.Address = address;

                    if (match != null)
                    {
                        currentSection.section_name = strokeValue.Remove(match.Index, match.Length);
                    }
                    else
                    {
                        currentSection.section_name = strokeValue;
                    }
                }
                else if (strokeValue.StartsWith("Таблица"))
                {
                    string t_name;
                    if (match != null)
                    {
                        t_name = strokeValue.Remove(match.Index, match.Length);
                    }
                    else
                    {
                        t_name = strokeValue;
                    }
                    if (currentSection == null)
                    {
                        currentSection = currentBranch.AddSection(new Section());
                        currentSection.section_name = "_";
                    }
                    currentSection.addTableLink(t_name, address);
                }
            }
            parseTables(compilation);
            Debug.WriteLine($"{JsonConvert.SerializeObject(compilation)}");
        }

        public static Excel.Worksheet OpenWorkBook(string path)
        {
            if (path == "" || path == null)
            {
                Debug.WriteLine($"path is Empty");
                return null;
            }
            _wb = _excel.Workbooks.Open(path);
            _ws = (Excel.Worksheet)_wb.ActiveSheet;
            _excel.Visible = true;
            return _ws;
        }
        public async void preProcessor(string path)
        {
            try
            {
                {
                    _ws = OpenWorkBook(path);
                    Excel.Range last = _ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    Debug.WriteLine($"{last.Address}");
                    //((Range)ws.Rows[2, Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
                    bool emptyStroke;
                    bool mergedCells;
                    bool emptyString;
                    bool isNull;
                    List<int> rows4deleting = new List<int>();
                    for (int row = 1; row <= last.Row; row++)
                    {
                        emptyStroke = true;
                        for (int column = 1; column <= last.Column; column++)
                        {

                            Range r = (Range)_ws.Cells[row, column];
                            mergedCells = r?.MergeCells;
                            emptyString = r?.Value?.ToString() == "";
                            isNull = r?.Value == null;
                            if (
                                mergedCells ||
                                !isNull)
                            {
                                emptyStroke = false;
                                Debug.WriteLine($"{r.Address} : val:{r?.Value?.ToString()} empty:{emptyString} merg:{mergedCells} null:{isNull}");
                                break;
                            }
                        }
                        if (emptyStroke)
                        {
                            rows4deleting.Add(row);
                            Debug.WriteLine($"{row} => FAILED");
                        }
                        else
                        {
                            //Debug.WriteLine($"{row} => norm");
                        }
                    }
                    rows4deleting.Reverse();
                    foreach (int row in rows4deleting)
                    {
                        _ws.Rows[row].Delete(1);
                    }
                }

            }
            catch (Exception e)
            {
                Debug.WriteLine(e.StackTrace);
            }
            
        }
        public void getContentName(Compilation compilation)
        {
            Excel.Range findRange = _ws.Cells;
            Excel.Range rangeTSNName = findRange.Find("ТСН-2001.");
            if (rangeTSNName == null) Debug.WriteLine("Compilation.Parse() error : rangeTSNName == null");

            compilation.compilation_name = rangeTSNName.Value;
            Debug.WriteLine($"compilation_name = {rangeTSNName.Value}");
        }

        public void getContentRange( out (Range, Range) fromToRange)
        {
            string findText = "Техническая часть";
            Excel.Range rangeFrom = _ws.Cells.Find(findText);
            Excel.Range rangeTo = rangeFrom.Find(findText);
            fromToRange = (rangeFrom, rangeTo);
        }

        private void parseTables( Compilation compilation)
        {
            List<string> addressesOfContinuationTables = ContinuationOfTablesSearcher.getAddressesOfContinuationTables(_ws);
            Dictionary<string, TSNCommonTable> continuationTablesDict = new Dictionary<string, TSNCommonTable>(addressesOfContinuationTables.Count);
            if (addressesOfContinuationTables != null && addressesOfContinuationTables.Count > 0)
            {
                foreach (string address in addressesOfContinuationTables)
                {
                    Range rCode = _ws.Range[address].Find("Код");
                    TSNCommonTable table = ExcelTablesParser.extractCommonTable(rCode);
                    Match match = Regex.Match(_ws.Range[address].Value, TSNRegexPatterns.ShortTableNamePattern, RegexOptions.IgnoreCase );
                    if (match.Success) 
                    { 
                        if (continuationTablesDict.ContainsKey(match.Value))
                        {
                            continuationTablesDict[match.Value].Merge(table);
                        }
                        else
                        {
                            continuationTablesDict.Add(match.Value, table);
                        }
                    }
                }
            }
            foreach (Branch branch in compilation.branches)
            {
                foreach (Section section in branch.sections)
                {
                    foreach (KeyValuePair<string, string> kvTableLink in section.tableLinks)
                    {
                        Range secondRange = _ws.Range[kvTableLink.Value].Find( kvTableLink.Key );
                        Range findFromRange = _ws.Range[kvTableLink.Value];
                        Range r2 = _ws.Range[secondRange.Address].Find("Код");
                        TSNCommonTable table = ExcelTablesParser.extractCommonTable(r2);
                        Match match = Regex.Match(secondRange.Value, TSNRegexPatterns.ShortTableNamePattern, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            if(continuationTablesDict.ContainsKey(match.Value))
                            {
                                table.Merge(continuationTablesDict[match.Value]);
                                continuationTablesDict.Remove(match.Value);
                            }
                        }

                        section.AddTable(table);
                    }
                }
            }
        }
    }
}
