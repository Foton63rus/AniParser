using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace AniParser.Entity.TSN
{
    public class TSNCompilationExtractor
    {
        public static Excel.Application _excel = new Excel.Application();
        public Compilation compilation;

        private Branch currentBranch;
        private Section currentSection;

        public void Parse( string path)
        {
            Excel.Worksheet ws = OpenWorkBook(path);
            ExcelTablesParser.Init(_excel, ws);
            Compilation compilation = new Compilation();
            string numberInTheEndPattern = @"(\s+\d+)$";
            getContentName(ws, compilation);
            getContentRange(ws, out (Range, Range) fromToRange);
            string fromToAddress = $"{fromToRange.Item1.Address}:{fromToRange.Item2.Address}";
            Range contentRange = ws.Range[fromToAddress];
            for (int i = fromToRange.Item1.Row+1; i < fromToRange.Item2.Row; i++)
            {
                string address = "A" + i ;
                string strokeValue = $"{ws.Range[address].Value}".Trim();
                Match match = Regex.Match(strokeValue, numberInTheEndPattern, RegexOptions.IgnoreCase);

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
            parseTables(compilation, ws);
            Debug.WriteLine($"{JsonConvert.SerializeObject(compilation)}");
        }

        public static Excel.Worksheet OpenWorkBook(string path)
        {
            if (path == "" || path == null)
            {
                Debug.WriteLine($"path is Empty");
                return null;
            }
            Excel.Workbook _wb = _excel.Workbooks.Open(path);
            Excel.Worksheet _ws = (Excel.Worksheet)_wb.ActiveSheet;
            _excel.Visible = true;
            return _ws;
        }

        public void getContentName(Excel.Worksheet ws, Compilation compilation)
        {
            Excel.Range findRange = ws.Cells;
            Excel.Range rangeTSNName = findRange.Find("ТСН-2001.");
            if (rangeTSNName == null) Debug.WriteLine("Compilation.Parse() error : rangeTSNName == null");

            compilation.compilation_name = rangeTSNName.Value;
            Debug.WriteLine($"compilation_name = {rangeTSNName.Value}");
        }

        public void getContentRange( Excel.Worksheet ws, out (Range, Range) fromToRange)
        {
            string findText = "Техническая часть";
            Excel.Range rangeFrom = ws.Cells.Find(findText);
            Excel.Range rangeTo = rangeFrom.Find(findText);
            fromToRange = (rangeFrom, rangeTo);
        }

        private void parseTables( Compilation compilation, Worksheet ws)
        {
            foreach (Branch branch in compilation.branches)
            {
                foreach (Section section in branch.sections)
                {
                    foreach (KeyValuePair<string, string> kvTableLink in section.tableLinks)
                    {
                        Range secondRange = ws.Range[kvTableLink.Value].Find( kvTableLink.Key );
                        Range findFromRange = ws.Range[kvTableLink.Value];
                        Range r2 = ws.Range[secondRange.Address].Find("Код");
                        TSNCommonTable table = ExcelTablesParser.extractCommonTable(r2);
                        section.AddTable(table);
                    }
                }
            }
        }
    }
}
