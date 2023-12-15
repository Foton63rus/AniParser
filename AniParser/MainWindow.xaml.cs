using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using UglyToad.PdfPig;

namespace AniParser
{
    public partial class MainWindow : Window
    {
        string path = "";
        string pathOut = "";
        string pathJson = "tsn_regex.json";
        int pageFrom = 1;
        int pageTo = 1;

        Dictionary<string, string> regexTmp = new Dictionary<string, string>();
        public MainWindow()
        {
            InitializeComponent();
            Debug.DebugLogAction += ConsoleWriteLine;
            Debug.DebugClearAction += ConsoleClear;

            pageFrom = 1;
            pageTo = 1;

            loadRegex();
        }

        private void getPathToFileTSNIndexes(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                path = openFileDialog.FileName;
                var a = openFileDialog.FileName.Split('.');
                pathOut = a[0] + "_parse.csv";
                fileIndexes.Text = path;
            }
        }

        private void parse(object sender, RoutedEventArgs e)
        {
            if (path == "")
            {
                MessageBox.Show("Укажи путь до файла с индексами");
                return;
            }
            try
            {
                PdfDocument pdf = PdfDocument.Open(path);
                StringBuilder sb = new StringBuilder();
                for (int i = pageFrom; i <= pageTo; i++)
                {
                    Regex regex = new Regex(tbRegex.Text);
                    MatchCollection matches = regex.Matches(pdf.GetPage(i).Text);

                    foreach (Match match in matches)
                    {
                        sb.Append(match.Value);
                        sb.Append(Environment.NewLine);
                    }
                    File.WriteAllText(pathOut, sb.ToString());
                }
                MessageBox.Show($"Результат записан в {pathOut}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void preview(object sender, RoutedEventArgs e)
        {
            if (path == "")
            {
                MessageBox.Show("Укажи путь до файла с индексами");
                return;
            }
            try
            {
                PdfDocument pdf = PdfDocument.Open(path);
                var page = pdf.GetPage(pageFrom);
                //MessageBox.Show(page.Text);
                File.WriteAllText(pathOut, page.Text);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tbPageFrom_KeyUp(object sender, KeyEventArgs e)
        {
            int prev = pageFrom;
            if (tbPageFrom.Text == "")
            {
                pageFrom = 1;
            }
            if (int.TryParse(tbPageFrom.Text, out pageFrom))
            {
                if (pageFrom > pageTo) pageTo = pageFrom;
                if (pageFrom < 1) pageFrom = 1;
            }
            else
            {
                pageFrom = prev;
            }
            tbPageFrom.Text = pageFrom.ToString();
            tbPageTo.Text = pageTo.ToString();
            tbPageFrom.SelectionStart = pageFrom.ToString().Length;
        }

        private void tbPageTo_KeyUp(object sender, KeyEventArgs e)
        {
            int prev = pageTo;
            if (tbPageTo.Text == "")
            {
                pageTo = pageFrom;
            }
            if (int.TryParse(tbPageTo.Text, out pageTo))
            {
                if (pageTo < pageFrom) pageTo = pageFrom;
            }
            else
            {
                pageTo = prev;
            }
            tbPageTo.Text = pageTo.ToString();
            tbPageTo.SelectionStart = pageTo.ToString().Length;
        }

        private void loadRegex()
        {
            if (File.Exists(pathJson))
            {
                string jsontext = File.ReadAllText(pathJson);
                regexTmp = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsontext);
            }
            else
            {
                regexTmp.Add("", "");
                regexTmp.Add("№п/п Шифр Коэффициент", "(\\d{0,6}\\s?)(\\d+\\.\\d+-\\d+-\\d+\\s?)+(\\d+\\,\\d+)");
                regexTmp.Add("№п/п Шифр Коэффициент ЭМ МР", "(\\d+\\s)(\\d+\\.\\d+\\s?\\-\\s?\\d+\\s?\\-\\s?\\d+\\s){1,2}(\\-\\s?|\\d+\\,\\d+\\s?){2}");
            }
            cbTmp.ItemsSource = regexTmp.Keys;
        }

        private void CbTmp_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            tbRegex.Text = regexTmp[cbTmp.SelectedItem.ToString()];
        }

        private void Test(object sender, RoutedEventArgs e)
        {
            ExcelTablesParser.Parse(path);
        }

        void openWithOpenXML(string path)
        {
            //if (File.Exists(path))
            //{
            //    using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            //    {
            //        WorkbookPart workbookPart = document.WorkbookPart;
            //        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

            //        Sheet sheet = sheets.First();
            //        var a = sheet;

            //        WorkbookPart wbPart = document.WorkbookPart;
            //        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            //        Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == "A200").FirstOrDefault();
            //        MessageBox.Show($" {theCell.CellReference} {theCell.CellValue.InnerText}");
            //    }
            //}
        }

        private void ConsoleWriteLine(string str)
        {
            tbConsole.AppendText($"{str}\n");
        }
        private void ConsoleClear()
        {
            tbConsole.Document.Blocks.Clear();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ConsoleClear();
        }
    }
}
