using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using UglyToad.PdfPig;

namespace AniParser
{
    public class PDFTablesOfIndexesParser
    {
        public string Parse(string path, string regEx, int pageFrom = 1, int pageTo = 1)
        {
            if (path == "")
            {
                MessageBox.Show("Укажи путь до файла с индексами");
                return null;
            }
            try
            {
                PdfDocument pdf = PdfDocument.Open(path);
                StringBuilder sb = new StringBuilder();
                for (int i = pageFrom; i <= pageTo; i++)
                {
                    Regex regex = new Regex(regEx);
                    MatchCollection matches = regex.Matches(pdf.GetPage(i).Text);

                    foreach (Match match in matches)
                    {
                        sb.Append(match.Value);
                        sb.Append(Environment.NewLine);
                    }
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }
}
