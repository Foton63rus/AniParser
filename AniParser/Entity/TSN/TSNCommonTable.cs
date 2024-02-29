using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AniParser.Entity
{
    [Serializable]
    public class TSNCommonTable
    {
        public string table_name;
        public string table_quantity;
        public string table_uom;
        public object table_conditions = null;
        public List<Paragraph> paragraphs = new List<Paragraph>();

        public void AddParagraph( Paragraph newParagraph)
        {
            paragraphs.Add(newParagraph);
        }

        public void Merge(TSNCommonTable mergedTable)
        {
            if (mergedTable.paragraphs.Count > 0 && mergedTable != null)
            {
                foreach (Paragraph paragraph in mergedTable.paragraphs)
                {
                    AddParagraph(paragraph);
                }
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (Paragraph record in paragraphs)
            {
                sb.Append($"\n>  {record}");
            }
            return $"TSNCommonTable: name:{table_name} [{paragraphs.Count}]{{{sb.ToString()}}}";
        }

        public TSNCommonTable(string name, string raw_measure) 
        {
            table_name = name;
            if (raw_measure != null)
            {
                Match match = Regex.Match(raw_measure, TSNRegexPatterns.MeasurePattern, RegexOptions.IgnoreCase);
                this.table_quantity = match.Groups[1].Value;
                this.table_uom = match.Groups[2].Value;
            }
        }
    }
}
