using System;
using System.Collections.Generic;
using System.Linq;
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
        public List<TSNCommonTableRecord> Records = new List<TSNCommonTableRecord>();

        private string measurePattern = @"^.*Измеритель:\s*(\d+)\s+(.+)";

        public void AddRecord( TSNCommonTableRecord newRecord)
        {
            Records.Add(newRecord);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (TSNCommonTableRecord record in Records)
            {
                sb.Append($"\n>  {record}");
            }
            return $"TSNCommonTable: name:{table_name} [{Records.Count}]{{{sb.ToString()}}}";
        }

        public TSNCommonTable(string name, string raw_measure) 
        {
            table_name = name;
            if (raw_measure != null)
            {
                Match match = Regex.Match(raw_measure, measurePattern, RegexOptions.IgnoreCase);
                this.table_quantity = match.Groups[1].Value;
                this.table_uom = match.Groups[2].Value;
            }
        }
    }
}
