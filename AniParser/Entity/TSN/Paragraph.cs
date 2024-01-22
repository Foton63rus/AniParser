using AniParser.Entity.TSN;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AniParser.Entity
{
    [Serializable]
    public class Paragraph
    {
        public string paragraph_name;
        public List<string> paragraph_conditions = new List<string>();
        public List<Expense> expenses = new List<Expense>();
        [JsonIgnore]
        public string conditionsInLine => String.Join("_", paragraph_conditions);

        public override string ToString()
        {
            return $"type:{paragraph_name} header:{conditionsInLine} expenses count:{expenses.Count}";
        }
    }
}
