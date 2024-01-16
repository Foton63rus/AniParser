using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AniParser.Entity
{
    [Serializable]
    public class TSNCommonTableRecord
    {
        public string RecordType;
        [JsonIgnore]
        public List<string> Header;
        public string Code;
        public string Work;
        public string Dimension;
        public string Value;

        public string HeaderInLine => String.Join("_", Header);

        public override string ToString()
        {
            return $"type:{RecordType} header:{HeaderInLine} code:{Code} work:{Work} dimension:{Dimension} value:{Value}";
        }
    }
}
