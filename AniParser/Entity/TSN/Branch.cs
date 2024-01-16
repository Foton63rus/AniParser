using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AniParser.Entity.TSN
{
    [Serializable]
    public class Branch // Отдел
    {
        public string branch_name;
        public object branch_conditions = null;
        public List<Section> sections = new List<Section>();

        private string address;

        public Section AddSection( Section section)
        {
            if (!sections.Contains(section))
            {
                sections.Add(section);
                return section;
            }
            else
            {
                return null;
            }
        }
        [JsonIgnore]
        public string Address
        {
            get { return address; }
            set { address = value; }
        }
    }
}

