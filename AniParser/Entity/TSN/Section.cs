using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace AniParser.Entity.TSN
{
    [Serializable]
    public class Section // Раздел
    {
        public string section_name;
        public object section_conditions = null;
        public List<TSNCommonTable> tables = new List<TSNCommonTable>();

        private string address;
        [JsonIgnore]
        public Dictionary<string, string> tableLinks = new Dictionary<string, string>();

        [JsonIgnore]
        public string Address
        {
            get { return address; }
            set { address = value; }
        }

        public void addTableLink(string table_name, string address)
        {
            tableLinks[table_name] = address;
        }

        public void AddTable( TSNCommonTable table)
        {
            if (!tables.Contains(table))
            {
                tables.Add(table);
            }
        }
    }
}