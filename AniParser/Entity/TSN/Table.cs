using System;

namespace AniParser.Entity.TSN
{
    [Serializable]
    public class Table
    {
        public string table_name;
        public object table_conditions = null;

        private string address;

        public string GetAddress() => address;
    }
}
