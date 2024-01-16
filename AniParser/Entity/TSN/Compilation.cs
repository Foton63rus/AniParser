using System;
using System.Collections.Generic;

namespace AniParser.Entity.TSN
{
    [Serializable]
    public class Compilation //Сборник
    {
        public string compilation_name;
        public object compilation_conditions = null;
        public List<Branch> branches = new List<Branch>();

        public Branch AddBranch(Branch branch)
        {
            if (!branches.Contains(branch))
            {
                branches.Add(branch);
                return branch;
            }
            else
            {
                return null;
            }
        }
    }
}
