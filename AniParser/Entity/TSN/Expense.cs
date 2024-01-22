using System;

namespace AniParser.Entity.TSN
{
    [Serializable]
    public class Expense
    {
        public string expense_name;
        public string expense_value;
        public string expense_uom;

        public override string ToString()
        {
            return $"expense_name: {expense_name}, expense_value: {expense_value}, expense_uom: {expense_uom}";
        }
    }
}
