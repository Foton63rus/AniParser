using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace AniParser.Entity
{
    public static class ContinuationOfTablesSearcher
    {
        private static Match match = null;
        private static Range rangeFrom;
        private static List<string> addressList = new List<string>();
        public static List<string> getAddressesOfContinuationTables(Worksheet ws)
        {
            rangeFrom = ws.Cells.Find("продолжение");
            findAllContinuations();
            return addressList;
        }
        private static void findAllContinuations()
        {
            if (rangeFrom == null)
            {
                Debug.WriteLine($"Поиск не дал результатов");
                return;
            }
            else
            {
                match = Regex.Match(rangeFrom.Value as string, TSNRegexPatterns.ShortTableNamePattern, RegexOptions.IgnoreCase);
                if (match == null)
                {
                    return;
                }
                else
                {
                    if (addressList.Contains(rangeFrom.Address))
                    {
                        return;
                    }
                    else
                    {
                        addressList.Add(rangeFrom.Address);
                        rangeFrom = rangeFrom.Find("продолжение");
                        findAllContinuations();
                    }
                }
            }
        }
    }
}
