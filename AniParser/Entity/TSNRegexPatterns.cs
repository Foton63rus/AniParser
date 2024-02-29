namespace AniParser.Entity
{
    public static class TSNRegexPatterns
    {
        public const string ShortTableNamePattern = @"Таблица\s+\d+\-\d+";
        public const string NumberAtTheEndOfTheLinePattern = @"(\s+\d+)$";
        public const string MeasurePattern = @"^.*Измеритель:\s*(\d+)\s+(.+)";
    }
}
