using System;

namespace AniParser
{
    internal static class Debug
    {
        public static Action<string> DebugLogAction;
        public static Action DebugClearAction;
        public static void WriteLine(string line)
        {
            DebugLogAction(line);
        }
        public static void ConsoleClear(string line)
        {
            DebugClearAction();
        }
    }
}
