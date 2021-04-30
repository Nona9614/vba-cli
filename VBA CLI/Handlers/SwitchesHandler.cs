using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Handlers
{
    public static class SwitchesHandler
    {
        private const string ONLY_EXECUTABLE_SWITCH = "/e";
        public static sbyte UsesExecutablePaths(ref List<string> parameters)
        {
            if (parameters == null) return -1;
            if (parameters.Contains(ONLY_EXECUTABLE_SWITCH))
            {
                Console.WriteLine("Using default files");
                parameters.Remove(ONLY_EXECUTABLE_SWITCH);
                return 1;
            }
            else return 0;
        }
    }
}
