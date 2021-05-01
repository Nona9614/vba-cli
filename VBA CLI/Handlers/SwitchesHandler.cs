using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Handlers
{
    public static class SwitchesHandler
    {
        private const string ONLY_EXECUTABLE_SWITCH = "/e";
        private const string IGNORE_CONFIGURATION_SWITCH = "/ic";
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
        public static sbyte UsesIgnoreConfiguration(ref List<string> parameters)
        {
            if (parameters == null) return -1;
            if (parameters.Contains(IGNORE_CONFIGURATION_SWITCH))
            {
                Console.WriteLine("Using project route as default");
                parameters.Remove(IGNORE_CONFIGURATION_SWITCH);
                return 1;
            }
            else return 0;
        }

    }
}
