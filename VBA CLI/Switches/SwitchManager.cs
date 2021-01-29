using System;
using System.Collections.Generic;
using System.Text;
using VBA.Switches;

namespace VBA
{
    public static class SwitchManager
    {
        public static bool SelectSwitch(string command, List<string> parameters)
        {
            ISwitch @switch = null;
            switch (command)
            {
                case "generate":
                case "g":
                    @switch = SwitchGenerate.Instance;
                    break;
                default:
                    Console.WriteLine($"Switch '{command}' not invalid");
                    break;
            }
            bool result = @switch != null;
            if (result) { result = @switch.Call(parameters); }
            return result;
        }

    }
}
