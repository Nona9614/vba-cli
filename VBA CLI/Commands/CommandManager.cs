using System;
using System.Collections.Generic;
using System.Text;
using VBA.Switches;

namespace VBA
{
    public static class CommandManager
    {
        public static bool SelectSwitch(string command, List<string> parameters)
        {
            ICommand @switch = null;
            switch (command)
            {
                case "generate":
                case "g":
                    @switch = SwitchGenerate.Instance;
                    break;
                case "add":
                case "a":
                    @switch = CommandAdd.Instance;
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
