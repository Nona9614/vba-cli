using System;
using System.Collections.Generic;

namespace VBA
{
    class Program
    {
        public enum ReturnCodes
        {
            NonValidSwitch = 1,
            SwitchFailed = 2,
            NonValidInput = 3,
            NoInput = 4,
            SwitchSucceed = 5
        }

        private const string SWITCH_KEY = "/";
        private const string NO_INPUT = null;

        static int Main(string[] args)
        {
            ReturnCodes code = 0;
            string input = args.Length < 1 ? null : args[0].ToLower().Substring(0, 1);

            switch (input)
            {
                case SWITCH_KEY:
                    string command = args[0].ToLower()[1..args[0].Length];
                    List<string> paramaters = GetParameters(args);
                    code = SwitchManager.SelectSwitch(command, paramaters) ? ReturnCodes.SwitchSucceed : ReturnCodes.SwitchFailed;
                    break;
                case NO_INPUT:
                    Console.WriteLine("Welcome to VBA! Write '/help' for more information or try 'VBA generate project'.");
                    code = ReturnCodes.NoInput;
                    break;
                default:
                    Console.WriteLine($"Input '{args[0]}' invalid");
                    code = ReturnCodes.NonValidInput;
                    break;
            }
    
            return (int)code;
        }

        private static List<string> GetParameters(string[] args)
        {
            List<string> paramaters = null;
            if (args.Length > 1)
            {
                paramaters = new List<string>();
                for (int i = 1; i < args.Length; i++)
                {
                    paramaters.Add(args[i]);
                }
            }
            return paramaters;
        }

    }
}
