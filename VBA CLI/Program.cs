using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using VBA_CLI.Switches;

namespace VBA_CLI
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
            int code = 0;
            string input = args.Length < 1 ? null : args[0].ToLower().Substring(0, 1);
            string command;

            switch (input)
            {
                case SWITCH_KEY:
                    command = args[0].ToLower()[1..args[0].Length];
                    List<string> paramaters = GetParameters(args);
                    code = SelectSwitch(command, paramaters);
                    break;
                case NO_INPUT:
                    Console.WriteLine("Welcome to VBA! Write '/help' for more information or try 'VBA generate project'.");
                    code = (int)ReturnCodes.NoInput;
                    break;
                default:
                    Console.WriteLine($"Input '{args[0]}' invalid");
                    code = (int)ReturnCodes.NonValidInput;
                    break;
            }

            return code;
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

        private static int SelectSwitch(string command, List<string> parameters)
        {
            int code;
            switch (command)
            {
                case "generate":
                case "g":
                    code = parameters != null ? SwitchGenerate.Call(parameters[0]) : (int)ReturnCodes.NonValidInput;
                    break;
                default:
                    code = (int)ReturnCodes.NonValidSwitch;
                    break;
            }
            return code;
        }
    }
}
