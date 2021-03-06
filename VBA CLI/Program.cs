﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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

        private enum InputType
        {
            Switch = 1,
            Empty = 2,
            Invalid = 3
        }

        static int Main(string[] args)
        {
            ReturnCodes code = 0;
            InputType type = InputType.Empty;
            string input = args.Length < 1 ? null : args[0].ToLower().Substring(0, 1);
            if (input != null) { type = Regex.Match(input, "[a-zA-Z]").Success ? InputType.Switch : InputType.Invalid; }

            switch (type)
            {
                case InputType.Switch:
                    string command = args[0];
                    List<string> paramaters = GetParameters(args);
                    code = CommandManager.SelectSwitch(command, paramaters) ? ReturnCodes.SwitchSucceed : ReturnCodes.SwitchFailed;
                    break;
                case InputType.Empty:
                    Console.WriteLine("Welcome to VBA! Write 'help' for more information or try 'VBA generate project'.");
                    code = ReturnCodes.NoInput;
                    break;
                case InputType.Invalid:
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
