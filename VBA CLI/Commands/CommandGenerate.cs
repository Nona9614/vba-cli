using System.Collections.Generic;
using Excel = VBA.ExcelHandler;
using static VBA.Program;
using System;

namespace VBA.Switches
{
    public class CommandGenerate : ICommand, IDisposable
    {
        private static CommandGenerate instance;
        public static CommandGenerate Instance
        {
            get
            {
                instance ??= new CommandGenerate();
                return instance;
            }
        }

        public bool Call(List<string> parameters)
        {
            bool result;
            switch (parameters[0])
            {
                case "file":
                    // Checks if a directory value was set
                    result = Excel.CreateExcelFile(parameters[1], parameters.Count > 2 ? parameters[2] : null);
                    break;
                default:
                    Console.WriteLine($"Option '{parameters[0]}' is not valid");
                    result = false;
                    break;
            }
            return result;
        }

        public void Dispose()
        {
            instance.Dispose();
        }

    }

}
