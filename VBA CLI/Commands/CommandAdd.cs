using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = VBA.ExcelHandler;

namespace VBA.Switches
{
    class CommandAdd: ICommand, IDisposable
    {
        private static CommandAdd instance;
        public static CommandAdd Instance
        {
            get
            {
                instance ??= new CommandAdd();
                return instance;
            }
        }

        public bool Call(List<string> parameters)
        {
            bool result;
            switch (parameters[0])
            {
                case "customUI":
                    // Checks if a directory value was set
                    if (parameters.Count == 3)
                    {
                        result = Excel.AddCustomUI(parameters[1], parameters[2], true);
                    }
                    else
                    {
                        Console.WriteLine("Not enough parameters");
                        result = false;
                    }
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
