using System.Collections.Generic;
using Excel = VBA.ExcelHandler;
using static VBA.Program;
using System;
using System.IO;
using System.Text.RegularExpressions;
using VBA.Handlers;

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
            if (parameters == null)
            {
                Console.WriteLine("This command needs parameters");
                return false;
            }
            bool result = false;
            switch (parameters[0])
            {
                case "customUI":
                    XMLHandler.GenerateDefaultCustomUI();
                    break;
                case "excel-file":
                    string name = null;
                    switch (parameters.Count)
                    {
                        case 2:
                            name = parameters[1];
                            name = Regex.Match(name, @"\w\.*").Length < 0 ? $"{name}.xlsm" : name;
                            name = Project.Paths.CheckForDefaultPath(name);
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters"); 
                            break;
                    }
                    if (name != null) 
                    {
                        result = Excel.CreateExcelFile(name); 
                    }
                    break;
                default:
                    Console.WriteLine($"Option '{parameters[0]}' is not valid");
                    result = false;
                    break;
            }
            return result;
        }
        // Disposes of resources 
        public void Dispose()
        {
            instance.Dispose();
        }

    }

}
