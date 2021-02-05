using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
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
            if (parameters == null)
            {
                Console.WriteLine("This command needs parameters");
                return false;
            }
            bool result = false;
            switch (parameters[0])
            {
                case "customUI":
                    string customUI = null;
                    string excel = null;
                    bool fileNotFound = false;
                    switch (parameters.Count)
                    {
                        case 2:
                            // CustomUI file name was not set, default will be used
                            customUI = @$"{Project.Files.CustomUI}";
                            excel = SetDefaultPath(parameters[1], $"{Directory.GetCurrentDirectory()}");
                            break;
                        case 3:
                            customUI = SetDefaultPath(parameters[2], $@"{Directory.GetCurrentDirectory()}\resources");
                            excel = SetDefaultPath(parameters[1], Directory.GetCurrentDirectory());
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    if (!fileNotFound)
                    {
                        result = Excel.AddCustomUI(excel, customUI);
                    }
                    break;
                default:
                    Console.WriteLine($"Option '{parameters[0]}' is not valid");
                    break;
            }
            return result;
        }
        // If it is a valid file name with no path, a default will be used
        private static string SetDefaultPath(string name, string path)
        {
            // Checks if the file name comes alone or with a path
            if (File.Exists(name))
            {
                return name;
            }
            else
            {
                return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 ? $@"{path}\{name}" : null;
            }
        }
        public void Dispose()
        {
            instance.Dispose();
        }
    }
}
