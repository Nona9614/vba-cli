using System;
using System.Collections.Generic;
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
                        case 1:
                            // CustomUI file name was not set, default will be used
                            // Excel file name was not set, will search for <project>.xlsm
                            customUI = @$"{Project.Files.VBE.CustomUI.Default}";
                            excel = Project.Paths.CheckForDefaultPath("project");
                            break;
                        case 2:
                            // CustomUI file name was not set, default will be used
                            customUI = @$"{Project.Files.VBE.CustomUI.Default}";
                            excel = Project.Paths.CheckForDefaultPath(parameters[1]);
                            break;
                        case 3:
                            customUI = Project.Paths.CheckForDefaultPath(parameters[2], $@"{Project.Paths.Resources}");
                            excel = Project.Paths.CheckForDefaultPath(parameters[1]);
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
        // Release all resources
        public void Dispose()
        {
            instance.Dispose();
        }
    }
}
