using System;
using System.Collections.Generic;
using VBA.Handlers;
using VBA.Project;
using Configuration = VBA.Handlers.ConfigurationFileHandler;

namespace VBA.Commands
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
                    string excel = null;
                    string excelPath = null;
                    string customUI = null;
                    string customUIPath = null;
                    bool exe = SwitchesHandler.UsesExecutablePaths(ref parameters) == 1;
                    bool ignore = SwitchesHandler.UsesIgnoreConfiguration(ref parameters) == 1;

                    if (exe && ignore)
                    {
                        Console.WriteLine("Only one switch can be used between '/e' and '/ic'");
                        return false;
                    }

                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 1:
                            break;
                        case 2:
                            excel = parameters[1];
                            break;
                        case 3:
                            excel = parameters[1];
                            customUI = parameters[2];
                            break;
                        case 4:
                            excel = parameters[1];
                            excelPath = parameters[2];
                            customUI = parameters[3];
                            break;
                        case 5:
                            excel = parameters[1];
                            excelPath = parameters[2];
                            customUI = parameters[3];
                            customUIPath = parameters[4];
                            break;
                        default:
                            Console.WriteLine("Not valid number of parameters");
                            break;
                    }
                    result = AdderHandler.AddCustomUI(exe, ignore, excel, excelPath, customUI, customUIPath);
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
