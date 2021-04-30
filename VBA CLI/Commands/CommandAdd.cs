using System;
using System.Collections.Generic;
using VBA.Handlers;
using VBA.Project;

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
                    string excel;
                    string excelPath;
                    string customUI;
                    string customUIPath;
                    if (SwitchesHandler.UsesExecutablePaths(ref parameters) == 0)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        excel = "project.xslm";
                        customUI = "customUI.xml";
                        excelPath = Executable.Paths.Base;
                        customUIPath = Executable.Paths.VBE.CustomUI;
                    }
                    else
                    {
                        if (!ConfigurationFileHandler.CheckForFileExistence()) return false;
                        excel = ConfigurationFileHandler.GetProjectName();
                        excelPath = Paths.Base;
                        if (parameters.Count > 3) ConfigurationFileHandler.SetCustomUIDefaultName(parameters[3]);
                        customUI = ConfigurationFileHandler.GetCustomUIDefaultName();
                        customUIPath = Paths.VBE.CustomUI;
                    }
                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 2:
                            excel = parameters[1];
                            break;
                        case 3:
                            excel = parameters[1];
                            excelPath = parameters[2];
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
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    result = AdderHandler.AddCustomUI(excel, excelPath, customUI, customUIPath);
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
