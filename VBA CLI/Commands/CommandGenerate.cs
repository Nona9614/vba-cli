using System.Collections.Generic;
using System;
using VBA.Handlers;
using VBA.Project;

namespace VBA.Commands
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
            string project = null;
            string projectPath = null;
            switch (parameters[0])
            {
                case "customUI":
                    CustomUIHandler.GenerateDefaultCustomUI();
                    break;
                case "excel-file":
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
                    result = GeneratorHandler.CreateExcelFile(excel, excelPath, customUI, customUIPath);
                    break;
                case "project":
                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 1:
                            if (ConfigurationFileHandler.CheckForFileExistence()) project = ConfigurationFileHandler.GetProjectName();
                            projectPath = Paths.Base;
                            break;
                        case 2:
                            project = parameters[1];
                            projectPath = Paths.Base;
                            break;
                        case 3:
                            project = parameters[1];
                            projectPath = parameters[2];
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    result = GeneratorHandler.CreateProject(project, projectPath);
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
