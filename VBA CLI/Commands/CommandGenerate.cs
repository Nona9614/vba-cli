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
                        case 1:
                            break;
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
                            Console.WriteLine("Not valid number of parameters");
                            return false;
                    }
                    result = GeneratorHandler.CreateExcelFile(exe, ignore, excel, excelPath, customUI, customUIPath);
                    break;
                case "project":
                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 1:
                            break;
                        case 2:
                            project = parameters[1];
                            break;
                        case 3:
                            project = parameters[1];
                            projectPath = parameters[2];
                            break;
                        default:
                            Console.WriteLine("Not valid number of parameters");
                            return false;
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
