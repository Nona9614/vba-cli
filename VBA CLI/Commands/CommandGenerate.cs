using System.Collections.Generic;
using System;
using VBA.Handlers;
using VBA.Project;

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
            string name = null;
            string _base = null;
            string customUI = null;
            switch (parameters[0])
            {
                case "customUI":
                    XMLHandler.GenerateDefaultCustomUI();
                    break;
                case "excel-file":
                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 1:
                            name = "project";
                            _base = Paths.Base;
                            break;
                        case 2:
                            name = parameters[1];
                            _base = Paths.Base;
                            break;
                        case 3:
                            name = parameters[1];
                            _base = parameters[2];
                            break;
                        case 4:
                            name = parameters[1];
                            _base = parameters[2];
                            customUI = parameters[3];
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    result = GeneratorHandler.CreateExcelFile(name, _base, customUI);
                    break;
                case "project":
                    switch (parameters.Count)
                    {
                        // If not name is asigned, excel will try to create 'project.xlsm' file
                        case 1:
                            name = "project";
                            _base = Paths.Base;
                            break;
                        case 2:
                            name = parameters[1];
                            _base = Paths.Base;
                            break;
                        case 3:
                            name = parameters[1];
                            _base = parameters[2];
                            break;
                        case 4:
                            name = parameters[1];
                            _base = parameters[2];
                            customUI = parameters[3];
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters"); 
                            break;
                    }
                    result = GeneratorHandler.CreateProject(name, _base, customUI);
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
