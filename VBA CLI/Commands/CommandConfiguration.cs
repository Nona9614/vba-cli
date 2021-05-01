using System;
using System.Collections.Generic;
using System.Text;
using VBA.Handlers;
using VBA.Commands;

namespace VBA.Commands
{
    class CommandConfiguration : ICommand, IDisposable
    {
        private static CommandConfiguration instance;
        public static CommandConfiguration Instance
        {
            get
            {
                instance ??= new CommandConfiguration();
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
            if (ConfigurationFileHandler.FileExists()) ConfigurationFileHandler.LoadFile();
            else
            {
                if (!ConfigurationFileHandler.CreateConfigurationFile()) return false;
                ConfigurationFileHandler.SaveFile();
            }
            switch (parameters[0])
            {
                case "default-customUI":
                    switch (parameters.Count)
                    {
                        case 2:
                            if (ConfigurationFileHandler.SetCustomUIDefaultName(parameters[1])) ConfigurationFileHandler.SaveFile();
                            else return false;
                        break;
                        default:
                            Console.WriteLine("Not correct number of parameters");
                            break;
                    }
                    break;
                // Versionage as 'a.b.c.d'
                case "specify-version":
                    string versionage = null;
                    string release = null;
                    string feature = null;
                    string bugfix = null;
                    string optimization = null;
                    switch (parameters.Count)
                    {
                        case 2:
                            versionage = parameters[1];
                            // Specifying versionage 'd'
                            optimization = parameters[1];
                            break;
                        case 3:
                            // Specifying versionage 'c.d'
                            bugfix = parameters[2];
                            optimization = parameters[1];
                            break;
                        case 4:
                            // Specifying versionage 'b.c.d'
                            feature = parameters[3];
                            bugfix = parameters[2];
                            optimization = parameters[1];
                            break;
                        case 5:
                            // Specifying versionage 'a.b.c.d'
                            release = parameters[4];
                            feature = parameters[3];
                            bugfix = parameters[2];
                            optimization = parameters[1];
                            break;
                        default:
                            Console.WriteLine("Not valid number of parameters");
                            break;
                    }
                    if (ConfigurationFileHandler.SetVersion(versionage, true)) ConfigurationFileHandler.SaveFile();
                    else
                    {
                        if (ConfigurationFileHandler.SetVersion(release, feature, bugfix, optimization)) ConfigurationFileHandler.SaveFile();
                        else return false;
                    }
                    break;
                case "update-version":
                    // Updating versionage 'a.b.c.d' 
                    bool a = false;
                    bool b = false;
                    bool c = false;
                    bool d = false;
                    switch (parameters.Count)
                    {
                        // Defuault will only update versionage of 'd'
                        case 1:
                            d = true;
                            result = true;
                            break;
                        case 2:
                            switch (parameters[1])
                            {
                                case "release":
                                case "rls":
                                    a = true;
                                    break;
                                case "feature":
                                case "ftr":
                                    b = true;
                                    break;
                                case "bugfix":
                                case "bfx":
                                    c = true;
                                    break;
                                case "optimization":
                                case "opt":
                                    d = true;
                                    break;
                            }
                            result = true;
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    ConfigurationFileHandler.UpdateVersion(a, b, c, d);
                    ConfigurationFileHandler.SaveFile();
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
