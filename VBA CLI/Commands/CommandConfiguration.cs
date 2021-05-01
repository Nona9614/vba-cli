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
            if (!ConfigurationFileHandler.CheckForFileExistence()) return false;
            ConfigurationFileHandler.SaveFile();
            switch (parameters[0])
            {
                // Versionage as 'a.b.c.d'
                case "specify-version":
                    string versionage = "";
                    switch (parameters.Count)
                    {
                        case 2:
                            // Specifying versionage 'd'
                            versionage = $"0.0.0.{parameters[1]}";
                            result = true;
                            break;
                        case 3:
                            // Specifying versionage 'c.d'
                            versionage = $"0.0.{parameters[1]}.{parameters[2]}";
                            result = true;
                            break;
                        case 4:
                            // Specifying versionage 'b.c.d'
                            versionage = $"0.{parameters[1]}.{parameters[2]}.{parameters[3]}";
                            result = true;
                            break;
                        case 5:
                            // Specifying versionage 'a.b.c.d'
                            versionage = $"{parameters[1]}.{parameters[2]}.{parameters[3]}.{parameters[4]}";
                            result = true;
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    if (result) if (!ConfigurationFileHandler.SpecifyVersion(versionage)) ConfigurationFileHandler.SaveFile();
                    break;
                case "update-version":
                    // Updating versionage 'a.b.c.d' 
                    switch (parameters.Count)
                    {
                        // Defuault will only update versionage of 'd'
                        case 1:
                            ConfigurationFileHandler.UpdateVersion(d: true);
                            ConfigurationFileHandler.SaveFile();
                            result = true;
                            break;
                        case 2:
                            bool release = false;
                            bool feature = false;
                            bool bugfix = false;
                            bool optimization = false;
                            switch (parameters[1])
                            {
                                case "release":
                                case "rls":
                                    release = true;
                                    break;
                                case "feature":
                                case "ftr":
                                    feature = true;
                                    break;
                                case "bugfix":
                                case "bfx":
                                    bugfix = true;
                                    break;
                                case "optimization":
                                case "opt":
                                    optimization = true;
                                    break;
                            }
                            ConfigurationFileHandler.UpdateVersion(release, feature, bugfix, optimization);
                            ConfigurationFileHandler.SaveFile();
                            result = true;
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
                    }
                    ConfigurationFileHandler.UpdateVersion(d: true);
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
