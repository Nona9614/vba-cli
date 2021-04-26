using System;
using System.Collections.Generic;
using System.Text;
using VBA.Handlers;
using VBA.Switches;

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
            ConfigurationFileHandler.CheckForFileExistence();
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
                    if (result) if (!ConfigurationFileHandler.SpecifyVersion(versionage)) ConfigurationFileHandler.SaveChanges();
                    break;
                case "update-version":
                    // Updating versionage 'a.b.c.d' 
                    switch (parameters.Count)
                    {
                        // Defuault will only update versionage of 'd'
                        case 1:
                            ConfigurationFileHandler.UpdateVersion(d: true);
                            result = true;
                            break;
                        case 2:
                            switch (parameters[1])
                            {
                                case "release":
                                case "rls":
                                    ConfigurationFileHandler.UpdateVersion(a: true);
                                    ConfigurationFileHandler.SaveChanges();
                                    result = true;
                                    break;
                                case "feature":
                                case "ftr":
                                    ConfigurationFileHandler.UpdateVersion(b: true);
                                    ConfigurationFileHandler.SaveChanges();
                                    result = true;
                                    break;
                                case "bugfix":
                                case "bfx":
                                    ConfigurationFileHandler.UpdateVersion(c: true);
                                    ConfigurationFileHandler.SaveChanges();
                                    result = true;
                                    break;
                                case "optimization":
                                case "opt":
                                    ConfigurationFileHandler.UpdateVersion(d: true);
                                    ConfigurationFileHandler.SaveChanges();
                                    result = true;
                                    break;
                            }
                            break;
                        default:
                            Console.WriteLine("Not recognized parameters");
                            break;
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
