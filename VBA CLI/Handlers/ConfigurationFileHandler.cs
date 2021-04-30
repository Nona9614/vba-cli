using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Newtonsoft;
using VBA.Models;
using System.Text.RegularExpressions;

namespace VBA.Handlers
{
    public static class ConfigurationFileHandler
    {
        private static string FileName
        {
            get { return $@"{Project.Paths.Base}/configuration.json"; }
        }

        private static ConfigurationFileModel _model;

        private static ConfigurationFileModel Model
        {
            get
            {
                if (_model == null)
                {
                    string fileContents = File.ReadAllText(FileName);
                    _model = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigurationFileModel>(fileContents);
                }
                return _model;
            }
        }

        public static void SaveChanges()
        {
            File.WriteAllText(FileName, Newtonsoft.Json.JsonConvert.SerializeObject(_model));
            _model = null;
            Console.WriteLine("Saving changes to configuration file");
        }

        public static bool CheckForFileExistence()
        {
            if (!File.Exists(FileName))
            {
                string project = null;
                int x = 0;
                ConsoleKeyInfo key;
                Console.WriteLine("\nConfiguration file not found, type the project name then press 'enter' to finish or 'esc' to cancel.");
                Console.Write("Project name --> ");
                int pos = Console.CursorLeft;
                while (!(x == 0xa || x == 0xd))
                {
                    key = Console.ReadKey();
                    x = key.KeyChar;
                    if (char.IsLetterOrDigit(key.KeyChar) || key.KeyChar == 0x28 || key.KeyChar == 0x29 || key.KeyChar == 0x2D || key.KeyChar >= 0x5F && key.KeyChar <= 0x5B) project += key.KeyChar;
                    if (x == 0x8) {
                        if (pos < Console.CursorLeft)
                        {
                            Console.Write(" \b");
                        }
                        else
                        {
                            Console.Write(" ");
                            Console.CursorLeft = pos;
                        }
                    };
                    if (x == 0x1b) 
                    {
                        Console.Write("Canceled\n");
                        return false;
                    }
                }
                if(!CreateConfigurationFile(project)) return false;
            }
            return true;
        }

        public static bool CreateConfigurationFile(string project)
        {
            File.WriteAllText(FileName, "{}");
            Model.Version = "0.0.0.0";
            if (!SetProjectName(project)) return false;
            if (!SetCustomUIDefaultName("customUI")) return false;
            Console.WriteLine("Configuration file created succesfully");
            SaveChanges();
            return true;
        }

        public static int[] VersionageStringToArrayInt(string value)
        {
            string[] versions = value.Split(".");
            if (versions.Length != 4) {
                Console.WriteLine($"Only 4 greater than zero integer values separated by a a dot '.' string is valid");
                return null;
            }

            int[] results = new int[4];

            for (int i = 0; i <= 3; i++)
            {
                if (!int.TryParse(versions[i], out results[i]))
                {
                    Console.WriteLine($"'{versions[i]}' is not a number, only greater than zero integer values are valid");
                    return null;
                }
                if (results[i] < 0)
                {
                    Console.WriteLine($"'{versions[i]}' is a negative value, only greater than zero integer values are valid");
                    return null;
                }
            }

            return results;
        }

        // Versionage as '0.0.0.0'
        public static void UpdateVersion(bool a = false, bool b = false, bool c = false, bool d = false)
        {
            int[] versions = Array.ConvertAll(Model.Version.Split("."), int.Parse);

            Model.Version =
                (versions[0] + (a ? 1 : 0)).ToString() + "." +
                (versions[1] + (b ? 1 : 0)).ToString() + "." +
                (versions[2] + (c ? 1 : 0)).ToString() + "." +
                (versions[3] + (d ? 1 : 0)).ToString();

            Console.WriteLine($"Setting project version {Model.Version}");
        }

        public static void SpecifyVersion(int? a = null, int? b = null, int? c = null, int? d = null)
        {
            int[] versions = Array.ConvertAll(Model.Version.Split("."), int.Parse);

            Model.Version =
                (a != null ? a : versions[0]).ToString() + "." +
                (b != null ? b : versions[1]).ToString() + "." +
                (c != null ? c : versions[2]).ToString() + "." +
                (d != null ? d : versions[3]).ToString();

            Console.WriteLine($"Setting project version {Model.Version}");
        }

        public static bool SpecifyVersion(string versionage)
        {
            int[] versions = VersionageStringToArrayInt(versionage);
            if (versions == null) return false;

            Model.Version = $"{versions[0]}.{versions[1]}.{versions[2]}.{versions[3]}";

            Console.WriteLine($"Setting project version {Model.Version}");
            return true;
        }
        public static bool IsValidProjectName(string name)
        {
            return Regex.Match(name, @"^[\w-]+$").Length > 0;
        }
        public static bool SetProjectName(string name)
        {
            if (IsValidProjectName(name))
            {
                Model.ProjectName = name;
                return true;
            }
            else
            {
                Console.WriteLine($"'{name}' is not a valid project name");
                return false;
            }
        }
        public static string GetProjectName()
        {
            if (!IsValidProjectName(Model.ProjectName))
            {
                Console.WriteLine($"'{Model.ProjectName}' is not a valid project name");
                return null;
            }
            else return Model.ProjectName;
        }
        public static bool SetCustomUIDefaultName(string name)
        {

            if (CustomUIHandler.IsXMLFile(name) && Project.Files.Verify.IsValidFileName(name))
            {
                Model.CustomUIDefaultName = name;
                return true;
            }
            else
            {
                Console.WriteLine($"'{name}' is not a customUI name");
                return false;
            }
        }
        public static string GetCustomUIDefaultName()
        {
            if (CustomUIHandler.IsXMLFile(Model.CustomUIDefaultName) && Project.Files.Verify.IsValidFileName(Model.CustomUIDefaultName)) return Model.CustomUIDefaultName;
            else
            {
                Console.WriteLine($"'{Model.CustomUIDefaultName}' is not a customUI name");
                return null;
            }
        }
    }
}
