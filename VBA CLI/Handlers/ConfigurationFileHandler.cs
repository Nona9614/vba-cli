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
            get { return $@"{Project.Paths.Base}\configuration.json"; }
        }

        private static ConfigurationFileModel _model;

        private static ConfigurationFileModel Model
        {
            get
            {
                if (_model == null)
                {
                    _model = new ConfigurationFileModel();
                }
                return _model;
            }
        }
        public static void LoadFile()
        {
            string fileContents = File.ReadAllText(FileName);
            _model = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigurationFileModel>(fileContents);
            Console.WriteLine($"Loading configuration file from: '{FileName}'");
        }

        public static void SaveFile()
        {
            File.WriteAllText(FileName, Newtonsoft.Json.JsonConvert.SerializeObject(_model));
            Console.WriteLine($"Saving changes to configuration file in: '{FileName}'");
        }

        public static bool FileExists() => File.Exists(FileName);

        public static bool CheckForFileExistence()
        {
            if (!File.Exists(FileName))
            {
                Console.WriteLine("\nConfiguration file not found");
                if(!SetProjectNameByConsole() && !CreateConfigurationFile(Model.ProjectName)) return false;
            }
            return true;
        }
        public static bool CreateConfigurationFile() => CreateConfigurationFileInternal(true);
        public static bool CreateConfigurationFile(string project) => CreateConfigurationFileInternal(false, project);
        private static bool CreateConfigurationFileInternal(bool useConsoleForSettingName, string project = null)
        {
            Model.Version = "0.0.0.0";
            if (useConsoleForSettingName)
            { 
                if (!SetProjectNameByConsole()) return false; 
            }
            else 
            { 
                if (!SetProjectName(project)) return false; 
            }
            if (!SetCustomUIDefaultName("customUI.xml")) return false;
            File.WriteAllText(FileName, "{}");
            Console.WriteLine("Configuration file created succesfully");
            return true;
        }

        public static int[] VersionageStringToArrayInt(string value, bool ignoreMessages = false)
        {
            string[] versions = value.Split(".");
            if (versions.Length != 4) {
                if (!ignoreMessages) Console.WriteLine($"Only 4 greater than zero integer values separated by a a dot '.' string is valid");
                return null;
            }

            int[] results = new int[4];

            for (int i = 0; i <= 3; i++)
            {
                if (!int.TryParse(versions[i], out results[i]))
                {
                    if (!ignoreMessages) Console.WriteLine($"'{versions[i]}' is not a number, only greater than zero integer values are valid");
                    return null;
                }
                if (results[i] < 0)
                {
                    if (!ignoreMessages) Console.WriteLine($"'{versions[i]}' is a negative value, only greater than zero integer values are valid");
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
        private static int IsValidVersionNumber(string number, string of)
        {
            if (!int.TryParse(number, out int result))
            {
                Console.WriteLine($"The {of} number '{number}' is not valid, only greater than zero integer values are valid");
                return -1;
            }
            if (result < 0)
            {
                Console.WriteLine($"The {of} number '{number}' is a negative value, only greater than zero integer values are valid");
                return -1;
            }
            return result;
        }
        public static bool SetVersion(string versionage, bool ignoreMessages = false)
        {
            int[] _versionage = VersionageStringToArrayInt(versionage, ignoreMessages);
            if (versionage == null) return false;

            int _release =_versionage[0];
            int _feature =_versionage[1];
            int _bugfix = _versionage[2];
            int _optimization = _versionage[3];

            if (_release < 0 || _feature < 0 || _bugfix < 0 || _optimization < 0) return false;
            else Model.Version = $"{_release}.{_feature}.{_bugfix}.{_optimization}";
            Console.WriteLine($"Setting project version {Model.Version}");
            return true;

        }
        public static bool SetVersion(string release = null, string feature = null, string bugfix = null, string optimization = null)
        {
            int[] versionage = VersionageStringToArrayInt(Model.Version);
            if (versionage == null) return false;

            int _release = release == null ? versionage[0] : IsValidVersionNumber(release, "release");
            int _feature = feature == null ? versionage[1] : IsValidVersionNumber(feature, "feature");
            int _bugfix = bugfix == null ? versionage[2] : IsValidVersionNumber(bugfix, "bugfix");
            int _optimization = optimization == null ? versionage[3] : IsValidVersionNumber(optimization, "optimization");

            if ( _release < 0 || _feature < 0 || _bugfix < 0 || _optimization < 0 ) return false;
            else Model.Version = $"{_release}.{_feature}.{_bugfix}.{_optimization}";
            Console.WriteLine($"Setting project version {Model.Version}");
            return true;
        }

        public static bool IsValidProjectName(string name) => Regex.Match(name, @"^[\w-]+$").Success;
        private static bool IsNullContent(string key, string value)
        {
            if (value == null)
            {
                Console.WriteLine($"The configuration file does not have a '{key}' property");
                return true;
            }
            else return false;
        }
        public static bool SetProjectName(string name)
        {
            if (name != null && IsValidProjectName(name))
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
        public static bool SetProjectNameByConsole()
        {
            int x = 0;
            string project = null;
            ConsoleKeyInfo key;
            Console.WriteLine("Type the project name then press 'enter' to finish or 'esc' to cancel.");
            Console.Write("Project name --> ");
            int pos = Console.CursorLeft;
            while (!(x == 0xa || x == 0xd))
            {
                key = Console.ReadKey();
                x = key.KeyChar;
                if (char.IsLetterOrDigit(key.KeyChar) || key.KeyChar == 0x28 || key.KeyChar == 0x29 || key.KeyChar == 0x2D || key.KeyChar >= 0x5F && key.KeyChar <= 0x5B) project += key.KeyChar;
                if (x == 0x8)
                {
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
            return SetProjectName(project);
        }
        public static bool GetProjectName(ref string name)
        {
            name = Model.ProjectName;
            if (IsNullContent("ProjectName", name)) return false;
            if (!IsValidProjectName(name))
            {
                Console.WriteLine($"'{Model.ProjectName}' is not a valid project name");
                return false;
            }
            else return true;
        }
        public static bool SetCustomUIDefaultName(string name)
        {

            if (name != null && CustomUIHandler.IsXMLFile(name) && Project.Files.Verify.IsValidFileName(name))
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
        public static bool GetCustomUIDefaultName(ref string name)
        {
            name = Model.CustomUIDefaultName;
            if (IsNullContent("CustomUIDefaultName", name)) return false;
            if (!CustomUIHandler.IsXMLFile(name) && !Project.Files.Verify.IsValidFileName(name))
            {
                Console.WriteLine($"'{Model.CustomUIDefaultName}' is not a customUI name");
                return false;
            }
            else return true;
        }
    }
}
