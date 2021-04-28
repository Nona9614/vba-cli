using System;
using System.IO;
using System.Text.RegularExpressions;
using VBA.Project;
using Excel = VBA.ExcelHandler;

namespace VBA.Handlers
{
    public static class GeneratorHandler
    {
        public static void VerifyProjectName(ref string name, ref string _base)
        {
            bool isValidName = Regex.Match(name, @"^[\w\-\(\)\[\]\/\\]+$").Length > 0;
            bool isValidRoute = Directory.Exists(_base);
            if (isValidName && isValidRoute)
            {
                _base = Regex.Replace($"{_base}\\{name}", "[/]", "\\");
                name = $"{_base}".Split("\\")[^1];
            }
            else
            {
                if (!isValidName) Console.WriteLine($"Name '{name}' has invalid format");
                if (!isValidRoute) Console.WriteLine($"Route'{_base}' doesn't exists");
                name = null;
                _base = null;
            }
        }
        public static void VerifyProjectCustomUI(ref string customUI)
        {
            if (customUI != null)
            {
                if (Regex.Match(Path.GetExtension(customUI), @".*\.[xX][mM][lL]").Length <= 0) customUI += ".xml";
                if (!File.Exists(customUI)) customUI = $"{Paths.VBE.CustomUI}\\{customUI}";
            }
            else
            {
                customUI = Project.Files.VBE.CustomUI.Default;
            }
        }
        public static bool CreateExcelFile(string name, string _base, string customUI)
        {
            VerifyProjectName(ref name, ref _base);
            if (name != null && _base != null)
            {
                Paths.Base = _base;
                VerifyProjectCustomUI(ref customUI);
                Directory.CreateDirectory(Paths.Base);
                if (!Excel.CreateExcelFile($"{Paths.Base}\\{name}", customUI)) return false;
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool CreateProject(string name, string _base, string customUI)
        {
            VerifyProjectName(ref name, ref _base);
            if (name != null && _base != null)
            {
                Paths.Base = _base;
                VerifyProjectCustomUI(ref customUI);
                Directory.CreateDirectory(Paths.Base);
                Structure.CreateFolders();
                Structure.CreateDefaultFiles();
                if (!Excel.CreateExcelFile($"{Paths.Base}\\{name}", customUI)) return false;
                if (!ConfigurationFileHandler.CreatingConfigurationFile(name)) return false;
                Console.WriteLine("Project Ready!");
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
