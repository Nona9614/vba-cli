using System;
using System.IO;
using VBA.Project;
using VBA.Project.Files;
using Excel = VBA.Handlers.ExcelHandler;
using CustomUI = VBA.Handlers.CustomUIHandler;
using Configuration = VBA.Handlers.ConfigurationFileHandler;
using VBE = VBA.Executable.Files.VBE;
using System.Collections.Generic;

namespace VBA.Handlers
{
    public static class GeneratorHandler
    {
        public static bool CreateExcelFile( string excel, string excelPath, string customUI, string customUIPath)
        {
            if (excel!=null) Excel.AddMacroEnabledExtension(ref excel);
            if (!Verify.Name(ref excel, ref excelPath)) return false;
            if (customUI != null) CustomUI.AddXmlExtension(ref customUI);
            if (!Verify.Name(ref customUI, ref customUIPath)) return false;
            if (excel == null || excelPath == null || customUI == null || customUIPath == null) return false;
            Directory.CreateDirectory(excelPath);
            if (!Excel.CreateExcelFile($"{excelPath}\\{excel}", $"{customUIPath}\\{customUI}")) return false;
            return true;
        }
        public static bool CreateProject(string project, string projectPath)
        {
            if (!Verify.Name(ref project, ref projectPath)) return false;
            projectPath += "\\" + project;
            Paths.Base = projectPath;
            Directory.CreateDirectory(projectPath);
            if (!Configuration.CreateConfigurationFile(project)) return false;
            string excel = project;
            Excel.AddMacroEnabledExtension(ref excel);
            Directory.CreateDirectory(projectPath);
            Structure.CreateFolders();
            Structure.CreateDefaultFiles();
            if (!Excel.CreateExcelFile($"{projectPath}\\{excel}", VBE.CustomUI.Default)) return false;
            Console.WriteLine("Project Ready!");
            return true;
        }
    }
}
