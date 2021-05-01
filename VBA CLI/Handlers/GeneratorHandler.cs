using System;
using System.IO;
using VBA.Project;
using VBA.Project.Files;
using Excel = VBA.Handlers.ExcelHandler;
using CustomUI = VBA.Handlers.CustomUIHandler;
using Configuration = VBA.Handlers.ConfigurationFileHandler;
using VBE = VBA.Executable.Files.VBE;

namespace VBA.Handlers
{
    public static class GeneratorHandler
    {
        public static bool CreateExcelFile(bool exe, bool ignore, string excel, string excelPath, string customUI, string customUIPath)
        {
            if (exe || ignore)
            {
                // If not name is asigned, excel will try to create 'project.xlsm' file
                excel ??= Excel.DefaultExcelFileName;
                customUI ??= CustomUI.FileDefaultName;
                excelPath ??= Project.Paths.Base;
                customUIPath ??= exe && !ignore ? Executable.Paths.VBE.CustomUI : Project.Paths.VBE.CustomUI;
                Excel.AddMacroEnabledExtension(ref excel);
                if (!Verify.Name(ref excel, ref excelPath)) return false;
            }
            else
            {
                customUIPath ??= Paths.VBE.CustomUI;
                excelPath ??= Paths.Base;
                if (Configuration.FileExists())
                {
                    Configuration.LoadFile();
                    if (excel == null)
                    {
                        if (!Configuration.GetProjectName(ref excel)) return false;
                        Excel.AddMacroEnabledExtension(ref excel);
                    }
                    if (customUI == null)
                    {
                        if (!Configuration.GetCustomUIDefaultName(ref customUI)) return false;
                    }
                    if (excel != null) Excel.AddMacroEnabledExtension(ref excel);
                    if (!Verify.Name(ref excel, ref excelPath)) return false;
                    Directory.CreateDirectory(excelPath);
                }
                else
                {
                    customUI ??= CustomUI.FileDefaultName;
                    if (!Configuration.SetCustomUIDefaultName(customUI)) return false;
                    if (excel == null)
                    {
                        if (!Configuration.SetProjectNameByConsole()) return false;
                        if (!Configuration.GetProjectName(ref excel)) return false;
                    }
                    else
                    {
                        if (!Configuration.SetProjectName(excel)) return false;
                    }
                    Excel.AddMacroEnabledExtension(ref excel);
                    if (!Verify.Name(ref excel, ref excelPath)) return false;
                    excelPath += "\\" + excel;
                    Paths.Base = excelPath;
                    Directory.CreateDirectory(excelPath);
                    if (!Configuration.CreateConfigurationFile(excel)) return false;
                    Configuration.SaveFile();
                }   
            }
            if (customUI != null) CustomUI.AddXmlExtension(ref customUI);
            if (!Verify.Name(ref customUI, ref customUIPath)) return false;
            if (!CustomUI.FileExists($"{customUIPath}\\{customUI}")) return false;
            if (!Excel.CreateExcelFile($"{excelPath}\\{excel}", $"{customUIPath}\\{customUI}")) return false;
            return true;
        }
        public static bool CreateProject(string project, string projectPath)
        {
            if (projectPath == null) projectPath = Paths.Base;
            if (project == null)
            {
                if (!Configuration.SetProjectNameByConsole()) return false;
                if (!Configuration.GetProjectName(ref project)) return false;
            }
            else
            {
                if (!Configuration.SetProjectName(project)) return false;
            }
            if (!Verify.Name(ref project, ref projectPath)) return false;
            projectPath += "\\" + project;
            Paths.Base = projectPath;
            Directory.CreateDirectory(projectPath);
            if (!Configuration.CreateConfigurationFile(project)) return false;
            Configuration.SaveFile();
            string excel = project;
            Excel.AddMacroEnabledExtension(ref excel);
            Structure.CreateFolders();
            Structure.CreateDefaultFiles();
            if (!Excel.CreateExcelFile($"{projectPath}\\{excel}", VBE.CustomUI.Default)) return false;
            Console.WriteLine("Project Ready!");
            return true;
        }
    }
}
