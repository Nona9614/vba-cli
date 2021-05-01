using VBA.Project;
using VBA.Project.Files;
using Excel = VBA.Handlers.ExcelHandler;
using CustomUI = VBA.Handlers.CustomUIHandler;
using Configuration = VBA.Handlers.ConfigurationFileHandler;
using VBE = VBA.Executable.Files.VBE;
using System.IO;
using System;

namespace VBA.Handlers
{
    public static class AdderHandler
    {
        public static bool AddCustomUI(bool exe, bool ignore, string excel, string excelPath, string customUI, string customUIPath)
        {
            if (exe || ignore)
            {
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
                    }
                    Excel.AddMacroEnabledExtension(ref excel);
                    if (customUI == null)
                    {
                        if (!Configuration.GetCustomUIDefaultName(ref customUI)) return false;
                    }
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
                    Paths.Base = excelPath;
                    Directory.CreateDirectory(excelPath);
                    if (!Configuration.CreateConfigurationFile(excel)) return false;
                    Configuration.SaveFile();
                }
            }
            if (customUI != null) CustomUI.AddXmlExtension(ref customUI);
            if (!Verify.Name(ref customUI, ref customUIPath)) return false;
            if (!CustomUI.FileExists($"{customUIPath}\\{customUI}")) return false;
            if (!Excel.FileExists($"{excelPath}\\{excel}")) return false;
            if (!CustomUI.AddCustomUI($"{excelPath}\\{excel}", $"{customUIPath}\\{customUI}")) return false;
            return true;
        }
    }
}
