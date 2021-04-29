using VBA.Project;
using VBA.Project.Files;
using Excel = VBA.Handlers.ExcelHandler;
using CustomUI = VBA.Handlers.CustomUIHandler;
using VBE = VBA.Executable.Files.VBE;

namespace VBA.Handlers
{
    public static class AdderHandler
    {
        public static bool AddCustomUI(string excel, string excelPath, string customUI, string customUIPath)
        {
            if(!Excel.IsExcelFile(excel)) Excel.AddMacroEnabledExtension(ref excel);
            Verify.Name(ref excel, ref excelPath);
            CustomUI.AddXmlExtension(ref customUI);
            Verify.Name(ref customUI, ref customUIPath);
            if (excel != null && excelPath != null && customUI != null && customUIPath != null)
            {
                if (!CustomUI.AddCustomUI($"{excelPath}\\{excel}", $"{customUIPath}\\{customUI}")) return false;
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
