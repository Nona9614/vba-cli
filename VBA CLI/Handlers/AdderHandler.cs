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
            if (excel != null && !Excel.IsExcelFile(excel)) Excel.AddMacroEnabledExtension(ref excel);
            if (!Verify.Name(ref excel, ref excelPath)) return false;
            if (customUI != null) CustomUI.AddXmlExtension(ref customUI);
            if (!Verify.Name(ref customUI, ref customUIPath)) return false;
            if (!CustomUI.AddCustomUI($"{excelPath}\\{excel}", $"{customUIPath}\\{customUI}")) return false;
            return true;
        }
    }
}
