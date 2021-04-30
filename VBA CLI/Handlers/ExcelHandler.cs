using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using VBE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using CustomUI = VBA.Handlers.CustomUIHandler;

namespace VBA.Handlers
{
    public static class ExcelHandler
    {
        private static readonly List<string> xlMacroEnabledExtensions = new List<string> { "xlsm", "xltm", "xlsb", "xla", "xlam", "xll" };
        public static bool IsExcelFile(string name)
        {
            return Regex.Match(Path.GetExtension(name), @".*\.[xX][lL]*").Length > 0;
        }
        public static bool IsMacroEnabledExcelFile(string name)
        {
            return xlMacroEnabledExtensions.Contains(Path.GetExtension(name));
        }
        // If has not valid macro-enabled extensions adds one
        public static void AddMacroEnabledExtension(ref string name)
        {
            if (!IsMacroEnabledExcelFile(name)) name += ".xlsm";
        }
        //  Returns true if the creation succeed, if not returns false
        public static bool CreateExcelFile(string name, string customUI = null)
        {
            // Validate overriding
            if (File.Exists(name))
            {
                Console.Write("There is an excel file already created with this name and path. \nWould you like to override? (y/n) --> ");
                ConsoleKeyInfo key;
                int x = 0x0;
                int pos = Console.CursorLeft;
                bool canceled = true;
                while (x != 0x1B)
                {
                    key = Console.ReadKey();
                    x = key.KeyChar;
                    if (!(x == 0x59 || x == 0x79 || x == 0x4E || x == 0x6E))
                    {
                        if (x != 0x1B) 
                        {
                            if (x != 0x8) { Console.Write("\b \b"); } else { Console.Write("  "); }
                            Console.CursorLeft = pos;
                        }
                        else
                        {
                            Console.Write("xn");
                        }
                    }
                    else
                    {
                        canceled = x == 0x4E || x == 0x6E;
                        break;
                    }
                }

                if (canceled)
                {
                    Console.WriteLine("\nProccess Canceled");
                    return false;
                }
                else
                {
                    File.Delete(name);
                    Console.WriteLine("\nOverriding file...");
                }
            }

            // Gets the current office version
            Excel.Application xlApp;
            if (version == null)
            {
                xlApp = new Excel.Application();
                version = xlApp.Version;
                xlApp.Quit();
                while (Marshal.ReleaseComObject(xlApp) != 0) { }
            }

            DisableTrustCenterSecurity();

            xlApp = new Excel.Application();
            Excel.Workbooks xlWbks = xlApp.Workbooks;
            Excel.Workbook xlWbk = xlWbks.Add();
            Excel.XlFileFormat xlFileFormat = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;

            //  Creating file
            AddModule(xlWbk, Executable.Files.VBE.Modules.Callbacks);

            name = Regex.Replace(name, "[/]", "\\");
            xlWbk.SaveAs(name, xlFileFormat);
            xlWbk.Close(true);
            while (Marshal.ReleaseComObject(xlWbk) != 0) { }

            xlWbks.Close();
            while (Marshal.ReleaseComObject(xlWbks) != 0) { }

            xlApp.Quit();
            while (Marshal.ReleaseComObject(xlApp) != 0) { }

            Console.WriteLine(@$"Excel file successfully created: '{name}'");

            bool added = CustomUI.AddCustomUI(name, customUI);

            EnableTrustCenterSecurity();

            return added;
        }
        private static void AddModule(Excel.Workbook xlWbk, string source)
        {

            xlWbk.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow;
            Excel.Application xlApp = xlWbk.Application;
            VBE.VBProject vbProject = xlWbk.VBProject;
            VBE.vbext_ComponentType vbComponentType = VBE.vbext_ComponentType.vbext_ct_StdModule;
            VBE.VBComponent vbModule = vbProject.VBComponents.Add(vbComponentType);
            VBE.CodeModule vbCode = vbModule.CodeModule;

            // Adds code to the module
            vbCode.Name = Path.GetFileName(source).Split(".")[0];
            vbCode.InsertLines(1, File.ReadAllText(source));
            xlApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityByUI;

            // Clean up
            while (Marshal.ReleaseComObject(vbProject) != 0) { }
            while (Marshal.ReleaseComObject(vbModule) != 0) { }
            while (Marshal.ReleaseComObject(vbCode) != 0) { }
        }
        private static void DisableTrustCenterSecurity()
        {
            // Disables the security for VBA Object Model
            RegistryKey VBOMKey = Registry.CurrentUser.OpenSubKey(subkey, true);
            if (VBOMKey == null) { VBOMKey = Registry.CurrentUser.CreateSubKey(subkey, true); }
            VBOMKey.SetValue(key, 0x01);
            VBOMKey.Close();
        }

        private static void EnableTrustCenterSecurity()
        {
            // Enables the security for VBA Object Model
            RegistryKey VBOMKey = Registry.CurrentUser.OpenSubKey(subkey, true);
            VBOMKey.DeleteValue(key);
            VBOMKey.Close();
        }

        private static string version = null;
        private static readonly string subkey = @$"Software\Microsoft\Office\{version}\Excel\Security";
        private const string key = "AccessVBOM";

    }
}
