using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using VBE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace VBA
{
    public static class ExcelHandler
    {
        //  Returns true if the creation succeed, if not returns false
        public static bool CreateExcelFile(string name, string path)
        {
            // If was not set any directory then use current
            path ??= Directory.GetCurrentDirectory();

            // Check for valid file name or valid route
            if (!(name.IndexOfAny(Path.GetInvalidFileNameChars()) == -1)) {
                Console.WriteLine($"The name '{name}' is not a valid value");
                return false;
            } 
            else if(Directory.Exists(path))
            {
                Console.WriteLine($"The path '{path}' is not a valid value");
                return false;
            }

            // Gets the current office version
            Excel.Application xlApp;
            if (version == null)
            {
                xlApp = new Excel.Application();
                version = xlApp.Version;
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            DisableTrustCenterSecurity();

            xlApp = new Excel.Application();
            Excel.Workbook xlWbk = xlApp.Workbooks.Add();
            Excel.XlFileFormat xlFileFormat = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;

            //  Creating file
            string fullName = $@"{path}/{name}.xlsm";
            AddCallbacksModule(xlWbk);

            xlApp.Visible = true;
            xlWbk.SaveAs(fullName, xlFileFormat);
            xlWbk.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWbk);

            AddCustomUI(fullName, null);

            EnableTrustCenterSecurity();

            return true;
        }

        private static void AddCustomUI(string fileRoute, string uiContent)
        {

            FileStream _stream = File.Open(fileRoute, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            ZipArchive archive = new ZipArchive(_stream, ZipArchiveMode.Update);
            ZipArchiveEntry rels = archive.GetEntry("_rels/.rels");
            ZipArchiveEntry customUI = archive.GetEntry("customUI/customUI.xml") ?? archive.CreateEntry("customUI/customUI.xml");

            byte[] relsBytes = File.ReadAllBytes(@$"{projectPath}/resources/.rels");
            byte[] customUIBytes = File.ReadAllBytes(@$"{projectPath}/resources/customUI.xml");

            rels.Open().Write(relsBytes, 0, relsBytes.Length);
            customUI.Open().Write(customUIBytes, 0, customUIBytes.Length);

            archive.Dispose();
        }

        private static void AddCallbacksModule(Excel.Workbook xlWbk)
        {

            xlWbk.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow;

            VBE.vbext_ComponentType type = VBE.vbext_ComponentType.vbext_ct_StdModule;
            VBE.VBComponent defaultModule = xlWbk.VBProject.VBComponents.Add(type);
            VBE.CodeModule code = defaultModule.CodeModule;

            // Adds code to the module
            code.Name = "Callbacks";
            code.InsertLines(1, File.ReadAllText($"{projectPath}/resources/Callbacks.bas"));

            xlWbk.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityByUI;
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

#if (DEBUG)
        private static string projectPath = @"D:\Documents\Personal\repos\apps\vba-cli\VBA CLI";
#else
        private static string projectPath = Assembly.GetExecutingAssembly().CodeBase.Replace($"/{Assembly.GetExecutingAssembly().GetName().Name}.dll", "").Replace("file:///", "");
#endif
    }
}
