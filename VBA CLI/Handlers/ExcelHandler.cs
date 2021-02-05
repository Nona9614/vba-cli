using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using VBE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace VBA
{
    public static class ExcelHandler
    {
        //  Returns true if the creation succeed, if not returns false
        public static bool CreateExcelFile(string name)
        {
            // Adds macro enabled extension
            name = $"{name}.xlsm";

            // Validate overriding
            if (File.Exists(name))
            {
                Console.Write("There is a file already created with this name and path. \nWould you like to override? (y/n) --> ");
                if (!(Regex.Match(Console.ReadLine().Trim(), "^y*").Length > 0))
                {
                    Console.WriteLine("Proccess Canceled");
                    return false;
                }
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
            xlApp.DisplayAlerts = false;

            //  Creating file
            AddCallbacksModule(xlWbk);

            xlWbk.SaveAs(name, xlFileFormat);
            xlApp.DisplayAlerts = true;
            xlWbk.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWbk);

            AddCustomUI(name, Executable.Files.CustomUI);

            EnableTrustCenterSecurity();

            Console.WriteLine(@$"Created successfully file: '{name}'");

            return true;
        }
        public static bool AddCustomUI(string excelFileName, string customUIName)
        {
            if (!IsXMLFile(customUIName) || !IsExcelFile(excelFileName))
            {
                return false;
            }

            FileStream _stream = File.Open(excelFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            ZipArchive archive = new ZipArchive(_stream, ZipArchiveMode.Update);
            ZipArchiveEntry rels = archive.GetEntry("_rels/.rels");
            ZipArchiveEntry customUI = archive.GetEntry("customUI/customUI.xml") ?? archive.CreateEntry("customUI/customUI.xml");

            byte[] relsBytes = File.ReadAllBytes(@$"{Executable.Files.Rels}");
            byte[] customUIBytes = File.ReadAllBytes(customUIName ?? @$"{Executable.Files.CustomUI}");

            Stream _rels = rels.Open();
            Stream _customUI = customUI.Open();

            _rels.Write(relsBytes, 0, relsBytes.Length);
            _rels.Dispose();
            _customUI.Write(customUIBytes, 0, customUIBytes.Length);
            _customUI.Dispose();

            archive.Dispose();
            Console.WriteLine("Custom UI added successfully");

            return true;
        }
        private static bool IsExcelFile(string name)
        {
            if (!File.Exists(name))
            {
                Console.WriteLine($"The excel file '{name}' was not found");
                return false;
            }
            if (Regex.Match(Path.GetExtension(name), @".*\.[xX][lL]*").Length > 0)
            {
                return true;
            }
            else 
            {
                Console.WriteLine($"This is not an excel file '{name}'");
                return false; ;
            }
        }
        private static bool IsXMLFile(string name)
        {
            if (!File.Exists(name))
            {
                Console.WriteLine($"The xml file '{name}' was not found");
                return false;
            }
            if (Regex.Match(Path.GetExtension(name), @".*\.[xX][mM][lL]").Length > 0)
            {
                return true;
            }
            else
            {
                Console.WriteLine($"This is not an xml file '{name}'");
                return false; ;
            }
        }
        private static void AddCallbacksModule(Excel.Workbook xlWbk)
        {

            xlWbk.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow;

            VBE.vbext_ComponentType type = VBE.vbext_ComponentType.vbext_ct_StdModule;
            VBE.VBComponent defaultModule = xlWbk.VBProject.VBComponents.Add(type);
            VBE.CodeModule code = defaultModule.CodeModule;

            // Adds code to the module
            code.Name = "Callbacks";
            code.InsertLines(1, File.ReadAllText($"{Executable.Files.CallbacksModule}"));

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

        private static bool IsValidFileName(string name)
        {
            return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;
        }

        private static string version = null;
        private static readonly string subkey = @$"Software\Microsoft\Office\{version}\Excel\Security";
        private const string key = "AccessVBOM";

    }
}
