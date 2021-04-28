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
        public static bool CreateExcelFile(string name, string customUI = null)
        {
            // Adds macro enabled extension
            name = $"{name}.xlsm";
            string _route = Directory.GetParent(name).FullName;
            string _name = Directory.GetParent(name).Name;
            // Validate overriding
            if (File.Exists(name))
            {
                Console.Write("There is an excel file already created with this name and path. \nWould you like to override? (y/n) --> ");
                ConsoleKeyInfo key;
                int x = 0x1B;//0;
                int pos = Console.CursorLeft;
                bool canceled = false;//true;
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

            AddCustomUI(name, customUI);

            EnableTrustCenterSecurity();

            Console.WriteLine(@$"File successfully created: '{name}'");

            return true;
        }
        public static bool AddCustomUI(string excelFileName, string customUIName = null)
        {
            if (!IsXMLFile(customUIName) || !IsExcelFile(excelFileName))
            {
                return false;
            }

            FileStream _stream = File.Open(excelFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            ZipArchive archive = new ZipArchive(_stream, ZipArchiveMode.Update);
            ZipArchiveEntry rels = archive.GetEntry("_rels/.rels");
            ZipArchiveEntry customUI = archive.GetEntry("customUI/customUI.xml") ?? archive.CreateEntry("customUI/customUI.xml");

            byte[] relsBytes = File.ReadAllBytes(@$"{Executable.Files.VBE.CustomUI.Rels}");
            byte[] customUIBytes = File.ReadAllBytes(customUIName ?? @$"{Executable.Files.VBE.CustomUI.Default}");

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
