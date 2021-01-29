using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using static VBA_CLI.Program;
using Excel = Microsoft.Office.Interop.Excel;

namespace VBA_CLI.Switches
{
    public static class SwitchGenerate
    {
        public static int Call(string parameter)
        {

            //  Checks if the passed parameter contains any character that may conflict with the file creation
            if (parameter.IndexOfAny(Path.GetInvalidFileNameChars()) == -1)
            {
                CreateExcelFile(parameter);
                return (int)ReturnCodes.SwitchSucceed;
            }
            else
            {
                return (int)ReturnCodes.SwitchFailed;
            }

        }

        private static void CreateExcelFile(string fileFullName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWbk = xlApp.Workbooks.Add();

            //  Add code here
            fileFullName = @"D:\Downloads\TestExcel.xlsm";
            if(!File.Exists(fileFullName))
            {
                xlWbk.SaveAs2(fileFullName, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
            }

            xlWbk.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWbk);

            AddCustomUI(fileFullName, null);
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

#if (DEBUG)
        private static string projectPath = @"D:\Documents\Personal\repos\apps\vba-cli\VBA CLI";
#else
        private static string projectPath = Assembly.GetExecutingAssembly().CodeBase.Replace($"/{Assembly.GetExecutingAssembly().GetName().Name}.dll", "").Replace("file:///", "");
#endif
    }

}
