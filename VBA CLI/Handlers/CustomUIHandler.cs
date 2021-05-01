using System;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Excel = VBA.Handlers.ExcelHandler;

namespace VBA.Handlers
{
    public static class CustomUIHandler
    {
        public const string FileDefaultName = "customUI.xml";
        // If has not xml extension then adds to it
        public static void AddXmlExtension(ref string name)
        {
            if (!IsXMLFile(name)) name += ".xml";
        }
        public static bool FileExists(string customUI)
        {
            if (!File.Exists(customUI))
            {
                Console.WriteLine($"The customUI file '{customUI}' doesn't exist");
                return false;
            }
            else return true;
        }
        //Checks for valid if xml file is valid
        public static bool IsXMLFile(string name) => Regex.Match(Path.GetExtension(name), @".*\.[xX][mM][lL]").Length > 0;
        public static bool AddCustomUI(string excel, string customUI)
        {
            if (!IsXMLFile(customUI)) return false;

            FileStream _stream = File.Open(excel, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            ZipArchive archive = new ZipArchive(_stream, ZipArchiveMode.Update);
            ZipArchiveEntry entryRels = archive.GetEntry("_rels/.rels");
            ZipArchiveEntry entryCustomUI = archive.GetEntry("customUI/customUI.xml") ?? archive.CreateEntry("customUI/customUI.xml");

            if (entryRels == null)
            {
                Console.WriteLine($"The archive '{excel}' has no internal '_rels/.rels' compressed entry");
                return false;
            }

            byte[] relsBytes = File.ReadAllBytes(@$"{Executable.Files.VBE.CustomUI.Rels}");
            byte[] customUIBytes = File.ReadAllBytes(customUI ?? @$"{Executable.Files.VBE.CustomUI.Default}");

            Stream _rels = entryRels.Open();
            Stream _customUI = entryCustomUI.Open();

            _rels.Write(relsBytes, 0, relsBytes.Length);
            _rels.Dispose();
            _customUI.Write(customUIBytes, 0, customUIBytes.Length);
            _customUI.Dispose();

            archive.Dispose();
            Console.WriteLine(@$"Custom UI added successfully from: '{customUI}'");

            return true;
        }
        public static bool GenerateDefaultCustomUI()
        {
            // Validate overriding
            if (File.Exists(Project.Files.VBE.CustomUI.Default))
            {
                Console.Write("There is a customUI file already created. \nWould you like to override it? (y/n) --> ");
                if (!(Regex.Match(Console.ReadLine().Trim(), "^y*").Length > 0))
                {
                    Console.WriteLine("Proccess Canceled");
                    return false;
                }
                else
                {
                    File.Delete(Project.Files.VBE.CustomUI.Default);
                    File.Copy(Executable.Files.VBE.CustomUI.Default, Project.Files.VBE.CustomUI.Default);
                    Console.WriteLine("CustomUI file created properly");
                }
            }
            else
            {
                Directory.CreateDirectory(Project.Paths.Resources);
                File.Copy(Executable.Files.VBE.CustomUI.Default, Project.Files.VBE.CustomUI.Default);
                Console.WriteLine("CustomUI file created successfully");
            }
            return true;
        }

    }
}
