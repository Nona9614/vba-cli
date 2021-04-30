using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;

namespace VBA.Project
{
    public static partial class Structure
    {
        public static void CreateFolders()
        {
            Console.WriteLine("Creating forms folder...");
            Directory.CreateDirectory(Paths.VBE.Forms);
            Console.WriteLine("Creating classes folder...");
            Directory.CreateDirectory(Paths.VBE.Classes);
            Console.WriteLine("Creating modules folder...");
            Directory.CreateDirectory(Paths.VBE.Modules);
            Console.WriteLine("Creating customUI folder...");
            Directory.CreateDirectory(Paths.VBE.CustomUI);
        }
        public static void CreateDefaultFiles()
        {
            Console.WriteLine("Creating default VBA class...");
            File.Copy(Executable.Files.VBE.Classes.Default, Files.VBE.Classes.Default, true);
            Console.WriteLine("Creating default VBA callbacks module...");
            File.Copy(Executable.Files.VBE.Modules.Callbacks, Files.VBE.Modules.Callbacks, true);
            Console.WriteLine("Creating default VBA customUI...");
            File.Copy(Executable.Files.VBE.CustomUI.Default, Files.VBE.CustomUI.Default, true);
        }
    }
    public static class Paths
    {
#if (DEBUG)
        private static readonly string _dfbase = @"D:\Documents\Personal\repos\apps\vba-cli\VBA CLI";
        private static string _base = _dfbase;
#else
        private static string _base = Directory.GetCurrentDirectory();
#endif
        public static string Base
        {
            get { return _base; }
            set { _base = value; }
        }
#if (DEBUG)
        public static string _vbe = "\\VBE";
        public static string Resources {
            get {
                if (string.Equals(_base, _dfbase, StringComparison.OrdinalIgnoreCase))
                {
                    _vbe = "\\VBE";
                    return $@"{_base}\Resources";
                }
                else return $@"{_base}\resources";
            } 
        }
#else
        public static string _vbe = "";
        public static string Resources { get { return $@"{Base}\resources"; } }
#endif
        public static class VBE
        {
            public static string Forms { get { return $@"{Resources}{_vbe}\forms"; } }
            public static string Modules { get { return $@"{Resources}{_vbe}\modules"; } }
            public static string Classes { get { return $@"{Resources}{_vbe}\classes"; } }
            public static string CustomUI { get { return $@"{Resources}{_vbe}\customUI"; } }
        }
    }
}
namespace VBA.Executable
{
    public static class Paths
    {
        public static string Base { get {
                string _name = Assembly.GetExecutingAssembly().GetName().Name;
                string _codebase = Regex.Replace(Assembly.GetExecutingAssembly().CodeBase, "[/]", "\\");
#if DEBUG
                string _bin = @"\bin\Debug\netcoreapp3.1\";
#else
                // If the program is installed uses the exe path as base
                string _bin = @"\bin\Release\netcoreapp3.1\";
#endif
                string _uri = @"file:\\\";
                return _codebase.Replace(_uri, "").Replace(@$"{_bin}{_name}.dll", ""); } 
        }
        public static string Resources { get { return $@"{Base}\Resources"; } }
        public static class VBE
        {
            public static string Forms { get { return $@"{Resources}\VBE\forms"; } }
            public static string Modules { get { return $@"{Resources}\VBE\modules"; } }
            public static string Classes { get { return $@"{Resources}\VBE\classes"; } }
            public static string CustomUI { get { return $@"{Resources}\VBE\customUI"; } }
        }

    }

}
