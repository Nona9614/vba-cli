﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;

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
        // If it is a valid file name with no path, a default will be used
        public static string CheckForDefaultPath(string name, string path)
        {
            if (File.Exists(name))
            {
                return name;
            }
            else
            {
                return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 ? $@"{path}\{name}" : null;
            }
        }
        // This overload uses the project path as the preset
        public static string CheckForDefaultPath(string name)
        {
            if (File.Exists(name))
            {
                return name;
            }
            else
            {
                return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 ? $@"{Base}/{name}" : null;
            }
        }

#if (DEBUG)
        private static string _base = @"D:/Documents/Personal/repos/apps/vba-cli/VBA CLI";
        public static string Base
        { 
            get { return _base; }
            set { _base = value; }
        }
#else
        private static string _base = Directory.GetCurrentDirectory();
        public static string Base 
        { 
            get { return Directory.GetCurrentDirectory(); } 
            set { _base = value; }
        }
#endif
        public static string Resources { get { return $@"{Base}/resources"; } }

        public static class VBE
        {
            public static string Forms { get { return $@"{Resources}/VBE/forms"; } }
            public static string Modules { get { return $@"{Resources}/VBE/modules"; } }
            public static string Classes { get { return $@"{Resources}/VBE/classes"; } }
            public static string CustomUI { get { return $@"{Resources}/VBE/customUI"; } }
        }
    }
}
namespace VBA.Executable
{
    public static class Paths
    {
#if DEBUG
        public static string Base { get { return Assembly.GetExecutingAssembly().CodeBase.Replace($"/{Assembly.GetExecutingAssembly().GetName().Name}.dll", "").Replace("file:///", "").Replace(@"/bin/Debug/netcoreapp3.1", ""); } }
#else
        public static string Base { get { return Assembly.GetExecutingAssembly().CodeBase.Replace($"/{Assembly.GetExecutingAssembly().GetName().Name}.dll", "").Replace("file:///", "").Replace(@"/bin/Release/netcoreapp3.1", ""); } }
#endif
        public static string Resources { get { return $@"{Base}/Resources"; } }
    }

}
