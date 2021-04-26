using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;

namespace VBA.Project
{
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
                return name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 ? $@"{Base}\{name}" : null;
            }
        }
#if (DEBUG)
        public static string Base { get { return @"D:/Documents/Personal/repos/apps/vba-cli/VBA CLI"; } }
#else
        public static string Base { get { return Directory.GetCurrentDirectory(); } }
#endif
        public static string Resources { get { return $@"{Base}/Resources"; } }
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
