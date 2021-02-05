using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;

namespace VBA.Project
{
    public static class Paths
    {
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
        public static string Base { get { return Assembly.GetExecutingAssembly().CodeBase.Replace($"/{Assembly.GetExecutingAssembly().GetName().Name}.dll", "").Replace("file:///", "").Replace(@"/bin/Release/netcoreapp3.1", ""); } }
        public static string Resources { get { return $@"{Base}/Resources"; } }
    }
}
