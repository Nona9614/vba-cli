using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Project
{
    public static class Files
    {
        public static string CustomUI{ get { return $@"{Paths.Resources}/customUI.xml"; } }
    }
}
namespace VBA.Executable
{
    public static class Files
    {
        public static string CustomUI { get { return $@"{Paths.Resources}/customUI.xml"; } }
        public static string CallbacksModule { get { return $@"{Paths.Resources}/Callbacks.bas"; } }
        public static string Rels { get { return $@"{Paths.Resources}/.rels"; } }
        public static string HelpContent { get { return $@"{Paths.Resources}/help.json"; } }
    }
}
