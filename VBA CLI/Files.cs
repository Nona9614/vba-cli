using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Project
{
    namespace Files
    {
        namespace VBE
        {
            public static class Classes
            {
                public static string Default { get { return $@"{Paths.VBE.Classes}/Class.cls"; } }
                public static string StaticClass { get { return $@"{Paths.VBE.Classes}/StaticClass.cls"; } }
                public static string Interface { get { return $@"{Paths.VBE.Classes}/Interface.cls"; } }
            }
            public static class CustomUI
            {
                public static string Default { get { return $@"{Paths.VBE.CustomUI}/customUI.xml"; } }
            }
            public static class Modules
            {
                public static string Default { get { return $@"{Paths.VBE.Modules}/Module.bas"; } }
                public static string Callbacks { get { return $@"{Paths.VBE.Modules}/Callbacks.bas"; } }
            }
        }
    }
}
namespace VBA.Executable
{
    namespace Files
    {
        namespace VBE
        {
            public static class Classes
            {
                public static string Default { get { return $@"{Paths.Resources}/VBE/Classes/Class.cls"; } }
                public static string StaticClass { get { return $@"{Paths.Resources}/VBE/Classes/StaticClass.cls"; } }
                public static string Interface { get { return $@"{Paths.Resources}/VBE/Classes/Interface.cls"; } }
            }
            public static class CustomUI
            {
                public static string Rels { get { return $@"{Paths.Resources}/VBE/CustomUI/customUI.xml"; } }
                public static string Default { get { return $@"{Paths.Resources}/VBE/CustomUI/customUI.xml"; } }
            }
            public static class Modules
            {
                public static string Default { get { return $@"{Paths.Resources}/VBE/Modules/Module.bas"; } }
                public static string Callbacks { get { return $@"{Paths.Resources}/VBE/Modules/Callbacks.bas"; } }
            }
        }
        public static class CLI
        {
            public static string HelpContent { get { return $@"{Paths.Resources}/help.json"; } }
        }
    }
}
