using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace VBA.Project
{
    namespace Files
    {
        public static class Verify
        {
            // Important to notice this modifies the _base route during all the code
            public static void Name(ref string name, ref string _base)
            {
                bool isValidName = Regex.Match(name, @"^[\w\-\(\)\[\]\/\\]+$").Length > 0;
                bool isValidRoute = Directory.Exists(_base);
                if (isValidName && isValidRoute)
                {
                    _base = Regex.Replace($"{_base}\\{name}", "[/]", "\\");
                    name = $"{_base}".Split("\\")[^1];
                }
                else
                {
                    if (!isValidName) Console.WriteLine($"Name '{name}' has invalid format");
                    if (!isValidRoute) Console.WriteLine($"Route'{_base}' doesn't exists");
                    name = null;
                    _base = null;
                }
            }
        }
        namespace VBE
        {
            public static class Classes
            {
                public static string Default { get { return $@"{Paths.VBE.Classes}\Class.cls"; } }
                public static string StaticClass { get { return $@"{Paths.VBE.Classes}\StaticClass.cls"; } }
                public static string Interface { get { return $@"{Paths.VBE.Classes}\Interface.cls"; } }
            }
            public static class CustomUI
            {
                public static string Default { get { return $@"{Paths.VBE.CustomUI}\customUI.xml"; } }
            }
            public static class Modules
            {
                public static string Default { get { return $@"{Paths.VBE.Modules}\Module.bas"; } }
                public static string Callbacks { get { return $@"{Paths.VBE.Modules}\Callbacks.bas"; } }
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
                public static string Default { get { return $@"{Paths.Resources}\VBE\Classes\Class.cls"; } }
                public static string StaticClass { get { return $@"{Paths.Resources}\VBE\Classes\StaticClass.cls"; } }
                public static string Interface { get { return $@"{Paths.Resources}\VBE\Classes\Interface.cls"; } }
            }
            public static class CustomUI
            {
                public static string Rels { get { return $@"{Paths.Resources}\VBE\CustomUI\.rels"; } }
                public static string Default { get { return $@"{Paths.Resources}\VBE\CustomUI\customUI.xml"; } }
            }
            public static class Modules
            {
                public static string Default { get { return $@"{Paths.Resources}\VBE\Modules\Module.bas"; } }
                public static string Callbacks { get { return $@"{Paths.Resources}\VBE\Modules\Callbacks.bas"; } }
            }
        }
        public static class CLI
        {
            public static string HelpContent { get { return $@"{Paths.Resources}\help.json"; } }
        }
    }
}
