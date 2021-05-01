using System;
using System.IO;
using System.Text.RegularExpressions;

namespace VBA.Project
{
    namespace Files
    {
        public static class Verify
        {
            // Checks for invalid chars
            public static bool IsValidFileName(string name) => Regex.Match(name, @"^[^\\\/:*<>|?]+$").Success;
            public static bool IsValidFileNameWithMessage(string name)
            {
                if (!IsValidFileName(name))
                {
                    Console.WriteLine($"The file '{name}' has not valid format");
                    return false;
                }
                return true;
            }

            public static bool IsValidFolderName(string name) => Regex.Match(name, @"^[a-zA-Z]:(\\|\/)[^:*<>|?.]+$").Success;
            public static bool IsValidFolderNameWithMessage(string name)
            {
                if (!IsValidFolderName(name))
                {
                    Console.WriteLine($"The folder '{name}' has not valid format");
                    return false;
                }
                return true;
            }
            public static bool Name(ref string name, ref string _base)
            {
                bool isNameInvalid = name == null;
                bool isBaseInvalid = _base == null;
                if ((isNameInvalid || isBaseInvalid) || !Directory.Exists(_base))
                {
                    if (isNameInvalid) Console.WriteLine($"Name has null reference");
                    if (isBaseInvalid) Console.WriteLine($"Base has null reference");
                    else Console.WriteLine($"Directory '{_base}' doesn't exist");
                    name = null;
                    _base = null;
                    return false;
                }
                _base = Directory.GetParent(_base + "\\remove").FullName;
                MatchCollection matches = Regex.Matches(name, @"(\\|\/)");
                if (matches.Count > 0)
                {
                    int _index = matches[^1].Index;
                    string _subbase = name.Remove(_index);
                    name = name.Remove(0, _subbase.Length + 1);
                    _base += "\\" + _subbase;
                }
                isNameInvalid = !IsValidFileName(name);
                isBaseInvalid = !IsValidFolderName(_base);
                if (isNameInvalid || isBaseInvalid)
                {
                    if (isNameInvalid) Console.WriteLine($"Name '{name}' has invalid format");
                    if (isBaseInvalid) Console.WriteLine($"Directory '{_base}' has invalid format");
                    name = null;
                    _base = null;
                    return false;
                }
                return true;
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
