using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace VBA.Handlers
{
    public static class XMLHandler
    {
        public static bool GenerateDefaultCustomUI()
        {
            // Validate overriding
            if (File.Exists(Project.Files.CustomUI))
            {
                Console.Write("There is a customUI file already created. \nWould you like to override it? (y/n) --> ");
                if (!(Regex.Match(Console.ReadLine().Trim(), "^y*").Length > 0))
                {
                    Console.WriteLine("Proccess Canceled");
                    return false;
                }
                else
                {
                    File.Delete(Project.Files.CustomUI);
                    File.Copy(Executable.Files.CustomUI, Project.Files.CustomUI);
                    Console.WriteLine("CustomUI file created properly");
                }
            }
            else
            {
                Directory.CreateDirectory(Project.Paths.Resources);
                File.Copy(Executable.Files.CustomUI, Project.Files.CustomUI);
                Console.WriteLine("CustomUI file created successfully");
            }
            return true;
        }

    }
}
