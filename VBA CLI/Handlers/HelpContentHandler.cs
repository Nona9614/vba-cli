using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using VBA.Resources.Models;

namespace VBA.Handlers
{
    public static class HelpContentHandler
    {
        private static string FileName
        {
            get { return $@"{Project.Paths.Resources}/help.json"; }
        }

        private static HelpContentModel _model;
        private static HelpContentModel Model {
            get {
                if (_model == null)
                {
                    string fileContents = File.ReadAllText(FileName);
                    _model = JsonConvert.DeserializeObject<HelpContentModel>(fileContents);
                }
                return _model;
            }
        }
        public static string MainMessage()
        {
            string message = "These are the current commands, input 'help <command>' for more information of each:\n";

            foreach (CommandHelpContentModel command in Model.Commands)
            {
                message += $"- '{command.Name}':\n    {command.Description}\n\n";
            }

            return message;
        }
        public static string CommandContent(string name)
        {
            CommandHelpContentModel command = Model.Commands.Find(x => x.Name.Contains(name));

            if (command == null)
            {
                Console.WriteLine($"The command '{name}' doesn't exist");
                return null;
            }

            string message = $"'{name}' subcommands:\n\n";

            foreach (SubcommandHelpContentModel subcommand in command.Subcommands)
            {
                message += $"Use: {subcommand.Use}\n- '{subcommand.Name}': {subcommand.Description}\n\n";
            }

            return message;

        }

        public static string SubcommandContent(string command, string subcommand)
        {
            CommandHelpContentModel _command = Model.Commands.Find(x => x.Name.Contains(command));
            SubcommandHelpContentModel _subcommand = _command.Subcommands.Find(x => x.Name.Contains(subcommand));

            if (_command == null)
            {
                Console.WriteLine($"The command '{command}' doesn't exist");
                return null;
            }
            if (_subcommand == null)
            {
                Console.WriteLine($"The command '{command}' doesn't have a subcommand {subcommand}");
                return null;
            }

            string message = $"Use: {_subcommand.Use}\n'{subcommand}' arguments:\n\n";
            string optional = "";
            foreach (ArgumentHelpContentModel argument in _subcommand.Arguments)
            {
                optional = argument.Required ? optional : "[Optional] ";
                message += $"- { optional }'{argument.Name}': {argument.Description}\n";
                if (argument.Default != null)
                {
                    message += $"   Default value: {argument.Default}\n";
                }
                if (argument.Options != null)
                {
                    if (argument.Default != null) message += "\n";
                    message += $"   Options:\n";
                    foreach (string option in argument.Options)
                    {
                        message += $"    - {option}\n";
                    }
                }
                message += "\n";
            }

            return message;

        }

    }
}
