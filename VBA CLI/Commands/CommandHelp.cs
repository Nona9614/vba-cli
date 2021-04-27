using System;
using System.Collections.Generic;
using VBA.Handlers;
using VBA.Switches;

public class CommandHelp: ICommand, IDisposable
{
	public CommandHelp()
	{

	}

    private static CommandHelp instance;
    public static CommandHelp Instance
    {
        get
        {
            instance ??= new CommandHelp();
            return instance;
        }
    }

    public bool Call(List<string> parameters)
    {
        string message = null;
        bool result = false;
        if (parameters == null)
        {
            message = HelpContentHandler.MainMessage();
        }
        else
        {
            if ( parameters.Count == 1)
            {
                message = HelpContentHandler.CommandContent(parameters[0]);
            }
            else if (parameters.Count == 2) {
                message = HelpContentHandler.SubcommandContent(parameters[0], parameters[1]);
            }            
        }
        if (message != null)
        {
            Console.WriteLine(message);
            result = true;
        }
        return result;
    }

    public void Dispose()
    {
        instance.Dispose();
    }
}
