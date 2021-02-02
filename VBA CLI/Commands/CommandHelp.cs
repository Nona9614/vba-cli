using System;
using System.Collections.Generic;
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
        return true;
    }

    public void Dispose()
    {
        throw new NotImplementedException();
    }
}
