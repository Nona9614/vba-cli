using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Commands
{
    interface ICommand
    {
        bool Call(List<string> parameters);
    }
}
