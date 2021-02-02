using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Switches
{
    interface ICommand
    {
        bool Call(List<string> parameters);
    }
}
