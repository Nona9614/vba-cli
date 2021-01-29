using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Switches
{
    interface ISwitch
    {
        bool Call(List<string> parameters);
    }
}
