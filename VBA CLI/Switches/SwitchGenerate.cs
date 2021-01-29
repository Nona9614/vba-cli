using System.Collections.Generic;
using Excel = VBA.ExcelHandler;
using static VBA.Program;
using System;

namespace VBA.Switches
{
    public class SwitchGenerate : ISwitch, IDisposable
    {
        private static SwitchGenerate instance;
        public static SwitchGenerate Instance
        {
            get
            {
                instance ??= new SwitchGenerate();
                return instance;
            }
        }

        public bool Call(List<string> parameters)
        {
            bool result;
            switch (parameters[0])
            {
                case "file":
                    // Checks if a directory value was set
                    result = Excel.CreateExcelFile(parameters[1], parameters.Count > 2 ? parameters[1] : null);
                    break;
                default:
                    result = false;
                    break;
            }
            return result;
        }

        public void Dispose()
        {
            instance.Dispose();
        }

    }

}
