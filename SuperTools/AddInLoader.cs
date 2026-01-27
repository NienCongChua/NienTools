using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperTools
{
    public class AddInLoader : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
