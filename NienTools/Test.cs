using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace NienTools
{
    internal class Test
    {
        [ExcelFunction(Description = "Tra ve loi chao")]
        public static string HELLO(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Hello!";
            return "Hello, " + name.Trim() + "!";
        }
    }
}
