using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadLibraryMessageTable;

namespace RlmtTestShell
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, string> Messages = null;

            Messages = ReadMessageTable.EnumerateMessageTable(@"C:\Windows\System32\msobjs.dll");

            string Message = ReadMessageTable.ReadModuleSingleMessage(@"C:\Windows\system32\msobjs.dll", 1612);

            Console.WriteLine(Message);

            Console.ReadLine();
        }
    }
}
