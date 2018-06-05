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

            //System.Threading.Thread.Sleep(1000);

            /// zune has multiple languages
            Messages = ReadMessageTable.EnumerateMessageTable(@"C:\Program Files\Zune\ZuneResources.dll");
            //Messages = ReadMessageTable.EnumerateMessageTable(@"C:\Windows\System32\notepad.exe");
            string Message = ReadMessageTable.ReadModuleSingleMessage(@"C:\Program Files\Zune\ZuneResources.dll", 3405643777);
            // decimal representation of hex country code (https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-LCID/[MS-LCID].pdf) 
            Message = ReadMessageTable.ReadModuleSingleMessage(@"C:\Program Files\Zune\ZuneResources.dll", 3222087522, 1028);

            Console.WriteLine(Messages.Keys.Count.ToString());

        }
    }
}
