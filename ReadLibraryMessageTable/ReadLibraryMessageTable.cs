﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ReadLibraryMessageTable
{
    public class ReadMessageTable
    {
        public static Dictionary<string, string> EnumerateMessageTable(string LibraryPath)
        {
            Dictionary<string, string> Messages = new Dictionary<string, string>();

            // TODO:MakeUseful - add code to load library, find resource, lock resources, load mesage table, enumerate and add to dictionary. Simple, eh?

            return Messages;
        } // public static Dictionary<string,string> EnumerateMessageTable(
    }
}
