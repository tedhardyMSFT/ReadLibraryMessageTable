using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ReadLibraryMessageTable
{
    class NativeMethods
    {
        /// <summary>
        /// For FormatMessage, the flag for retreiving a message by lookup.
        /// see: https://msdn.microsoft.com/en-us/library/windows/desktop/ms679351(v=vs.85).aspx
        /// </summary>
        [Flags]
        internal enum FormatMessageFlags : uint
        {
            FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100,
            FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200,
            FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000,
            FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x00002000,
            FORMAT_MESSAGE_FROM_HMODULE = 0x00000800,
            FORMAT_MESSAGE_FROM_STRING = 0x00000400,
        }

        /// <summary>
        /// Loads the specified module into the address space of the calling process. The specified module may cause other modules to be loaded.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern IntPtr LoadLibrary(string fileName);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern IntPtr LoadLibraryEx(string fileName, IntPtr hFile, int dwFlags);



        /// <summary>
        /// Determines the location of a resource with the specified type and name in the specified module.
        /// Reference: https://msdn.microsoft.com/en-us/library/windows/desktop/ms648042(v=vs.85).aspx
        /// </summary>
        /// <param name="hModule">A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process.</param>
        /// <param name="lpID">The name of the resource. Alternately, rather than a pointer, this parameter can be MAKEINTRESOURCE(ID), where ID is the integer identifier of the resource. For more information, see the Remarks section below.</param>
        /// <param name="lpType">The resource type. Alternately, rather than a pointer, this parameter can be MAKEINTRESOURCE(ID), where ID is the integer identifier of the given resource type.</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr FindResource(IntPtr hModule, int lpID, int lpType);
        //TODO:DOC
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);
        //TODO:DOC
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LockResource(IntPtr hResData);
        //TODO:DOC
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GetProcessHeap();
        //TODO:DOC
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int HeapFree(IntPtr hHeap, int dwFlags, IntPtr lpMem);

        /// <summary>
        /// Frees the loaded dynamic-link library (DLL) module and, if necessary, decrements its reference count. 
        /// When the reference count reaches zero, the module is unloaded from the address space of the calling process and the handle is no longer valid.
        /// </summary>
        /// <param name="hModule">A handle to the loaded library module</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern uint FreeLibrary(IntPtr hModule);
        //TODO:DOC
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int FormatMessage(
            FormatMessageFlags dwFlags,
            IntPtr lpSource,
            uint dwMessageId,
            uint dwLanguageId,
            ref IntPtr lpBuffer,
            uint nSize,
            IntPtr Arguments);

    } // class NativeMethods
} // namespace ReadLibraryMessageTable
