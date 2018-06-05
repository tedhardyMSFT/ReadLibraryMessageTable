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
        /// For FormatMessage, the flags for retreiving a message by lookup.
        /// see: https://msdn.microsoft.com/en-us/library/windows/desktop/ms679351(v=vs.85).aspx
        /// </summary>
        [Flags]
        internal enum FormatMessageFlags : uint
        {
            /// <summary>
            /// The function allocates a buffer large enough to hold the formatted message, and places a pointer to the allocated buffer at the address specified by lpBuffer. 
            /// </summary>
            FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100,
            /// <summary>
            /// Insert sequences in the message definition are to be ignored and passed through to the output buffer unchanged. 
            /// </summary>
            FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200,
            /// <summary>
            /// The function should search the system message-table resource(s) for the requested message. If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system message table if the message is not found in the module specified by lpSource. This flag cannot be used with FORMAT_MESSAGE_FROM_STRING.
            /// </summary>
            FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000,
            /// <summary>
            /// The Arguments parameter is not a va_list structure, but is a pointer to an array of values that represent the arguments.
            /// </summary>
            FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x00002000,
            /// <summary>
            /// The lpSource parameter is a module handle containing the message-table resource(s) to search. If this lpSource handle is NULL, the current process's application image file will be searched. This flag cannot be used with FORMAT_MESSAGE_FROM_STRING.
            /// </summary>
            FORMAT_MESSAGE_FROM_HMODULE = 0x00000800,
            /// <summary>
            /// The lpSource parameter is a pointer to a null-terminated string that contains a message definition. The message definition may contain insert sequences, just as the message text in a message table resource may. This flag cannot be used with FORMAT_MESSAGE_FROM_HMODULE or FORMAT_MESSAGE_FROM_SYSTEM.
            /// </summary>
            FORMAT_MESSAGE_FROM_STRING = 0x00000400,
        }

        /// <summary>
        /// Loads the specified module into the address space of the calling process. The specified module may cause other modules to be loaded.
        /// </summary>
        /// <param name="fileName">A string that specifies the file name of the module to load.</param>
        /// <param name="hFile">This parameter is reserved for future use. It must be NULL.</param>
        /// <param name="dwFlags">FormatMessageFlags specifying the action to be taken when loading the module. If no flags are specified, the behavior of this function is identical to that of the LoadLibrary function.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Retrieves a handle that can be used to obtain a pointer to the first byte of the specified resource in memory.
        /// </summary>
        /// <param name="hModule">A handle to the module whose executable file contains the resource. If hModule is NULL, the system loads the resource from the module that was used to create the current process.</param>
        /// <param name="hResInfo">A handle to the resource to be loaded. This handle is returned by the FindResource or FindResourceEx function.</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

        /// <summary>
        /// Retrieves a pointer to the specified resource in memory.
        /// </summary>
        /// <param name="hResData">A handle to the resource to be accessed. The LoadResource function returns this handle. Note that this parameter is listed as an HGLOBAL variable only for backward compatibility. Do not pass any value as a parameter other than a successful return value from the LoadResource function.</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LockResource(IntPtr hResData);

        /// <summary>
        /// Retrieves a handle to the default heap of the calling process. This handle can then be used in subsequent calls to the heap functions.
        /// </summary>
        /// <returns> the function succeeds, the return value is a handle to the calling process's heap.
        /// If the function fails, the return value is NULL.To get extended error information, call GetLastError.</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GetProcessHeap();

        /// <summary>
        /// Frees a memory block allocated from a heap by the HeapAlloc or HeapReAlloc function.
        /// </summary>
        /// <param name="hHeap">handle to the heap whose memory block is to be freed. This handle is returned by either the HeapCreate or GetProcessHeap function.</param>
        /// <param name="dwFlags">The heap free options. Specifying the following value overrides the corresponding value specified in the flOptions parameter when the heap was created by using the HeapCreate function. </param>
        /// <param name="lpMem">A pointer to the memory block to be freed. This pointer is returned by the HeapAlloc or HeapReAlloc function. If this pointer is NULL, the behavior is undefined.</param>
        /// <returns>If the function succeeds, the return value is nonzero.
        ///  If the function fails, the return value is zero.An application can call GetLastError for extended error information</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int HeapFree(IntPtr hHeap, int dwFlags, IntPtr lpMem);

        /// <summary>
        /// If this value is used with LoadLibraryEx, the system maps the file into the calling process's virtual address space as if it were a data file. 
        /// Nothing is done to execute or prepare to execute the mapped file. Therefore, you cannot call functions like GetModuleFileName, GetModuleHandle
        /// or GetProcAddress with this DLL. Using this value causes writes to read-only memory to raise an access violation. Use this flag when you want 
        /// to load a DLL only to extract messages or resources from it.
        /// </summary>
        public const int LOAD_LIBRARY_AS_DATAFILE = 2;

        /// <summary>
        /// Frees the loaded dynamic-link library (DLL) module and, if necessary, decrements its reference count. 
        /// When the reference count reaches zero, the module is unloaded from the address space of the calling process and the handle is no longer valid.
        /// </summary>
        /// <param name="hModule">A handle to the loaded library module</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern uint FreeLibrary(IntPtr hModule);

        /// <summary>
        /// Formats a message string. The function requires a message definition as input. The message definition can come from a buffer passed into the function. It can come from a message table resource in an already-loaded module. Or the caller can ask the function to search the system's message table resource(s) for the message definition. The function finds the message definition in a message table resource based on a message identifier and a language identifier. The function copies the formatted message text to an output buffer, processing any embedded insert sequences if requested.
        /// </summary>
        /// <param name="dwFlags">The formatting options, and how to interpret the lpSource parameter. The low-order byte of dwFlags specifies how the function handles line breaks in the output buffer. The low-order byte can also specify the maximum width of a formatted output line.</param>
        /// <param name="lpSource">The location of the message definition. The type of this parameter depends upon the settings in the dwFlags parameter.</param>
        /// <param name="dwMessageId">The message identifier for the requested message. This parameter is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
        /// <param name="dwLanguageId">The language identifier for the requested message. This parameter is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
        /// <param name="lpBuffer">A pointer to a buffer that receives the null-terminated string that specifies the formatted message. If dwFlags includes FORMAT_MESSAGE_ALLOCATE_BUFFER, the function allocates a buffer using the LocalAlloc function, and places the pointer to the buffer at the address specified in lpBuffer.
        /// This buffer cannot be larger than 64K bytes.</param>
        /// <param name="nSize">f the FORMAT_MESSAGE_ALLOCATE_BUFFER flag is not set, this parameter specifies the size of the output buffer, in TCHARs. If FORMAT_MESSAGE_ALLOCATE_BUFFER is set, this parameter specifies the minimum number of TCHARs to allocate for an output buffer.
        /// The output buffer cannot be larger than 64K bytes.</param>
        /// <param name="Arguments"></param>
        /// <returns>An array of values that are used as insert values in the formatted message. A %1 in the format string indicates the first value in the Arguments array; a %2 indicates the second argument; and so on.</returns>
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
