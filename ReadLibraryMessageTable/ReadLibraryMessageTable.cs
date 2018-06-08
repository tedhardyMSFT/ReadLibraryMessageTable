namespace ReadLibraryMessageTable
{

    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Security.Permissions;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.ComponentModel;

    [Serializable]
    public class ReadLibraryException : Exception
    {
        /// <summary>
        /// Creates a new ReadLibraryException
        /// </summary>
        /// <param name="Message">Message describing the exception.</param>
        public ReadLibraryException(string Message) : base(Message)
        {
            Win32ErrorCode = -1;
            Win32ErrorMessage = null;
        }

        /// <summary>
        /// Creates a new ReadLibraryException with Win32Error code
        /// </summary>
        /// <param name="Message">Message describing the exception.</param>
        /// <param name="Win32Error">Native Windows API Error code returned</param>
        public ReadLibraryException(string Message, int Win32Error) : base(Message)
        {
            Win32ErrorCode = Win32Error;
            Win32ErrorMessage = new Win32Exception(Win32ErrorCode).Message;
        }


        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }

        /// <summary>
        /// Returns a string representation of the current ReadLibraryException object
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (Win32ErrorCode == -1)
            {
                return base.ToString();
            } else
            {
                return base.ToString() + string.Format("\nWin32 Error code: {0}\nWin32 Message:{1}", this.Win32ErrorCode, this.Win32ErrorMessage);
            }
        }

        /// <summary>
        /// Win32 error code returned from the underlying API.
        /// </summary>
        public Int32 Win32ErrorCode { get; private set; }

        /// <summary>
        /// Localized message for the supplied win32 error code.
        /// </summary>
        public string Win32ErrorMessage { get; private set; }
    } // public class ReadLibraryException : Exception

    /// <summary>
    /// Enables reading a Library or Executable message table contents.
    /// </summary>
    public class ReadMessageTable : IDisposable
    {
        /// <summary>
        /// Resource type for Message Table
        /// See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms648009(v=vs.85).aspx
        /// </summary>
        private const int RT_MESSAGETABLE = 11;

        /// <summary>
        /// Contains information about message strings with identifiers in the range indicated by the LowId and HighId members. 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        struct MESSAGE_RESOURCE_BLOCK
        {
            /// <summary>
            /// The lowest message identifier contained within this structure. 
            /// </summary>
            public int LowId;
            /// <summary>
            /// The highest message identifier contained within this structure. 
            /// </summary>
            public int HighId;
            /// <summary>
            /// The offset, in bytes, from the beginning of the MESSAGE_RESOURCE_DATA structure to the MESSAGE_RESOURCE_ENTRY structures in this MESSAGE_RESOURCE_BLOCK.
            /// The MESSAGE_RESOURCE_ENTRY structures contain the message strings.
            /// </summary>
            public int OffsetToEntries;
        }

        /// <summary>
        /// Contains the error message or message box display text for a message table resource. 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        struct MESSAGE_RESOURCE_ENTRY
        {
            /// <summary>
            /// The length, in bytes, of the MESSAGE_RESOURCE_ENTRY structure
            /// </summary>
            
            public Int16 Length;
            /// <summary>
            /// Indicates that the string is encoded in Unicode, if equal to the value 0x0001. 
            /// Indicates that the string is encoded in ANSI, if equal to the value 0x0000. 
            /// </summary>
            
            public Int16 Flags;
            /// <summary>
            /// Pointer to an array that contains the error message or message box display text. 
            /// </summary>
            public byte Text;
        }

        /// <summary>
        /// handle to the module loaded in the constructor.
        /// </summary>
        private IntPtr moduleHandle = IntPtr.Zero;

        /// <summary>
        /// used for HeapFree. Handle to the default process heap.
        /// </summary>
        private IntPtr hDefaultProcessHeap = NativeMethods.GetProcessHeap();

        /// <summary>
        /// Instantiates a new instance of ReadMessageTable class.
        /// When finished, use the FreeMessageTablefile
        /// </summary>
        /// <param name="ModulePath"></param>
        public ReadMessageTable(string ModulePath)
        {
            IntPtr hModule = NativeMethods.LoadLibraryEx(ModulePath,IntPtr.Zero, NativeMethods.LOAD_LIBRARY_AS_DATAFILE);
            int LastError = Marshal.GetLastWin32Error();
            if (hModule == IntPtr.Zero)
            {
                throw new ReadLibraryException(string.Format("Error loading library from location:{0}", ModulePath), LastError);
            } 
            else
            {
                this.moduleHandle = hModule;
            }
        } // public ReadMessageTable(string ModulePath)

        /// <summary>
        /// Disposes of the external resources and file handles.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
        } //public void FreeMessageTableFile()

        /// <summary>
        /// Disposes of unmanaged resources.
        /// </summary>
        /// <param name="freeAll">if true both managed and unmanaged resources will be freed.</param>
        protected virtual void Dispose(bool freeAll)
        {
            if (this.moduleHandle != IntPtr.Zero)
            {
                NativeMethods.FreeLibrary(this.moduleHandle);
            }
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Reads the message table for the library referenced by ModulePath and returns all of the messages found.
        /// </summary>
        /// <param name="ModulePath">Full path to the library</param>
        /// <returns></returns>
        public static Dictionary<string, string> EnumerateMessageTable(string ModulePath)
        {
            Dictionary<string, string> Messages = new Dictionary<string, string>();
            int LastError = -1;
            // steps overview - on any error throw ReadLibraryException otherwise continue to next step:
            // 1) load the library (into memory as a datafile - so no executable code)
            // 2) locate the resource type in the memory (we're interested in a Message Table)
            // 3) Get a handle that can be used to get a pointer
            // 4) get the above mentioned pointer to the first byte of the resource.
            // 5) get a count for the number of blocks
            // 6) for each block, read the ranges in the block

            // Loads the specified module into the address space of the calling process. The specified module may cause other modules to be loaded.
            // 1) load the library (into memory as a datafile - so no executable code)
            IntPtr hModule = NativeMethods.LoadLibraryEx(ModulePath, IntPtr.Zero, NativeMethods.LOAD_LIBRARY_AS_DATAFILE);
            LastError = Marshal.GetLastWin32Error();
            if (hModule == IntPtr.Zero)
            {
                
                throw new ReadLibraryException(string.Format("Error loading library from {0}.", ModulePath), LastError);
            }

            // 2) locate the resource type in the memory (we're interested in a Message Table)
            // 3) Get a handle that can be used to get a pointer
            IntPtr msgTable = LoadMessageTableResource(hModule);

            if (msgTable == IntPtr.Zero)
            {
                // no message table resources
                return Messages;
            }

            // Retrieves a pointer to the specified resource in memory. 
            // MSDN Remarks on LockResource function:
            // The pointer returned by LockResource is valid until the module containing the resource is unloaded. 
            //      It is not necessary to unlock resources because the system automatically deletes them when the process that created them terminates.
            // Do not try to lock a resource by using the handle returned by the FindResource or FindResourceEx function.Such a handle points to random data.
            // Note  LockResource does not actually lock memory; it is just used to obtain a pointer to the memory containing the resource data.
            //      The name of the function comes from versions prior to Windows XP, when it was used to lock a global memory block allocated by LoadResource.
            
            // 4) get the above mentioned pointer to the first byte of the resource.
            IntPtr memTable = NativeMethods.LockResource(msgTable);
            LastError = Marshal.GetLastWin32Error();
            if (memTable == IntPtr.Zero)
            {
                throw new ReadLibraryException("Error locking message table in memory.", LastError);
            }

            // memTable is a pointer to a MESSAGE_RESOURCE_DATA structure,
            // which is just a count (DWORD) of the blocks in it and an array of MESSAGE_RESOURCE_BLOCK pointers.
            // this code just reads the number of blocks, skips over the int (4 bytes) and starts processing each block.
            // 5) get a count for the number of blocks
            int numberOfBlocks = Marshal.ReadInt32(memTable);
            if (numberOfBlocks == 0)
            {
                // no blocks found, return empty Dictionary
                return Messages;
            }
            
            // skip past the integer (DWORD) read above.
            IntPtr blockPtr = IntPtr.Add(memTable, 4);

            // get the size of block in bytes for incrementing the pointer
            int messageBlockSize = Marshal.SizeOf<MESSAGE_RESOURCE_BLOCK>();

            // loop over all of the blocks
            for (int i = 0; i < numberOfBlocks; i++)
            {
                // https://msdn.microsoft.com/en-us/library/windows/desktop/ms648020(v=vs.85).aspx
                // Contains information about message strings with identifiers in the range indicated by the LowId and HighId members. 
                MESSAGE_RESOURCE_BLOCK block = Marshal.PtrToStructure<MESSAGE_RESOURCE_BLOCK>(blockPtr);

                // skip over the bytes in the block as indicated to get to the message_block_entry structure
                IntPtr entryPtr = IntPtr.Add(memTable, (int)block.OffsetToEntries);

                // iterate over all of the entries in the block
                for (int id = block.LowId; id <= block.HighId; id++)
                {

                    MESSAGE_RESOURCE_ENTRY entry = Marshal.PtrToStructure<MESSAGE_RESOURCE_ENTRY>(entryPtr);

                    IntPtr messagePtr = new IntPtr(entryPtr.ToInt32() + Marshal.SizeOf(entryPtr));
                    MESSAGE_RESOURCE_ENTRY MessageData = Marshal.PtrToStructure<MESSAGE_RESOURCE_ENTRY>(messagePtr);

                    IntPtr textData = IntPtr.Add(entryPtr, 4);
                    string Message = Marshal.PtrToStringAnsi(textData);

                    // pointer arithmetic.
                    // read the length, in bytes, of the MESSAGE_RESOURCE_ENTRY structure. 
                    Int16 length = Marshal.ReadInt16(entryPtr);
                    // flags: 0x0001 Indicates that the string is encoded in Unicode
                    //        0x0000 Indicates that the string is encoded in ANSI
                    Int16 flags = Marshal.ReadInt16(entryPtr, 2);
                    // get pointer 4 bytes ahead to a pointer for the message. 
                    IntPtr textPtr = IntPtr.Add(entryPtr, 4);
                    string testText = string.Empty;
                    string text = string.Empty;
                    // check if unicode flag is set.
                    if (flags == 0)
                    {
                        text = Marshal.PtrToStringAnsi(textPtr);
                    }
                    else
                    {
                        text = Marshal.PtrToStringUni(textPtr);
                    }
                    try
                    {
                        // add ID and test to output Dictionary
                        Messages.Add(((uint)id).ToString(), text);
                    } catch ( Exception ex)
                    {
                        Console.WriteLine("exception:{0}", ex);
                    }
                    // skip to next entry.
                    entryPtr = IntPtr.Add(entryPtr, length);
                }
                // skip to next block.
                blockPtr = IntPtr.Add(blockPtr, messageBlockSize);
            } // for (int i = 0; i < numberOfBlocks; i++)

            NativeMethods.FreeLibrary(hModule);

            return Messages;
        } // public static Dictionary<string,string> EnumerateMessageTable(

        /// <summary>
        /// Using the default language, searches the module for the message ID and returns it.
        /// </summary>
        /// <param name="MessageId">Message ID to search for.</param>
        /// <returns>string resource found. if nothing then empty string.</returns>
        public string ReadModuleMessage(uint MessageId)
        {
            return ReadModuleMessage(MessageId, 0);
        }

        /// <summary>
        /// Using the specified language ID, searches the module for the message ID and returns it.
        /// </summary>
        /// <param name="MessageId">Message ID to search for.</param>
        /// <param name="LangaugeId">Language ID to use.</param>
        /// <returns>string resource found. if nothing then empty string.</returns>
        public string ReadModuleMessage(uint MessageId, uint LangaugeId)
        {
            IntPtr stringBuffer = IntPtr.Zero;
            // attempt to read the specific message from the library
            // returns zero upon error, otherwise the number of TCHARS in output buffer.
            int returnVal = NativeMethods.FormatMessage(NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_FROM_HMODULE | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_ALLOCATE_BUFFER | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_IGNORE_INSERTS,
                this.moduleHandle,
                MessageId,
                LangaugeId, // specified language
                ref stringBuffer,
                0,
                IntPtr.Zero);
            int errorCode = Marshal.GetLastWin32Error();
            // check if no characters returned (error condition)
            if (returnVal == 0)
            {
                throw new ReadLibraryException(string.Format("Unable to retrieve message id {0} for languageID {1} FormatMessage.", MessageId, LangaugeId), errorCode);
            }
            // read output buffer
            string messageString = Marshal.PtrToStringAnsi(stringBuffer);

            // Free the buffer.
            NativeMethods.HeapFree(hDefaultProcessHeap, 0, stringBuffer);
            return messageString;
        } // public string ReadmoduleMessage(uint MessageId)

        /// <summary>
        /// Looks up a message Id in the module (dll or exe) specified using the default system LanguageId. If nothing found then returns an empty string.
        /// </summary>
        /// <param name="ModulePath">Path to the module containing the messages</param>
        /// <param name="MessageID">message ID to look up</param>
        /// <returns>Message string found. if nothing found then an empty string.</returns>
        public static string ReadModuleSingleMessage(string ModulePath, uint MessageID)
        {
            return ReadModuleSingleMessage(ModulePath, MessageID, 0);
        }


        /// <summary>
        /// Looks up a message Id in the module (dll or exe) specified using the supplied LanguageId. If nothing found then returns an empty string.
        /// </summary>
        /// <param name="ModulePath">Path to the module containing the messages</param>
        /// <param name="MessageID">message ID to look up</param>
        /// <param name="LanguageId">Language ID to use</param>
        /// <returns>Message string found. if nothing found then an empty string.</returns>
        public static string ReadModuleSingleMessage (string ModulePath, uint MessageID, uint LanguageId)
        {
            IntPtr stringBuffer = IntPtr.Zero;
            IntPtr hDefaultProcessHeap = NativeMethods.GetProcessHeap();

            IntPtr hModule = NativeMethods.LoadLibraryEx(ModulePath, IntPtr.Zero, NativeMethods.LOAD_LIBRARY_AS_DATAFILE);
            if (hModule == IntPtr.Zero)
            {
                int LastError = Marshal.GetLastWin32Error();
                throw new ReadLibraryException(string.Format("Error loading library {0}.", ModulePath), LastError);
            }
            int returnVal = NativeMethods.FormatMessage(NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_FROM_HMODULE | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_ALLOCATE_BUFFER | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_IGNORE_INSERTS,
                hModule,
                MessageID,
                LanguageId,
                ref stringBuffer,
                0,
                IntPtr.Zero);
            int errorCode = Marshal.GetLastWin32Error();
            //FormatMessage returns zero on error, otherwise number of chars in buffer.
            if (returnVal == 0)
            {
                throw new ReadLibraryException(string.Format("Unable to retrieve message id {0} from library: {1}", MessageID, ModulePath), errorCode);
            }
            // read buffer
            string messageString = Marshal.PtrToStringAnsi(stringBuffer);

            // Free the buffer.
            int heapFreeStatus = NativeMethods.HeapFree(hDefaultProcessHeap, 0, stringBuffer);
            // unload the library
            NativeMethods.FreeLibrary(hModule);
            return messageString;
        } // public static string ReadModuleMessage (string ModulePath, uint MessageID)

        /// <summary>
        /// using the supplied handle to the module, locates the first message table and obtains a handle to it.
        /// </summary>
        /// <param name="hModule">handle to the module file</param>
        /// <returns>Handle to first message table</returns>
        private static IntPtr LoadMessageTableResource (IntPtr hModule)
        {
            // return value: pointer to first message table
            IntPtr MessageTablePointer = IntPtr.Zero;
            // Win32 status code of last native call.
            int LastError = 0;

            // Determines the location of a resource with the specified type and name in the specified module.
            IntPtr msgTableInfo = NativeMethods.FindResource(hModule, 1, RT_MESSAGETABLE);
            LastError = Marshal.GetLastWin32Error();
            if (msgTableInfo == IntPtr.Zero)
            {
                
                throw new ReadLibraryException("Error finding message table in library.", LastError);
            }
            // Retrieves a handle that can be used to obtain a pointer to the first byte of the specified resource in memory.
            MessageTablePointer = NativeMethods.LoadResource(hModule, msgTableInfo);
            LastError = Marshal.GetLastWin32Error();
            if (MessageTablePointer == IntPtr.Zero)
            {
                throw new ReadLibraryException("Error loading message table from library.", LastError);
            }
            return MessageTablePointer;
        } // private static IntPtr LoadMessageTableResource (IntPtr hModule)
    } // public class ReadMessageTable : IDisposable
} // namespace ReadLibraryMessageTable
