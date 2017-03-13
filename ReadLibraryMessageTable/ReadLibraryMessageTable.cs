namespace ReadLibraryMessageTable
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Runtime.InteropServices;

    public class ReadMessageTable
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
        /// Instantiates a new instance of blah blah 
        /// TODO:DOC
        /// </summary>
        /// <param name="ModulePath"></param>
        public ReadMessageTable(string ModulePath)
        {
            IntPtr hModule = NativeMethods.LoadLibrary(ModulePath);
            if (hModule == IntPtr.Zero)
            {
                int LastError = Marshal.GetLastWin32Error();
                throw new DllNotFoundException(string.Format("Error loading library. Error code returned:{0}", LastError));
            } 
            else
            {
                this.moduleHandle = hModule;
            }
        } // public ReadMessageTable(string ModulePath)

        public static Dictionary<string, string> EnumerateMessageTable(string ModulePath)
        {
            int LastError = -1;
            Dictionary<string, string> Messages = new Dictionary<string, string>();

            // steps overview:
            // 1) load the library (into memory)
            // 2) locate the resource type in the memory (we're interested in a Message Table)
            // 3) Get a handle that can be used to get a pointer
            // 4) get the above mentioned pointer to the first byte of the resource.
            // 5) get a count for the number of blocks
            // 6) for each block, read the 

            // Loads the specified module into the address space of the calling process. The specified module may cause other modules to be loaded.
            IntPtr hModule = NativeMethods.LoadLibrary(ModulePath);
            LastError = Marshal.GetLastWin32Error();
            if (hModule == IntPtr.Zero)
            {
                 
                Console.WriteLine("Error loading library. Error code returned:{0}", LastError);
                return Messages;
            }

            IntPtr msgTable = LoadMessageTableResource(hModule);

            if (msgTable == IntPtr.Zero)
            {
                return Messages;
            }


            // Retrieves a pointer to the specified resource in memory. 
            // MSDN Remarks on LockResource function:
            // The pointer returned by LockResource is valid until the module containing the resource is unloaded. 
            //      It is not necessary to unlock resources because the system automatically deletes them when the process that created them terminates.
            // Do not try to lock a resource by using the handle returned by the FindResource or FindResourceEx function.Such a handle points to random data.
            // Note  LockResource does not actually lock memory; it is just used to obtain a pointer to the memory containing the resource data.
            //      The name of the function comes from versions prior to Windows XP, when it was used to lock a global memory block allocated by LoadResource.
            IntPtr memTable = NativeMethods.LockResource(msgTable);
            LastError = Marshal.GetLastWin32Error();
            if (memTable == IntPtr.Zero)
            {
                
                Console.WriteLine("Error locking message table in memory. Error code returned:{0}", LastError);
                return null;
            }

            // memTable is a pointer to a MESSAGE_RESOURCE_DATA structure,
            // which is just a count of the blocks in it and an array of MESSAGE_RESOURCE_BLOCK structures.
            // this code just reads the number of blocks, skips over the int (4 bytes) and starts processing each block.

            int numberOfBlocks = Marshal.ReadInt32(memTable);
            if (numberOfBlocks == 0)
            {
                Console.WriteLine("Zero entries found in message table.");
                return null;
            }
            
            // skip past the integerread above.
            IntPtr blockPtr = IntPtr.Add(memTable, 4);

            // get the size of block in bytes to increment the IntPtr by.
            int blockSize = Marshal.SizeOf<MESSAGE_RESOURCE_BLOCK>();

            // loop over all of the blocks
            for (int i = 0; i < numberOfBlocks; i++)
            {
                // https://msdn.microsoft.com/en-us/library/windows/desktop/ms648020(v=vs.85).aspx
                // Contains information about message strings with identifiers in the range indicated by the LowId and HighId members. 
                var block = Marshal.PtrToStructure<MESSAGE_RESOURCE_BLOCK>(blockPtr);

                // skip over the bytes in the block as indicated to get to the message_block_entry structure
                IntPtr entryPtr = IntPtr.Add(memTable, block.OffsetToEntries);

                // iterate over all of the entries in the block
                for (int id = block.LowId; id <= block.HighId; id++)
                {
                    var entry = Marshal.PtrToStructure<MESSAGE_RESOURCE_ENTRY>(entryPtr);

                    var messagePtr = new IntPtr(entryPtr.ToInt32() + Marshal.SizeOf(entryPtr));
                    var MessageData = Marshal.PtrToStructure<MESSAGE_RESOURCE_ENTRY>(messagePtr);

                    byte[] foo = new byte[entry.Length];
                    IntPtr textData = IntPtr.Add(entryPtr, 4);
                    //Marshal.Copy(textData, foo, 0, entry.Length);
                    string Message = Marshal.PtrToStringAuto(textData);

                    // pointer arithmetic.
                    // The length, in bytes, of the MESSAGE_RESOURCE_ENTRY structure. 
                    var length = Marshal.ReadInt16(entryPtr);
                    // Indicates that the string is encoded in Unicode, if equal to the value 0x0001. 
                    // Indicates that the string is encoded in ANSI, if equal to the value 0x0000. 
                    var flags = Marshal.ReadInt16(entryPtr, 2);
                    // Pointer to an array that contains the error message or message box display text. 
                    IntPtr textPtr = IntPtr.Add(entryPtr, 4);

                    var testText = string.Empty;
                    var text = "<<Message Unreadable>>";
                    if (flags == 0)
                    {
                        text = Marshal.PtrToStringAnsi(textPtr);
                        //testText = Marshal.PtrToStringAnsi(entry.Text);
                    }
                    else if (flags == 1)
                    {
                        text = Marshal.PtrToStringUni(textPtr);
                        //testText = Marshal.PtrToStringUni(entry.Text);
                    }
                    text = text.Replace("\r\n", "");
                    Messages.Add(id.ToString(), text);

                    // skip to next entry.
                    entryPtr = IntPtr.Add(entryPtr, length);
                }
                // skip to next block.
                blockPtr = IntPtr.Add(blockPtr, blockSize);
            } // for (int i = 0; i < numberOfBlocks; i++)

            // to implement?
            // unlock resource?
            // unload Library?

            NativeMethods.FreeLibrary(hModule);

            return Messages;
        } // public static Dictionary<string,string> EnumerateMessageTable(

        /// <summary>
        /// Using the default language, searches the module for the message ID and returns it.
        /// </summary>
        /// <param name="MessageId">Message ID to search for.</param>
        /// <returns>string resource found. if nothing then empty string.</returns>
        public string ReadmoduleMessage(uint MessageId)
        {
            IntPtr stringBuffer = IntPtr.Zero;
            int returnVal = NativeMethods.FormatMessage(NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_FROM_HMODULE | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_ALLOCATE_BUFFER,
                this.moduleHandle,
                MessageId,
                0, // default language
                ref stringBuffer,
                0,
                IntPtr.Zero);

            if (returnVal == 0)
            {
                int errorCode = Marshal.GetLastWin32Error();
                Console.WriteLine("unable to retrieve message, error code returned:{0}", errorCode);
                return string.Empty;
            }

            string messageString = Marshal.PtrToStringAnsi(stringBuffer).Replace("\r\n", "");

            // Free the buffer.
            NativeMethods.HeapFree(hDefaultProcessHeap, 0, stringBuffer);
            return messageString;
        } // public string ReadmoduleMessage(uint MessageId)

        /// <summary>
        /// Looks up a message Id in the module (dll or exe) specified. If nothing found then returns an empty string.
        /// </summary>
        /// <param name="ModulePath">Path to the module containing the messages</param>
        /// <param name="MessageID">message ID to look up</param>
        /// <returns>Message string found. if nothing found then an empty string.</returns>
        public static string ReadModuleSingleMessage (string ModulePath, uint MessageID)
        {
            IntPtr stringBuffer = IntPtr.Zero;
            IntPtr hDefaultProcessHeap = NativeMethods.GetProcessHeap();

            IntPtr hModule = NativeMethods.LoadLibrary(ModulePath);
            if (hModule == IntPtr.Zero)
            {
                int LastError = Marshal.GetLastWin32Error();
                Console.WriteLine("Error loading library. Error code returned:{0}", LastError);
                return string.Empty;
            }
            int returnVal = NativeMethods.FormatMessage(NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_FROM_HMODULE | NativeMethods.FormatMessageFlags.FORMAT_MESSAGE_ALLOCATE_BUFFER,
                hModule,
                MessageID,
                0, // default language
                ref stringBuffer,
                0,
                IntPtr.Zero);

            if (returnVal == 0)
            {
                int errorCode = Marshal.GetLastWin32Error();
                Console.WriteLine ("unable to retrieve message, error code returned:{0}", errorCode);
                return string.Empty;
            }
            string messageString = Marshal.PtrToStringAnsi(stringBuffer).Replace("\r\n","");

            // Free the buffer.
            int heapFreeStatus = NativeMethods.HeapFree(hDefaultProcessHeap, 0, stringBuffer);
            // unload the library
            NativeMethods.FreeLibrary(hModule);
            return messageString;
        } // public static string ReadModuleMessage (string ModulePath, uint MessageID)

        /// <summary>
        /// loads the module, locates the first message table and obtains a handle to it.
        /// </summary>
        /// <param name="ModulePath"></param>
        /// <returns>empty pointer on any error. otherwise handle to message table</returns>
        private static IntPtr LoadMessageTableResource (IntPtr hModule)
        {
            IntPtr MessageTablePointer = IntPtr.Zero;

            // Determines the location of a resource with the specified type and name in the specified module.
            IntPtr msgTableInfo = NativeMethods.FindResource(hModule, 1, RT_MESSAGETABLE);
            if (msgTableInfo == IntPtr.Zero)
            {
                int LastError = Marshal.GetLastWin32Error();
                Console.WriteLine("Error finding message table in library. Error code returned:{0}", LastError);
                return MessageTablePointer;
            }
            // Retrieves a handle that can be used to obtain a pointer to the first byte of the specified resource in memory.
            MessageTablePointer = NativeMethods.LoadResource(hModule, msgTableInfo);
            if (MessageTablePointer == IntPtr.Zero)
            {
                int LastError = Marshal.GetLastWin32Error();
                Console.WriteLine("Error loading message table from library. Error code returned:{0}", LastError);
                return MessageTablePointer;
            }
            return MessageTablePointer;
        } // private static IntPtr LoadMessageTableResource (IntPtr hModule)
    }
}
