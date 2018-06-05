using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using ReadLibraryMessageTable;

namespace RlmtTest
{
    [TestClass]
    public class UnitTest1
    {
        private string SetSystemTimeExpected = "Set System Time Privilege\r\n";
        [TestMethod]
        public void ReadSetSystemTimePrivilegeMessage()
        {
            
            string Message = ReadSinglemessage(@"C:\Windows\system32\msobjs.dll", 1612);
            Assert.AreEqual<string>(SetSystemTimeExpected, Message);
        }

        [TestMethod]
        public void ReadMsgobjsMessages()
        {
            Dictionary<string, string> messageTable = new Dictionary<string, string>();
            string LibraryPath = @"C:\Windows\system32\msobjs.dll";

            messageTable = ReadAllLibraryMessages(LibraryPath);
            int messagesFound = messageTable.Keys.Count;

            Assert.IsTrue((messagesFound >= 974));

        }

        [TestMethod]
        public void TestModuleLoading()
        {
            ReadMessageTable msgTbl = new ReadMessageTable(@"C:\Windows\system32\msobjs.dll");
            string timePriv = msgTbl.ReadmoduleMessage(1612);
            msgTbl.Dispose();
            Assert.AreEqual<string>(timePriv, SetSystemTimeExpected);
        }


        private string ReadSinglemessage(string libraryPath, uint messageId)
        {
            return ReadMessageTable.ReadModuleSingleMessage(libraryPath, messageId);
        }

        private Dictionary<string,string>ReadAllLibraryMessages(string libraryPath)
        {
            return ReadMessageTable.EnumerateMessageTable(libraryPath);
        }
    }
}
