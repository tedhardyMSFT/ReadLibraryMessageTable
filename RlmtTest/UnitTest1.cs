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
        public void TestReadSinglemessage_SetSystemTimePrivilegeMessage()
        {
            
            string Message = this.readSinglemessage(@"C:\Windows\system32\msobjs.dll", 1612);
            Assert.AreEqual<string>(SetSystemTimeExpected, Message);
        }

        [TestMethod]
        public void TestReadAllLibraryMessages_msgobs()
        {
            Dictionary<string, string> messageTable = new Dictionary<string, string>();
            string LibraryPath = @"C:\Windows\system32\msobjs.dll";

            messageTable = this.readAllLibraryMessages(LibraryPath);
            int messagesFound = messageTable.Keys.Count;

            Assert.IsTrue((messagesFound >= 974));

        }

        [TestMethod]
        [ExpectedException(typeof(ReadLibraryException))]
        public void TestReadAllLibraryMessages_FileNotExist()
        {
            Dictionary<string, string> messageTable = new Dictionary<string, string>();
            string LibraryPath = @"C:\Windows\system32\msobjs2.dll";

            messageTable = this.readAllLibraryMessages(LibraryPath);
        }

        [TestMethod]
        [ExpectedException(typeof(ReadLibraryException))]
        public void TestReadSinglemessage_InvalidMessageId()
        {
            string Message = this.readSinglemessage(@"C:\Windows\system32\msobjs.dll", 9999);
        }

        [TestMethod]
        public void TestReadModuleMessage_SetSystemTimePrivilegeMessage()
        {
            ReadMessageTable msgTbl = new ReadMessageTable(@"C:\Windows\system32\msobjs.dll");
            string timePriv = msgTbl.ReadModuleMessage(1612);
            msgTbl.Dispose();
            Assert.AreEqual<string>(SetSystemTimeExpected,timePriv);
        }



        private string readSinglemessage(string libraryPath, uint messageId)
        {
            return ReadMessageTable.ReadModuleSingleMessage(libraryPath, messageId);
        }

        private Dictionary<string,string> readAllLibraryMessages(string libraryPath)
        {
            return ReadMessageTable.EnumerateMessageTable(libraryPath);
        }
    }
}
