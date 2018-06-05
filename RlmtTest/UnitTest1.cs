using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using ReadLibraryMessageTable;

namespace RlmtTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void ReadSetSystemTimePrivilegeMessage()
        {
            string Expected = "Set System Time Privilege\r\n";
            string Message = ReadSinglemessage(@"C:\Windows\system32\msobjs.dll", 1612);
            Assert.AreEqual<string>(Expected, Message);
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
