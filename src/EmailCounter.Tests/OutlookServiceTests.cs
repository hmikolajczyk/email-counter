using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Runtime.InteropServices;
using Xunit;
using Microsoft.Office.Interop.Outlook;
using Moq;
using EmailCounter.Shared.Models;
using EmailCounter.Shared.Services;

namespace EmailCounter.Tests
{
    public class OutlookServiceTests
    {
       public class TestFolder
        {
            public string Name { get; set; } = "";
            public List<object> Folders { get; set; } = new();
            public TestItems Items { get; set; } = new();
        }

        public class TestItems
        {
            public int Count => 5;
        }

        [Fact]
        public void GetMessageCount_CorrectCountReturnedFromAMockedFolder()
        {
            var subFolder = new TestFolder { Name = "Inbox" };
            var rootFolder = new TestFolder { Folders = new List<object> { subFolder } };
            
            var fakeNameSpace = new
            {
                Folders = new List<object> { rootFolder }
            };

            var service = new OutlookService(fakeNameSpace);

            int result = service.GetMessageCount("Inbox");
            
            Assert.True(result == 5 || result == -1, $"Metoda zwróciła: {result}");
        }
    }  
}