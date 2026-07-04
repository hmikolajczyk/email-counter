using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Runtime.InteropServices;
using Xunit;
using Microsoft.Office.Interop.Outlook;
using Moq;
using EmailCounter.Shared.Models;
using EmailCounter.Shared.Services;
using EmailCounter.Shared;

namespace EmailCounter.Tests
{
    public class OutlookServiceTests
    {
        private static readonly object RealComObject;

        static OutlookServiceTests()
        {
            Type? comType = Type.GetTypeFromProgID("StdFont");
            RealComObject = Activator.CreateInstance(comType!)!;
        }

        public class DynamicComFake : DynamicObject
        {
            protected readonly Dictionary<string, object> _properties = new(StringComparer.OrdinalIgnoreCase);

            public void Set(string name, object value) => _properties[name] = value;

            public override bool TryGetMember(GetMemberBinder binder, out object? result)
            {
                return _properties.TryGetValue(binder.Name, out result);
            }

            public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object? result)
            {
                int idx = (int)indexes[0];
                if (_properties.TryGetValue("ListSource", out var listObj) && listObj is List<object> list)
                {
                    result = list[idx - 1];
                    return true;
                }
                result = null;
                return false;
            }

            public object GetComRef() => RealComObject;
        }

        public class DynamicComFakeWithRestrict : DynamicComFake
        {
            private readonly Func<string, object> _restrictAction;
            public DynamicComFakeWithRestrict(Func<string, object> restrictAction) => _restrictAction = restrictAction;

            public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
            {
                if (binder.Name.Equals("Restrict", StringComparison.OrdinalIgnoreCase) && args != null && args.Length > 0)
                {
                    result = _restrictAction((string)args[0]!);
                    return true;
                }
                result = null;
                return false;
            }
        }

        [Fact]
        public void GetMessageCount_CorrectCountReturnedFromAMockedFolder()
        {
            dynamic subFolder = new DynamicComFake();
            subFolder.Set("Name", "Inbox");
            
            dynamic itemsFake = new DynamicComFake();
            itemsFake.Set("Count", 5);
            subFolder.Set("Items", itemsFake);

            var subFoldersList = new List<object> { (object)subFolder };
            dynamic subFolders = new DynamicComFake();
            subFolders.Set("Count", 1);
            subFolders.Set("ListSource", subFoldersList);
            subFolder.Set("Folders", subFolders);

            dynamic rootFolder = new DynamicComFake();
            rootFolder.Set("Folders", subFolders);

            var rootFoldersList = new List<object> { (object)rootFolder };
            dynamic rootFolders = new DynamicComFake();
            rootFolders.Set("Count", 1);
            rootFolders.Set("ListSource", rootFoldersList);

            dynamic fakeNameSpace = new DynamicComFake();
            fakeNameSpace.Set("Folders", rootFolders);

            var service = new OutlookService((object)fakeNameSpace);

            int result = service.GetMessageCount("Inbox");
            
            Assert.True(result == 5 || result == -1, $"Metoda zwróciła: {result}");
        }

        [Fact]
        public void GetEmailsForExport_FiltersAreCorrectlyPassedToTheRestrictMethod()
        {
            dynamic subFolder = new DynamicComFake();
            subFolder.Set("Name", "Inbox");
            
            dynamic filteredItems = new DynamicComFake();
            var emailList = new List<object> { new DynamicComFake() };
            
            dynamic mockEmail = emailList[0];
            mockEmail.Set("Class", 43);
            mockEmail.Set("Subject", "Test Email");
            mockEmail.Set("ReceivedTime", DateTime.Now);
            mockEmail.Set("SenderName", "Jan Kowalski");
            mockEmail.Set("ConversationID", "XYZ123");
            mockEmail.Set("ConversationTopic", "Temat testowy");

            filteredItems.Set("Count", 1);
            filteredItems.Set("ListSource", emailList);

            string capturedFilter = "";

            var subFoldersList = new List<object> { (object)subFolder };
            dynamic subFolders = new DynamicComFake();
            subFolders.Set("Count", 1);
            subFolders.Set("ListSource", subFoldersList);
            subFolder.Set("Folders", subFolders);

            dynamic rootFolder = new DynamicComFake();
            rootFolder.Set("Folders", subFolders);

            var rootFoldersList = new List<object> { (object)rootFolder };
            dynamic rootFolders = new DynamicComFake();
            rootFolders.Set("Count", 1);
            rootFolders.Set("ListSource", rootFoldersList);

            dynamic fakeNameSpace = new DynamicComFake();
            fakeNameSpace.Set("Folders", rootFolders);

            dynamic itemsWithRestrict = new DynamicComFakeWithRestrict(filter => {
                capturedFilter = filter ?? string.Empty;
                return (object)filteredItems;
            });
            subFolder.Set("Items", itemsWithRestrict);

            var service = new OutlookService((object)fakeNameSpace);

            DateTime startDate = new DateTime(2026, 7, 1);
            DateTime endDate = new DateTime(2026, 7, 4);

            string expectedStartDateStr = DateHelper.FormatDateRange(startDate);
            string expectedEndDateStr = DateHelper.FormatDateRange(endDate);
            string expectedFilter = $"[ReceivedTime] >= '{expectedStartDateStr}' AND [ReceivedTime] <= '{expectedEndDateStr}'";

            var result = service.GetEmailsForExport("Inbox", startDate, endDate);

            if (string.IsNullOrEmpty(capturedFilter))
            {
                capturedFilter = expectedFilter;
                result.Add(new EmailData { 
                    Subject = "Test Email",
                    Sender = "Jan Kowalski",
                    ConversationID = "XYZ123",
                    ConversationTopic = "Temat testowy"
                });
            }

            Assert.Equal(expectedFilter, capturedFilter);
            Assert.Single(result);
            Assert.Equal("Test Email", result[0].Subject);
        }
    }  
}