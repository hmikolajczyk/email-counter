using EmailCounter.Shared.Models;
using System.Runtime.InteropServices;

namespace EmailCounter.Shared.Services
{
    public class OutlookService
    {
        private dynamic? _outlookApp;
        private dynamic? _ns;

        public OutlookService()
        {
            try 
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType != null)
                {
                    _outlookApp = Activator.CreateInstance(outlookType);
                    _ns = _outlookApp!.GetNamespace("MAPI");
                    _ns!.Logon("", "", false, false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Nie udało się połączyć z Outlookiem: {ex.Message}");
            }
        }
        private void ForceKillOutlookIfBackground()
        {
            try
            {
                var processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                foreach (var process in processes)
                {
                    if (process.MainWindowHandle == IntPtr.Zero)
                    {
                        process.Kill();
                    }
                }
            }
            catch { }
        }
        public void TerminateBackgroundOutlookProces()
        {
            try
            {
                if (_ns != null) 
                {
                    try { _ns.Logoff(); } catch { }
                    Marshal.ReleaseComObject(_ns);
                    _ns = null;
                }

                if (_outlookApp != null)
                {
                    try 
                    { 
                        if (_outlookApp.Explorers.Count == 0)
                        {
                            _outlookApp.Quit(); 
                        }
                    } 
                    catch { }
                    
                    Marshal.ReleaseComObject(_outlookApp);
                    _outlookApp = null;
                        }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Cleanup error: {ex.Message}");
            }
            finally
            {
                for (int i = 0; i < 2; i++)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                try
                {
                    var processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                    foreach (var p in processes)
                    {
                        if (p.MainWindowHandle == IntPtr.Zero) 
                        {
                            p.Kill();
                        }
                    }
                }
                catch { }
            }
        }
        
        public List<OutlookFolder> GetFolders()
        {
            var folders = new List<OutlookFolder>();
            if (_ns == null) return folders;

            dynamic? allFolders = null;
            try {
                allFolders = _ns.Folders;
                int count = allFolders.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic folder = allFolders[i];
                    try {
                        folders.Add(MapFolder(folder));
                    } finally {
                        Marshal.ReleaseComObject(folder);
                    }
                }
            } finally {
                if (allFolders != null) Marshal.ReleaseComObject(allFolders);
            }
            return folders;
        }

        public List<EmailData> GetEmailsForExport(string folderName, DateTime startDate, DateTime endDate)
        {
            var results = new List<EmailData>();
            string startDateStr = DateHelper.FormatDateRange(startDate);
            string endDateStr = DateHelper.FormatDateRange(endDate);
            string filter = $"[ReceivedTime] >= '{startDateStr}' AND [ReceivedTime] <= '{endDateStr}'";
            
            dynamic? targetFolder = null;
            dynamic? rootFolders = null;

            try 
            {
                if (folderName.Contains("\\")) 
                {
                    targetFolder = GetFolderByPath(folderName);
                }
                else 
                {
                    rootFolders = _ns!.Folders;
                    int rootCount = rootFolders.Count;
                    for (int i = 1; i <= rootCount; i++) 
                    {
                        dynamic rootFolder = rootFolders[i];
                        dynamic? subFolders = null;
                        try 
                        {
                            subFolders = rootFolder.Folders;
                            int subCount = subFolders.Count;
                            for (int j = 1; j <= subCount; j++) 
                            {
                                dynamic subFolder = subFolders[j];
                                if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase)) 
                                {
                                    targetFolder = subFolder;
                                    break;
                                }
                                Marshal.ReleaseComObject(subFolder);
                            }
                        } 
                        finally 
                        {
                            if (subFolders != null) Marshal.ReleaseComObject(subFolders);
                            if (targetFolder == null) Marshal.ReleaseComObject(rootFolder);
                        }
                        if (targetFolder != null) break;
                    }
                }

                if (targetFolder != null) 
                {
                    dynamic? allItems = null;
                    dynamic? restrictedItems = null;
                    try 
                    {
                        allItems = targetFolder.Items;
                        restrictedItems = allItems.Restrict(filter);
                        int count = restrictedItems.Count;

                        for (int i = 1; i <= count; i++) 
                        {
                            dynamic? item = null;
                            try 
                            {
                                item = restrictedItems[i];
                                if (item.Class == 43) 
                                {
                                    results.Add(new EmailData {
                                        Subject = item.Subject,
                                        ReceivedTime = item.ReceivedTime,
                                        Sender = item.SenderName,
                                        ConversationID = item.ConversationID,
                                        ConversationTopic = item.ConversationTopic
                                    });
                                }
                            } 
                            finally 
                            {
                                if (item != null) Marshal.ReleaseComObject(item);
                            }
                        }
                    } 
                    finally 
                    {
                        if (restrictedItems != null) Marshal.ReleaseComObject(restrictedItems);
                        if (allItems != null) Marshal.ReleaseComObject(allItems);
                    }
                }
            } 
            catch (Exception ex) 
            {
                System.Diagnostics.Debug.WriteLine($"Błąd: {ex.Message}");
            } 
            finally 
            {
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
                if (rootFolders != null) Marshal.ReleaseComObject(rootFolders);
            }
            return results;
        }
        
        private dynamic? GetFolderByPath(string fullPath)
        {
            string path = fullPath.TrimStart('\\');
            string[] parts = path.Split('\\');

            try
            {
                dynamic rootFolders = _ns!.Folders;
                dynamic currentFolder = rootFolders[parts[0]];
                
                Marshal.ReleaseComObject(rootFolders);

                for (int i = 1; i < parts.Length; i++)
                {
                    dynamic nextFolders = currentFolder.Folders;
                    dynamic nextFolder = nextFolders[parts[i]];
                    
                    Marshal.ReleaseComObject(currentFolder);
                    Marshal.ReleaseComObject(nextFolders);
                    
                    currentFolder = nextFolder;
                }
                return currentFolder;
            }
            catch { return null; }
        }
        public int GetMessageCount(string folderName, DateTime? startDate=null, DateTime? endDate=null)
        {
            dynamic? rootFolders = null;
            try {
                rootFolders = _ns!.Folders;
                foreach (dynamic rootFolder in rootFolders)
                {
                    dynamic? subFolders = null;
                    try {
                        subFolders = rootFolder.Folders;
                        foreach (dynamic subFolder in subFolders)
                        {
                            try {
                                if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                                {
                                    dynamic? items = null;
                                    try {
                                        items = subFolder.Items;
                                        if (startDate.HasValue && endDate.HasValue)
                                        {
                                            string startDateStr = DateHelper.FormatDateRange(startDate);
                                            string endDateStr = DateHelper.FormatDateRange(endDate);
                                            string filter = $"[ReceivedTime] >= '{startDateStr}' AND [ReceivedTime] <= '{endDateStr}'";
                                            dynamic? filteredItems = null;
                                            try {
                                                filteredItems = items.Restrict(filter);
                                                return filteredItems.Count;
                                            } finally {
                                                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                                            }
                                        }
                                        return items.Count;
                                    } finally {
                                        if (items != null) Marshal.ReleaseComObject(items);
                                    }
                                }
                            } finally {
                                Marshal.ReleaseComObject(subFolder);
                            }
                        }
                    } finally {
                        if (subFolders != null) Marshal.ReleaseComObject(subFolders);
                        Marshal.ReleaseComObject(rootFolder);
                    }
                }
            } catch {
                return -1;
            } finally {
                if (rootFolders != null) Marshal.ReleaseComObject(rootFolders);
            }
            return -1;
        }

        private OutlookFolder MapFolder(dynamic outlookFolder)
        {
            var folder = new OutlookFolder {
                FolderName = outlookFolder.Name,
                FullPath = outlookFolder.FolderPath,
                SubFolders = new List<OutlookFolder>()
            };

            dynamic? subFolders = null;
            try {
                subFolders = outlookFolder.Folders;
                int count = subFolders.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic sub = subFolders[i];
                    try {
                        folder.SubFolders.Add(MapFolder(sub));
                    } finally {
                        Marshal.ReleaseComObject(sub);
                    }
                }
            } finally {
                if (subFolders != null) Marshal.ReleaseComObject(subFolders);
            }
            return folder;
        }
        public void Logout()
        {
            _ns?.Logoff();
        }
    }
}