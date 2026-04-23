using System.Collections.Generic;

namespace EmailCounter.Shared.Models
{
    public class OutlookFolder
    {
        public string FolderName { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        
        public List<OutlookFolder> SubFolders { get; set; } = new List<OutlookFolder>();
    }
}