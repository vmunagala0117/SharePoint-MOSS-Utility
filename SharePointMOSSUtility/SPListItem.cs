using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointMOSSUtility
{
    public class SPListItem
    {
        public string ID { get; set; }
        public string Title { get; set; }
        public string FileDirRef { get; set; }
        public string FileRef { get; set; }
        public DateTime ModifiedDate { get; set; }
        public Int64 FileLength { get; set; }
        public string FilePath { get; set; }
        public int FilePathLength { get; set; }
    }
}
