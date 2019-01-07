using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointMOSSUtility
{
    
    public class SPListModified
    {
        public string Url { get; set; }
        public string ListName { get; set; }
        public DateTime ListLastModifiedDate { get; set; }
        public DateTime ListItemLastModifiedDate { get; set; }
    }
}
