using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration;

namespace SharePointMOSSUtility
{
    public sealed class SPListItemClassMap : ClassMap<SPListItem>
    {
        public SPListItemClassMap()
        {
            Map(m => m.Title).Name("Title");
            Map(m => m.FileDirRef).Name("FileDirRef");
            Map(m => m.FileRef).Name("FileRef");
            Map(m => m.FileRef).Name("ModifiedDate");
            Map(m => m.FileRef).Name("FileLength");
            Map(m => m.FileRef).Name("FilePath");
            Map(m => m.FileRef).Name("FilePathLength");
        }
    }
    public sealed class SPSubWebSizeClassMap : ClassMap<SPListItem>
    {
        public SPSubWebSizeClassMap()
        {
            Map(m => m.FileDirRef).Name("Url");
            Map(m => m.FileRef).Name("WebSize");
        }
    }
    public sealed class SPLongListItemClassMap : ClassMap<SPLongListItem>
    {
        public SPLongListItemClassMap()
        {
            Map(m => m.WebUrl).Name("WebUrl");
            Map(m => m.ListName).Name("ListName");
            Map(m => m.FilePath).Name("FilePath");
            Map(m => m.FilePathLength).Name("FilePathLength");
        }
    }

    public sealed class SPListModifiedClassMap : ClassMap<SPListModified>
    {
        public SPListModifiedClassMap()
        {
            Map(m => m.Url).Name("Url");
            Map(m => m.ListName).Name("ListName");
            Map(m => m.ListLastModifiedDate).Name("ListLastModifiedDate");
            Map(m => m.ListItemLastModifiedDate).Name("ListItemLastModifiedDate");
        }
    }

    public sealed class SPUserListItemClassMap : ClassMap<SPUserListItem>
    {
        public SPUserListItemClassMap()
        {
            Map(m => m.Title).Name("Title");
            Map(m => m.Name).Name("Name");
            Map(m => m.EMail).Name("EMail");
            Map(m => m.SipAddress).Name("SipAddress");
            Map(m => m.UserName).Name("UserName");
        }
    }

    public sealed class SPListThresholdLimitClassMap : ClassMap<SPListThresholdLimit>
    {
        public SPListThresholdLimitClassMap()
        {
            Map(m => m.ID).Name("ID");
            Map(m => m.FileDirRef).Name("FileDirRef");
            Map(m => m.ItemCount).Name("ItemCount");
        }
    }

    public sealed class SPListInformationPolicyClassMap : ClassMap<SPListInformationPolicy>
    {
        public SPListInformationPolicyClassMap()
        {
            Map(m => m.WebUrl).Name("WebUrl");
            Map(m => m.ListName).Name("ListName");
            Map(m => m.ContentTypeName).Name("ContentTypeName");
            Map(m => m.ContentTypeId).Name("ContentTypeId");
            Map(m => m.PolicyName).Name("PolicyName");
            Map(m => m.PolicyDescription).Name("PolicyDescription");
            Map(m => m.PolicyStatement).Name("PolicyStatement");
            Map(m => m.PolicyItemName).Name("PolicyItemName");
            Map(m => m.PolicyItemFeatureId).Name("PolicyItemFeatureId");
            Map(m => m.PolicyItemDescription).Name("PolicyItemDescription");
            Map(m => m.PolicyCustomData).Name("PolicyCustomData");
        }
    }
}
