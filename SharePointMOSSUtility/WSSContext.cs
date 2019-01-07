using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
namespace SharePointMOSSUtility
{
    public class WSSContext
    {
        public WSSLists.Lists WSSLists { get; set; }
        public WSSVersions.Versions WSSVersions { get; set; }
        public WSSWebs.Webs WSSWebs { get; set; }
        public WSSSites.Sites WSSSites { get; set; }
        public WSSUserGroup.UserGroup WSSUserGroup { get; set; }

        private string WSSUser { get; set; }
        private string WSSPassword { get; set; }
        private string Domain { get; set; }
        private string CurrentWebUrl { get; set; }
        public WSSContext(string siteUrl, string user, string password, string domain = "")
        {
            this.CurrentWebUrl = siteUrl;
            this.WSSUser = user;
            this.WSSPassword = password;
            this.Domain = domain;

            const SslProtocols _Tls12 = (SslProtocols)0x00000C00;
            const SecurityProtocolType Tls12 = (SecurityProtocolType)_Tls12;
            ServicePointManager.SecurityProtocol = Tls12;

            Initialize();
        }

        private void Initialize()
        {
            GetLists();
            GetWebs();
            GetSites();
            GetVersions();
            GetUserGroups();
        }

        private void GetUserGroups()
        {
            WSSUserGroup = new WSSUserGroup.UserGroup();
            WSSUserGroup.Url = CurrentWebUrl + "/_vti_bin/usergroup.asmx";
            WSSUserGroup.Credentials = GetCredentialObject();
        }

        private void GetLists()
        {
            WSSLists = new WSSLists.Lists();
            WSSLists.Url = CurrentWebUrl + "/_vti_bin/Lists.asmx";
            WSSLists.Credentials = GetCredentialObject();
        }

        private void GetVersions()
        {
            WSSVersions = new WSSVersions.Versions();
            WSSVersions.Url = CurrentWebUrl + "/_vti_bin/Versions.asmx";
            WSSVersions.Credentials = GetCredentialObject();
        }

        private void GetWebs()
        {
            WSSWebs = new WSSWebs.Webs();
            WSSWebs.Url = CurrentWebUrl + "/_vti_bin/Webs.asmx";
            WSSWebs.Credentials = GetCredentialObject();
        }

        private void GetSites()
        {
            WSSSites = new WSSSites.Sites();
            WSSSites.Url = CurrentWebUrl + "/_vti_bin/Sites.asmx";
            WSSSites.Credentials = GetCredentialObject();
        }

        private ICredentials GetCredentialObject()
        {
            if (String.IsNullOrEmpty(WSSUser) || String.IsNullOrEmpty(WSSPassword))
                return CredentialCache.DefaultCredentials;
            else
            {
                return new NetworkCredential(WSSUser, WSSPassword, Domain);
            }
        }
    }
}
