using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Globalization;
using System.IO;

namespace SharePointMOSSUtility
{
    public class Program
    {

        //private const string siteUrl = @"http://togethernet/test22/vmtest";
        private const string siteUrl = @"http://tgniteim";
        private const string userName = "munagava";
        private const string passWord = "Summer2019";
        private const string domain = "OGE";
        private static WSSContext SPContext = new WSSContext(siteUrl, userName, passWord, domain);

        static void Main(string[] args)
        {
            //GetAllWebsTemplates();
            //GetAllWebSizes();

            //GetLongFileUrls();           

            //GetListsLastModifiedDates();

            GetUserInformationList();

            //FixLongFileURLs();


            //GetThresholdLimitForEveryFolder("Shared Documents");

            //GetContentTypePolicies();

            //GetListsInformationPolicies();
        }

        public static void GetUserPermissions()
        {
            try
            {
                List<SPLongListItem> results = new List<SPLongListItem>();
            }
            catch (Exception ex)
            {

            }
        }

        private static void GetAllWebsTemplates()
        {
            try
            {
                List<SPSubWebSize> subWebSizes = new List<SPSubWebSize>();
                var allWebs = GetAllWebUrls();
                foreach (var webUrl in allWebs)
                {
                    try
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"Processing {webUrl} ....");

                        SPContext = new WSSContext(webUrl, userName, passWord, domain);
                        XmlNode xn = SPContext.WSSWebs.GetWeb(webUrl);
                        //SPContext.WSSSites.
                        var elements = xn.GetChildElements();

                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("-------------------- ERROR IN GetListItems --------------------");
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.InnerException);
                    }
                }
                string relativeURL = siteUrl.Replace("http://", "");
                var dirInfo = Directory.CreateDirectory(@"D:\Results\" + relativeURL);

                CsvWriterHelper.WriteCsvRecords(subWebSizes, Path.Combine(dirInfo.FullName, "WebSizes.csv"));
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.Write("EXCEPTION: " + ex.Message);
                Console.ReadLine();
            }
        }

        private static void GetLongFileUrls()
        {
            try
            {
                List<SPLongListItem> results = new List<SPLongListItem>();
                var allWebs = GetAllWebUrls();
                foreach (var webUrl in allWebs)
                {
                    try
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"Processing {webUrl} ....");

                        SPContext = new WSSContext(webUrl, userName, passWord, domain);
                        var libs = GetDocLists();
                        foreach (var libName in libs)
                        {
                            var listItems = GetListItems(libName);
                            var longFileListItems = listItems.Where(a => a.FilePathLength >= 245);
                            foreach (var longFileListItem in longFileListItems)
                            {
                                Console.ForegroundColor = ConsoleColor.Magenta;
                                Console.WriteLine($"Long File URL {longFileListItem.FilePath} \t : {longFileListItem.FilePathLength}....");
                                results.Add(new SPLongListItem()
                                {
                                    WebUrl = webUrl,
                                    ListName = libName,
                                    FilePath = longFileListItem.FilePath,
                                    FilePathLength = longFileListItem.FilePathLength
                                });
                            }
                        }
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Processed {webUrl}");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("-------------------- ERROR IN GetListItems --------------------");
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.InnerException);
                    }
                }
                string fileName = siteUrl.Replace("http://", "").Replace("/", "-") + "-LongFileUrls.csv";
                CsvWriterHelper.WriteCsvRecords(results, System.IO.Path.Combine(@"D:\Results", fileName));
            }
            catch (Exception ex)
            {
                Console.Write("EXCEPTION: " + ex.Message);
                Console.ReadLine();
            }
        }

        public static void GetAllWebSizes()
        {
            try
            {
                List<SPSubWebSize> subWebSizes = new List<SPSubWebSize>();
                var allWebs = GetAllWebUrls();
                foreach (var webUrl in allWebs)
                {
                    try
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"Processing {webUrl} ....");

                        SPContext = new WSSContext(webUrl, userName, passWord, domain);
                        decimal webSizeInMB = 0;
                        var libs = GetDocLists();
                        //List<SPListItem> tempLibSizes = new List<SPListItem>();

                        foreach (var libName in libs)
                        {
                            var listItems = GetListItems(libName);
                            //tempLibSizes.AddRange(listItems);
                            decimal webSizeInBytes = listItems.Sum(a => a.FileLength);
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Web {webUrl} \t  {libName} \t Size:{decimal.Round(webSizeInBytes / (1024 * 1024), 2, MidpointRounding.AwayFromZero)}....");
                            webSizeInMB += decimal.Round(webSizeInBytes / (1024 * 1024), 2, MidpointRounding.AwayFromZero);
                        }
                        subWebSizes.Add(new SPSubWebSize()
                        {
                            Url = webUrl,
                            WebSize = webSizeInMB
                        });
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Processed {webUrl}    Size: {webSizeInMB}");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("-------------------- ERROR IN GetListItems --------------------");
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.InnerException);
                    }
                }
                string relativeURL = siteUrl.Replace("http://", "");
                var dirInfo = Directory.CreateDirectory(@"D:\Results\" + relativeURL);

                CsvWriterHelper.WriteCsvRecords(subWebSizes, Path.Combine(dirInfo.FullName, "WebSizes.csv"));
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.Write("EXCEPTION: " + ex.Message);
                Console.ReadLine();
            }
        }

        public static void GetListsLastModifiedDates()
        {
            List<SPListModified> results = new List<SPListModified>();
            var allWebs = GetAllWebUrls();
            foreach (var webUrl in allWebs)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Processing {webUrl} ....");

                    SPContext = new WSSContext(webUrl, userName, passWord, domain);
                    //Get Lists
                    XmlNode xn = SPContext.WSSLists.GetListCollection();
                    var elements = xn.GetChildElements();
                    //filter out reserved or hidden lists
                    var lists = elements.Where(e => e.Attribute("Title").Value != "User Information List" &&
                                e.Attribute("Title").Value != "Master Page Gallery" &&
                                e.Attribute("Title").Value != "Content and Structure Reports" &&
                                e.Attribute("Hidden").Value != "True");

                    foreach (var list in lists)
                    {
                        var listTitle = list.Attribute("Title").Value;
                        var listModifiedDate = DateTime.ParseExact(list.Attribute("Modified").Value, "yyyyMMdd hh:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None);
                        var lastListItemModifiedDate = DateTime.MinValue;
                        int itemCount = int.Parse(list.Attribute("ItemCount").Value);
                        if (itemCount > 0)
                        {
                            lastListItemModifiedDate = GetLastModifiedListItem(listTitle);
                        }
                        results.Add(new SPListModified()
                        {
                            Url = webUrl,
                            ListName = listTitle,
                            ListLastModifiedDate = listModifiedDate,
                            ListItemLastModifiedDate = lastListItemModifiedDate
                        });
                    }
                    string fileName = siteUrl.Replace("http://", "") + "-LastModifiedDate.csv";
                    CsvWriterHelper.WriteCsvRecords(results, System.IO.Path.Combine(@"D:\Results", "bpi-PM-lastmodified.csv"));
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("-------------------- ERROR IN GetListItems --------------------");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.InnerException);
                }
            }
        }

        public static void GetUserInformationList()
        {
            try
            {
                var results = new List<SPUserListItem>();
                //Get Lists
                XmlNode xn = SPContext.WSSLists.GetListCollection();
                var elements = xn.GetChildElements();
                //filter out reserved or hidden lists
                var userInfoListTitle = elements
                    .Where(e => e.Attribute("Title").Value == "User Information List")
                    .Select(e => e.Attribute("Title").Value).First();

                xn = SPContext.WSSLists.GetList(userInfoListTitle);
                elements = xn.GetChildElements();

                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options
                ndViewFields.InnerXml = "<FieldRef Name='Title' /> " +
                              "<FieldRef Name='Name' /> <FieldRef Name = 'EMail' />" +
                              "<FieldRef Name='SipAddress' /> <FieldRef Name = 'UserName' />";

                ndQueryOptions.InnerXml = "<ViewAttributes Scope='RecursiveAll' IncludeRootFolder='True' />";
                XmlNode ndListItems = null;

                bool flag;
                do
                {
                    flag = false;
                    ndListItems = SPContext.WSSLists.GetListItems(userInfoListTitle, null, ndQuery, ndViewFields, "500", ndQueryOptions, null);
                    if (ndListItems != null)
                    {
                        XmlNode xmlPosition = ndListItems.SelectSingleNode("//@ListItemCollectionPositionNext");
                        foreach (XmlNode node in ndListItems.ChildNodes)
                        {
                            if (node.Name == "rs:data")
                            {
                                //rs:row
                                foreach (XmlNode childNode in node.ChildNodes)
                                {
                                    if (childNode.Name == "z:row")
                                    {
                                        XmlNodeReader objReader = new XmlNodeReader(childNode);
                                        while (objReader.Read())
                                        {
                                            //SPList Item
                                            results.Add(new SPUserListItem()
                                            {

                                                EMail = (objReader["ows_EMail"] == null) ? "" : objReader["ows_EMail"].ToString(),
                                                Name = (objReader["ows_Name"] == null) ? "" : objReader["ows_Name"].ToString(),
                                                SipAddress = (objReader["ows_SipAddress"] == null) ? "" : objReader["ows_SipAddress"].ToString(),
                                                UserName = (objReader["ows_UserName"] == null) ? "" : objReader["ows_UserName"].ToString(),
                                                Title = (objReader["ows_Title"] == null) ? "" : objReader["ows_Title"].ToString()
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        if (xmlPosition != null)
                        {
                            ndQueryOptions.InnerXml = "<Paging ListItemCollectionPositionNext='" + xmlPosition.InnerXml + "' /><MeetingInstanceID>-1</MeetingInstanceID><ViewAttributes Scope='RecursiveAll'  IncludeRootFolder='True' />";
                            flag = true;
                        }
                    }
                } while (flag);

                string relativeURL = siteUrl.Replace("http://", "");
                var dirInfo = Directory.CreateDirectory(@"D:\Results\" + relativeURL);

                CsvWriterHelper.WriteCsvRecords(results, Path.Combine(dirInfo.FullName, "tgniteim-is-UserInfoList.csv"));
            }
            catch (Exception ex)
            {
                Console.Write("EXCEPTION: " + ex.Message);
                Console.ReadLine();
            }
        }

        public static IEnumerable<string> GetAllWebUrls()
        {
            try
            {
                XmlNode xn = SPContext.WSSWebs.GetAllSubWebCollection();
                //XmlNode xn = SPContext.WSSWebs.GetWebCollection();
                var elements = xn.GetChildElements();
                var currentWebUrl = SPContext.WSSWebs.Url.Replace("/_vti_bin/Webs.asmx", "");
                var webs = elements.Where(e => e.Attribute("Url").Value.StartsWith(currentWebUrl)).Select(e => e.Attribute("Url").Value);
                return webs;
            }
            catch (FaultException fe)
            {
                MessageFault mf = fe.CreateMessageFault();
                if (mf.HasDetail)
                {
                    XmlElement fexe = mf.GetDetail<XmlElement>();
                    Console.WriteLine("\tError: " + fexe.OuterXml);
                }
                throw fe;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static IEnumerable<string> GetDocLists()
        {
            try
            {
                //Get Lists
                XmlNode xn = SPContext.WSSLists.GetListCollection();
                var elements = xn.GetChildElements();
                //filter out reserved or hidden lists
                return elements
                    .Where(e => e.Attribute("ServerTemplate").Value == "101" ||
                                e.Attribute("ServerTemplate").Value == "115" ||
                                e.Attribute("ServerTemplate").Value == "109"
                            /* &&
                            e.Attribute("Title").Value != "Style Library" &&
                            e.Attribute("Title").Value != "Content and Structure Reports" &&
                            e.Attribute("Title").Value != "Workflow Tasks" &&
                            e.Attribute("Title").Value != "Site Collection Documents" &&
                            e.Attribute("Title").Value != "Form Templates" &&
                            e.Attribute("Title").Value != "Reusable Content" &&
                            e.Attribute("Title").Value != "Site Assets"*/)
                    .Select(e => e.Attribute("Title").Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void GetThresholdLimitForEveryFolder(string listName)
        {
            var results = new List<SPListThresholdLimit>();
            var listItems = GetListItems(listName);
            var items = listItems.GroupBy(f => f.FileDirRef).Select(
                group => new
                {
                    FileDirRef = group.Key,
                    ItemsCount = group.Count()
                });

            foreach (var item in items)
            {
                results.Add(new SPListThresholdLimit()
                {
                    FileDirRef = item.FileDirRef,
                    ItemCount = item.ItemsCount
                });
            }
            var dirInfo = Directory.CreateDirectory(@"D:\Results\");
            CsvWriterHelper.WriteCsvRecords(results, Path.Combine(dirInfo.FullName, "ESSBUAuditSharedDocuments.csv"));
        }

        

        private static List<string> GetAllFolders(string listName)
        {
            try
            {
                var results = new List<string>();
                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options
                ndViewFields.InnerXml = "<FieldRef Name='EncodedAbsUrl' /> " +
                              "<FieldRef Name='LinkFilename' /> <FieldRef Name = 'Author' />" +
                              "<FieldRef Name='Editor' /> <FieldRef Name = 'Created' />" +
                              "<FieldRef Name='Modified' /> <FieldRef Name = 'ID' /> <FieldRef Name = 'Title' />" +
                              "<FieldRef Name='FileRef' /> <FieldRef Name = 'FileDirRef' /> <FieldRef Name='File_x0020_Size' />";

                ndQueryOptions.InnerXml = "<ViewAttributes Scope='RecursiveAll' IncludeRootFolder='True' />";
                ndQuery.InnerXml = @"<Where> 
                                <Eq>
                                <FieldRef Name='FSObjType' />
                                <Value Type='Lookup'>1</Value>
                                </Eq>
                            </Where>";
                XmlNode ndListItems = null;
                bool flag;
                do
                {
                    flag = false;
                    ndListItems = SPContext.WSSLists.GetListItems(listName, null, ndQuery, ndViewFields, "500", ndQueryOptions, null);
                    if (ndListItems != null)
                    {
                        XmlNode xmlPosition = ndListItems.SelectSingleNode("//@ListItemCollectionPositionNext");
                        foreach (XmlNode node in ndListItems.ChildNodes)
                        {
                            if (node.Name == "rs:data")
                            {
                                XmlNodeReader objReader = new XmlNodeReader(node);
                                while (objReader.Read())
                                {
                                    //SPList Item
                                    if (objReader["ows_EncodedAbsUrl"] != null && objReader["ows_LinkFilename"] != null)
                                    {
                                        //var fileRef = objReader["ows_FileRef"];
                                        var fileRef = objReader["ows_FileRef"].Split(new char[] { '#' })[1];
                                        results.Add(objReader["ows_ID"].ToString());
                                    }
                                }
                            }
                        }
                        if (xmlPosition != null)
                        {
                            ndQueryOptions.InnerXml = "<Paging ListItemCollectionPositionNext='" + xmlPosition.InnerXml + "' /><MeetingInstanceID>-1</MeetingInstanceID><ViewAttributes Scope='RecursiveAll'  IncludeRootFolder='True' />";
                            flag = true;
                        }
                    }
                } while (flag);
                return results;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DateTime GetLastModifiedListItem(string listName)
        {
            var results = new List<SPListItem>();
            try
            {
                //Get List's Regional Setting
                XmlNode xn = SPContext.WSSLists.GetList(listName);
                var elements = xn.GetChildElements();
                var regionalSettingsElement = elements.FirstOrDefault(e => e.Name.LocalName.Equals("RegionalSettings"));
                var timeZone = regionalSettingsElement.Elements().Where(e => e.Name.LocalName == "TimeZone").Single().Value;
                var utcOffset = new TimeSpan(0, int.Parse(timeZone), 0);
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().Where(t => t.BaseUtcOffset.Equals(utcOffset)).First();

                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options

                ndQuery.InnerXml = "<OrderBy><FieldRef Name='Modified' Ascending='FALSE'></FieldRef></OrderBy>";
                ndViewFields.InnerXml = "<FieldRef Name='EncodedAbsUrl' /> " +
                              "<FieldRef Name='LinkFilename' /> <FieldRef Name = 'Author' />" +
                              "<FieldRef Name='Editor' /> <FieldRef Name = 'Created' />" +
                              "<FieldRef Name='Modified' /> <FieldRef Name = 'ID' /> <FieldRef Name = 'Title' />";

                XmlNode ndListItems = null;
                ndListItems = SPContext.WSSLists.GetListItems(listName, null, ndQuery, ndViewFields, "1", ndQueryOptions, null);
                DateTime modifiedDate = DateTime.MinValue;
                if (ndListItems != null)
                {
                    foreach (XmlNode node in ndListItems.ChildNodes)
                    {
                        if (node.Name == "rs:data")
                        {
                            var listData = node.GetChildElements();
                            string lastModifiedDate = listData.Attributes().Where(e => e.Name.LocalName.Equals("ows_Modified")).Single().Value;
                            modifiedDate = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(lastModifiedDate), timeZoneInfo);
                        }
                    }
                }
                return modifiedDate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }        

        public static List<SPListItem> GetListItems(string listName)
        {
            var results = new List<SPListItem>();
            try
            {
                //Get List's Regional Setting
                XmlNode xn = SPContext.WSSLists.GetList(listName);
                var elements = xn.GetChildElements();
                var regionalSettingsElement = elements.FirstOrDefault(e => e.Name.LocalName.Equals("RegionalSettings"));
                //https://stackoverflow.com/questions/4265766/is-there-a-way-to-get-a-sharepoint-sites-locale-with-web-services
                var timeZone = regionalSettingsElement.Elements().Where(e => e.Name.LocalName == "TimeZone").Single().Value;
                var utcOffset = new TimeSpan(0, int.Parse(timeZone), 0);
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().Where(t => t.BaseUtcOffset.Equals(utcOffset)).First();

                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options

                
                ndViewFields.InnerXml = "<FieldRef Name='EncodedAbsUrl' /> " +
                              "<FieldRef Name='LinkFilename' /> <FieldRef Name = 'Author' />" +
                              "<FieldRef Name='Editor' /> <FieldRef Name = 'Created' />" +
                              "<FieldRef Name='Modified' /> <FieldRef Name = 'ID' /> <FieldRef Name = 'Title' />" +
                              "<FieldRef Name='FileRef' /> <FieldRef Name = 'FileDirRef' /> <FieldRef Name='File_x0020_Size' />";


                ndQueryOptions.InnerXml = "<ViewAttributes Scope='RecursiveAll' IncludeRootFolder='True' />";

                
                XmlNode ndListItems = null;
                bool flag;
                do
                {
                    flag = false;
                    ndListItems = SPContext.WSSLists.GetListItems(listName, null, ndQuery, ndViewFields, "500", ndQueryOptions, null);
                    if (ndListItems != null)
                    {
                        XmlNode xmlPosition = ndListItems.SelectSingleNode("//@ListItemCollectionPositionNext");
                        foreach (XmlNode node in ndListItems.ChildNodes)
                        {
                            if (node.Name == "rs:data")
                            {
                                XmlNodeReader objReader = new XmlNodeReader(node);
                                while (objReader.Read())
                                {
                                    //SPList Item
                                    if (objReader["ows_EncodedAbsUrl"] != null && objReader["ows_LinkFilename"] != null)
                                    {
                                        Int64 fileLength = 0;
                                        //if it is a file, then calculate fileLength
                                        if (objReader["ows_FSObjType"].Split(new char[] { '#' })[1] == "0")
                                        {
                                            fileLength = Int64.Parse(objReader["ows_File_x0020_Size"].Split(new char[] { '#' })[1]);
                                            //fileLength = GetVersions(objReader["ows_EncodedAbsUrl"]);
                                        }
                                        results.Add(new SPListItem()
                                        {
                                            FileDirRef = objReader["ows_FileDirRef"].Split(new char[] { '#' })[1],
                                            FileRef = objReader["ows_FileRef"].ToString(),
                                            ID = objReader["ows_ID"],
                                            ModifiedDate = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(objReader["ows_Modified"].ToString()), timeZoneInfo),
                                            Title = objReader["ows_Title"],
                                            FileLength = fileLength,
                                            FilePath = Uri.UnescapeDataString(objReader["ows_EncodedAbsUrl"].ToString()),
                                            FilePathLength = Uri.UnescapeDataString(objReader["ows_EncodedAbsUrl"].ToString()).Length
                                        });

                                    }
                                }
                            }
                        }
                        if (xmlPosition != null)
                        {
                            ndQueryOptions.InnerXml = "<Paging ListItemCollectionPositionNext='" + xmlPosition.InnerXml + "' /><MeetingInstanceID>-1</MeetingInstanceID><ViewAttributes Scope='RecursiveAll'  IncludeRootFolder='True' />";
                            flag = true;
                        }
                    }
                } while (flag);
                return results;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static List<SPListItem> GetListItemsBasedOnLastModifiedDate(string listName)
        {
            var results = new List<SPListItem>();
            try
            {
                //Get List's Regional Setting
                XmlNode xn = SPContext.WSSLists.GetList(listName);
                var elements = xn.GetChildElements();
                var regionalSettingsElement = elements.FirstOrDefault(e => e.Name.LocalName.Equals("RegionalSettings"));
                //https://stackoverflow.com/questions/4265766/is-there-a-way-to-get-a-sharepoint-sites-locale-with-web-services
                var timeZone = regionalSettingsElement.Elements().Where(e => e.Name.LocalName == "TimeZone").Single().Value;
                var utcOffset = new TimeSpan(0, int.Parse(timeZone), 0);
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().Where(t => t.BaseUtcOffset.Equals(utcOffset)).First();

                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options

                ndQuery.InnerXml = "<Where>" +
                    "<Geq>" +
                    "<FieldRef Name='Modified'/><Value Type='DateTime' IncludeTimeValue='FALSE'>2018-12-10</Value>" +
                    "</Geq>" +
                    "</Where>";


                ndViewFields.InnerXml = "<FieldRef Name='EncodedAbsUrl' /> " +
                              "<FieldRef Name='LinkFilename' /> <FieldRef Name = 'Author' />" +
                              "<FieldRef Name='Editor' /> <FieldRef Name = 'Created' />" +
                              "<FieldRef Name='Modified' /> <FieldRef Name = 'ID' /> <FieldRef Name = 'Title' />" +
                              "<FieldRef Name='FileRef' /> <FieldRef Name = 'FileDirRef' /> <FieldRef Name='File_x0020_Size' />";

                ndQueryOptions.InnerXml = "<ViewAttributes Scope='RecursiveAll' IncludeRootFolder='True' />";
                XmlNode ndListItems = null;
                bool flag;
                do
                {
                    flag = false;
                    ndListItems = SPContext.WSSLists.GetListItems(listName, null, ndQuery, ndViewFields, "500", ndQueryOptions, null);
                    if (ndListItems != null)
                    {
                        XmlNode xmlPosition = ndListItems.SelectSingleNode("//@ListItemCollectionPositionNext");
                        foreach (XmlNode node in ndListItems.ChildNodes)
                        {
                            if (node.Name == "rs:data")
                            {
                                XmlNodeReader objReader = new XmlNodeReader(node);
                                while (objReader.Read())
                                {
                                    //SPList Item
                                    if (objReader["ows_EncodedAbsUrl"] != null && objReader["ows_LinkFilename"] != null)
                                    {
                                        Int64 fileLength = 0;
                                        //if it is a file, then calculate fileLength
                                        if (objReader["ows_FSObjType"].Split(new char[] { '#' })[1] == "0")
                                        {
                                            fileLength = Int64.Parse(objReader["ows_File_x0020_Size"].Split(new char[] { '#' })[1]);
                                            //fileLength = GetVersions(objReader["ows_EncodedAbsUrl"]);
                                        }
                                        results.Add(new SPListItem()
                                        {
                                            FileDirRef = objReader["ows_FileDirRef"].Split(new char[] { '#' })[1],
                                            FileRef = objReader["ows_FileRef"].ToString(),
                                            ID = objReader["ows_ID"],
                                            ModifiedDate = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(objReader["ows_Modified"].ToString()), timeZoneInfo),
                                            Title = objReader["ows_Title"],
                                            FileLength = fileLength,
                                            FilePath = Uri.UnescapeDataString(objReader["ows_EncodedAbsUrl"].ToString()),
                                            FilePathLength = Uri.UnescapeDataString(objReader["ows_EncodedAbsUrl"].ToString()).Length
                                        });

                                    }
                                }
                            }
                        }
                        if (xmlPosition != null)
                        {
                            ndQueryOptions.InnerXml = "<Paging ListItemCollectionPositionNext='" + xmlPosition.InnerXml + "' /><MeetingInstanceID>-1</MeetingInstanceID><ViewAttributes Scope='RecursiveAll'  IncludeRootFolder='True' />";
                            flag = true;
                        }
                    }
                } while (flag);
                return results;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static SPListItem GetListItem(string listName, string fileUrl)
        {
            SPListItem results = null;
            try
            {


                //Get List's Regional Setting
                XmlNode xn = SPContext.WSSLists.GetList(listName);
                var elements = xn.GetChildElements();
                var regionalSettingsElement = elements.FirstOrDefault(e => e.Name.LocalName.Equals("RegionalSettings"));
                //https://stackoverflow.com/questions/4265766/is-there-a-way-to-get-a-sharepoint-sites-locale-with-web-services
                var timeZone = regionalSettingsElement.Elements().Where(e => e.Name.LocalName == "TimeZone").Single().Value;
                var utcOffset = new TimeSpan(0, int.Parse(timeZone), 0);
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().Where(t => t.BaseUtcOffset.Equals(utcOffset)).First();

                //XML Document object
                XmlDocument xmlDoc = new System.Xml.XmlDocument();
                XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");//Query
                XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");//Views fields
                XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");//Options
                ndViewFields.InnerXml = "<FieldRef Name='EncodedAbsUrl' /> " +
                              "<FieldRef Name='LinkFilename' /> <FieldRef Name = 'Author' />" +
                              "<FieldRef Name='Editor' /> <FieldRef Name = 'Created' />" +
                              "<FieldRef Name='Modified' /> <FieldRef Name = 'ID' /> <FieldRef Name = 'Title' />" +
                              "<FieldRef Name='FileRef' /> <FieldRef Name = 'FileDirRef' /> <FieldRef Name='File_x0020_Size' />";

                ndQueryOptions.InnerXml = "<ViewAttributes Scope='RecursiveAll' IncludeRootFolder='True' />";
                ndQuery.InnerXml = string.Format("<Where><Eq><FieldRef Name='{0}'/>" +
                                                "<Value Type='Text'>{1}</Value></Eq></Where>", "EncodedAbsUrl", Uri.EscapeUriString(fileUrl));
                XmlNode ndListItems = null;
                bool flag;
                do
                {
                    flag = false;
                    ndListItems = SPContext.WSSLists.GetListItems(listName, null, ndQuery, ndViewFields, "1", ndQueryOptions, null);
                    if (ndListItems != null)
                    {
                        XmlNode xmlPosition = ndListItems.SelectSingleNode("//@ListItemCollectionPositionNext");
                        foreach (XmlNode node in ndListItems.ChildNodes)
                        {
                            if (node.Name == "rs:data")
                            {
                                XmlNodeReader objReader = new XmlNodeReader(node);
                                while (objReader.Read())
                                {
                                    //SPList Item
                                    if (objReader["ows_EncodedAbsUrl"] != null && objReader["ows_LinkFilename"] != null)
                                    {
                                        Int64 fileLength = 0;
                                        //if it is a file, then calculate fileLength
                                        if (objReader["ows_FSObjType"].Split(new char[] { '#' })[1] == "0")
                                        {
                                            fileLength = Int64.Parse(objReader["ows_File_x0020_Size"].Split(new char[] { '#' })[1]);
                                            //fileLength = GetVersions(objReader["ows_EncodedAbsUrl"]);
                                            results = new SPListItem()
                                            {
                                                FileDirRef = objReader["ows_FileDirRef"].ToString(),
                                                FileRef = objReader["ows_FileRef"].ToString(),
                                                ID = objReader["ows_ID"],
                                                ModifiedDate = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(objReader["ows_Modified"].ToString()), timeZoneInfo),
                                                Title = objReader["ows_Title"],
                                                FileLength = fileLength,
                                                FilePath = objReader["ows_EncodedAbsUrl"].ToString(),
                                                FilePathLength = objReader["ows_EncodedAbsUrl"].ToString().Length
                                            };
                                        }
                                    }
                                }
                            }
                        }
                        if (xmlPosition != null)
                        {
                            ndQueryOptions.InnerXml = "<Paging ListItemCollectionPositionNext='" + xmlPosition.InnerXml + "' /><MeetingInstanceID>-1</MeetingInstanceID><ViewAttributes Scope='RecursiveAll'  IncludeRootFolder='True' />";
                            flag = true;
                        }
                    }
                } while (flag);
                return results;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static Int64 GetVersions(string fileUrl)
        {
            //Get List's Regional Setting
            XmlNode xn = SPContext.WSSVersions.GetVersions(fileUrl);
            var versionElement = xn.GetChildElements();
            Int64 fileSize = 0;
            if (versionElement != null)
            {
                //Get child element with name 'result' and whose attribute named 'version' contains '@'
                var rows = from row in versionElement
                           where row.Name.LocalName == "result" &&
                           row.Attribute("version").Value.Contains("@")
                           select row;

                int totalVersions = rows.Count();

                foreach (var row in rows)
                {
                    //var versionSize= decimal.Round(decimal.Parse(row.Attribute("size").Value) / (1024 * 1024), 2, MidpointRounding.AwayFromZero);
                    //Console.ForegroundColor = ConsoleColor.Magenta;
                    //Console.WriteLine($"{row.Attribute("version").Value} --- {fileUrl} --- {row.Attribute("size").Value}");
                    fileSize += Int64.Parse(row.Attribute("size").Value);
                }
                //CurrentFileVersion = CurrentFileVersion.Replace("@", string.Empty);                
            }
            return fileSize;
        }        

        private static string GetWebUrl(string fileUrl)
        {
            var requestUri = new Uri(fileUrl);
            var returnUrl = string.Empty;

            var baseUrl = requestUri.GetLeftPart(UriPartial.Authority);
            for (int i = requestUri.Segments.Length; i >= 0; i--)
            {
                var path = string.Join(string.Empty, requestUri.Segments.Take(i));
                string url = string.Format("{0}{1}", baseUrl, path);
                try
                {
                    SPContext = new WSSContext(url, userName, passWord, domain);
                    XmlNode node = SPContext.WSSWebs.GetWeb(url);
                    var childElements = node.GetChildElements();
                    returnUrl = url;
                }
                catch (Exception ex)
                {
                    break;
                }
            }
            return returnUrl;
        }

        private static void GetContentTypePolicies()
        {
            XmlNode xn = SPContext.WSSWebs.GetContentTypes();
            var elements = xn.GetChildElements();
            var ctypes = elements
               .Where(e => e.Attribute("Group") != null && e.Attribute("Group").Value.ToLower() != "_hidden");

            foreach (var contentType in ctypes)
            {
                xn = SPContext.WSSWebs.GetContentType(contentType.Attribute("ID").Value);
                elements = xn.GetChildElements();

                XNamespace p = "office.server.policy";
                var policies = elements.Descendants(p + "Policy");
                foreach (var policy in policies)
                {
                    var policyName = policy.Element(p + "Name").Value;
                    var policyDescription = policy.Element(p + "Description").Value;
                    var policyStatement = policy.Element(p + "Statement").Value;

                    //get Policy Items
                    var policyItems = policy.Descendants(p + "PolicyItem");
                    foreach (var policyItem in policyItems)
                    {
                        Console.WriteLine(policyItem.Attribute("featureId").Value);
                        var policyItemName = policyItem.Element(p + "Name").Value;
                        var policyItemDescription = policyItem.Element(p + "Description").Value;
                        var policyCustomData = policyItem.Element(p + "CustomData").ToString();

                        Console.WriteLine(policyItem);
                    }
                }

            }
        }

        private static void GetListsInformationPolicies()
        {
            var results = new List<SPListInformationPolicy>();
            var allWebs = GetAllWebUrls();
            foreach (var webUrl in allWebs)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Processing {webUrl} ....");

                    try
                    {
                        SPContext = new WSSContext(webUrl, userName, passWord, domain);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write("EXCEPTION: " + ex.Message);
                        Console.Write("Continuiing....");
                        continue;
                    }

                    //Get Lists
                    XmlNode xn = SPContext.WSSLists.GetListCollection();
                    var elements = xn.GetChildElements();

                    var listTitles = elements
                            .Where(e =>
                                    e.Attribute("Title").Value != "Style Library" &&
                                    e.Attribute("Title").Value != "Content and Structure Reports" &&
                                    e.Attribute("Title").Value != "Workflow Tasks" &&
                                    e.Attribute("Title").Value != "Site Collection Documents" &&
                                    e.Attribute("Title").Value != "Form Templates" &&
                                    e.Attribute("Title").Value != "Reusable Content" &&
                                    e.Attribute("Title").Value != "Site Assets")
                            .Select(e => e.Attribute("Title").Value);

                    foreach (var listTitle in listTitles)
                    {
                        //Get List's Regional Setting

                        xn = SPContext.WSSLists.GetListContentTypes(listTitle, null);
                        elements = xn.GetChildElements();
                        foreach (var contentType in elements)
                        {
                            XNamespace p = "office.server.policy";
                            var policies = contentType.Descendants(p + "Policy");
                            foreach (var policy in policies)
                            {
                                var contentTypeName = contentType.Attribute("Name").Value;
                                var contentTypeId = contentType.Attribute("ID").Value;
                                var policyName = policy.Element(p + "Name").Value;
                                var policyDescription = policy.Element(p + "Description").Value;
                                var policyStatement = policy.Element(p + "Statement").Value;

                                //get Policy Items
                                var policyItems = policy.Descendants(p + "PolicyItem");
                                foreach (var policyItem in policyItems)
                                {
                                    var policyItemFeatureId = policyItem.Attribute("featureId").Value;
                                    var policyItemName = policyItem.Element(p + "Name").Value;
                                    var policyItemDescription = policyItem.Element(p + "Description").Value;
                                    var policyCustomData = policyItem.Element(p + "CustomData").ToString();

                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine(policyItem);


                                    results.Add(new SPListInformationPolicy()
                                    {
                                        WebUrl = webUrl,
                                        ListName = listTitle,
                                        ContentTypeName = contentTypeName,
                                        ContentTypeId = contentTypeId,
                                        PolicyName = policyName,
                                        PolicyDescription = policyDescription,
                                        PolicyStatement = policyStatement,
                                        PolicyItemName = policyItemName,
                                        PolicyItemFeatureId = policyItemFeatureId,
                                        PolicyItemDescription = policyItemDescription,
                                        PolicyCustomData = policyCustomData
                                    });
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.Write("EXCEPTION: " + ex.Message);
                    Console.ReadLine();
                }
            }
            string fileName = siteUrl.Replace("http://", "").Replace("/", "-") + "-InfoPolicies.csv";
            CsvWriterHelper.WriteCsvRecords(results, System.IO.Path.Combine(@"D:\Results", fileName));

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("DONE");
            Console.ReadLine();
        }
    }
}
