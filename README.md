# SharePoint-MOSS-Utility

Traditional reporting tools are good for discovery purposes but only at a high level. This custom console application provides a great extention and provides you with in-depth analysis of sites, webs and libraries which will be useful for auditing your existing SharePoint platform. 
This particularly works with legacy SharePoint environments such as 2003 or 2007 and even 2010.

The tool helps to provide you with an in-depth analysis for the mentioned below.

1) <b> GetAllWebsTemplates() </b>: Gets all associated web sites templates.
2) <b> GetAllWebSizes() </b>: Gets sizes for each webs including subsites by looping through every document library and calculating the file sizes. There is a logic to process versionings too but is commented out due to performance issues.
3) <b> GetLongFileUrls() </b>: Gets all URLs exceeding more than 260 characters.
4) <b> GetListsLastModifiedDates() </b>: Gets the last modified entries for each list present in a site. Provides good information on which lists or libraries are old.
5) <b> GetUserInformationList() </b>: Gets  user information list present on a site.
6) <b> GetThresholdLimitForEveryFolder("doc library name") </b>: Gets the threshold limit for each and every folder present in the doc library. This gives a greater insight into which folders are exceeding the threshold limit of 5000 rather than the document library as a whole.
7) <b> GetContentTypePolicies() </b>: Gets the policies attached to the content types.
8) <b> GetListsInformationPolicies() </b>: Gets all information policies associated to the lists.
