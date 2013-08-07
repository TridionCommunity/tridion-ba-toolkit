#define DEBUG
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using CSV;
using Tridion.ContentManager.CoreService.Client;

// uses CSV read/write class from http://www.blackbeltcoder.com/Articles/files/reading-and-writing-csv-files-in-c

namespace BA_Toolkit
{
    internal struct CoreConfig
    {
        internal string CmsUrl, User, Pass, PublicationUri, FolderUri, OutputPath;

        public CoreConfig(string cmsUrl, string user, string pass, string publicationUri, string folderUri,
                          string outputPath)
        {
            CmsUrl = cmsUrl;
            User = user;
            Pass = pass;
            PublicationUri = publicationUri;
            FolderUri = folderUri;
            OutputPath = outputPath;
        }
    }

    public class FunctionalDetails
    {
        #region Declare Variables

        private static readonly ReadOptions DefaultReadOptions = new ReadOptions();
        private static readonly List<List<String>> Publications = new List<List<String>>();
        private static readonly List<List<String>> PublicationRights = new List<List<string>>();
        private static readonly List<List<string>> Groups = new List<List<String>>();
        private static List<List<string>>[] Folders = new List<List<string>>[] {};
        private static readonly List<List<string>> FunctionalDesign = new List<List<string>>();

        private static Hashtable groupsHashtable = new Hashtable();



        private static CoreConfig _cmsCoreConfig = new CoreConfig("http://cms-url", "username",
                                                          "password", "tcm:0-4-1", "tcm:4-5-2",
                                                          "Test.csv"); // change "tcm:0-4-1" to a publication id for "010 Schemas" and change "tcm:4-5-2" to id for "Building Blocks" folder


        private static List<CoreConfig> _coreConfigs = new List<CoreConfig>();

        private static int _count;
        private static ICoreService _core;

        #endregion

        private static void Main(string[] args)
        {
            #region Initialize

            _coreConfigs.Add(_cmsCoreConfig); // 0


            const int configSelection = 0;


            var coreServiceHandler = new CoreServiceHandler();
            var config = _coreConfigs[configSelection];
            _core = coreServiceHandler.GetNewClient(config.CmsUrl, config.User, config.Pass);

            var std = Console.Out;
            Console.SetWindowSize(Console.LargestWindowWidth, Console.LargestWindowHeight);
            Console.SetOut(std);

            var orgFilter = new OrganizationalItemItemsFilterData {ItemTypes = new[] {ItemType.Folder, ItemType.StructureGroup}};

            #endregion

            Trace.WriteLine(DateTime.Now.ToString() + " - Start of Main");
            var consoleTracer = new ConsoleTraceListener(false) {Name = "mainConsoleTracer"};

            var start = DateTime.Now;
            var lastStamp = DateTime.Now;
            TimeSpan delta = lastStamp - start;

            Trace.Listeners.Add(consoleTracer);

            #region Get Core Service Results

            lastStamp = DateTime.Now;

            GetPublicationsFromCore();

            delta = DateTime.Now - lastStamp;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Publications");

            lastStamp = DateTime.Now;
            GetPublicationRightsFromCore();

            delta = DateTime.Now - lastStamp;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Rights");

            lastStamp = DateTime.Now;
            GetGroupScopeFromCore();

            delta = DateTime.Now - lastStamp;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Scope");

            var bluePrintXml = _core.GetSystemWideListXml(new BluePrintFilterData());

            Folders = new List<List<string>>[bluePrintXml.Nodes().Count()];

            var pub = 0;
            foreach (XElement pubElement in bluePrintXml.Nodes())
            {
                lastStamp = DateTime.Now;

                Folders[pub] = new List<List<string>>();

                var idAttribute = pubElement.Attribute("ID");
                if (idAttribute == null) continue;
                var publicationUri = idAttribute.Value;

                GetFolderPermissionsFromCore(orgFilter, publicationUri, pub);
                pub++;

                delta = DateTime.Now - lastStamp;
                consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Publication" +
                                        pubElement.Attribute("Title"));
            }

            #endregion

            #region Arrange Results

            lastStamp = DateTime.Now;

            FunctionalDesign.AddRange(Publications);


            delta = DateTime.Now - lastStamp;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Added Publication Range");


            lastStamp = DateTime.Now;
            FunctionalDesign.AddRange(PublicationRights);

            delta = lastStamp - start;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Added PublicationRights");

            lastStamp = DateTime.Now;
            FunctionalDesign.AddRange(Groups);

            delta = DateTime.Now - lastStamp;
            consoleTracer.WriteLine(DateTime.Now - start + ", " + delta + ",Added Groups");

            for (var i = 0; i < Folders.Count(); i++)
            {
                FunctionalDesign.AddRange(Folders[i]);
            }

            #endregion

            #region Write and Read

            using (var writer = new CsvFileWriter(_coreConfigs[configSelection].OutputPath))
                Write(FunctionalDesign, writer);

            Console.WriteLine("Saved to file {0}. Check the Debug or Release folder.\nPress any key to end trace and call dispose...", _coreConfigs[configSelection].OutputPath);
            Console.ReadKey();

            #endregion

            Trace.Flush();
            Trace.Listeners.Remove(consoleTracer);
            consoleTracer.Close();

            Trace.WriteLine(DateTime.Now.ToString() + " - End of Main");
            Trace.Close();
            coreServiceHandler.Dispose();
            Console.WriteLine("Press any key to end...", _coreConfigs[configSelection].OutputPath);
            Console.ReadKey();
        }

        private static void Write(IEnumerable<IEnumerable<string>> stringMatrix, CsvFileWriter writer)
        {
            foreach (var row in stringMatrix.Where(row => writer != null))
            {
                writer.WriteRow((List<string>) row);
            }
        }

        private static void GetPublicationsFromCore()
        {
            var bluePrintXml = _core.GetSystemWideListXml(new BluePrintFilterData());

            foreach (XElement pubElement in bluePrintXml.Nodes())
            {
                var titleAttribute = pubElement.Attribute("Title");
                if (titleAttribute == null) continue;
                var titleValue = titleAttribute.Value;
                Publications.Add(new List<string> {titleValue}); // Publications is a List<string>
            }
        }

        private static void GetPublicationRightsFromCore()
        {
            var bluePrintXml = _core.GetSystemWideListXml(new BluePrintFilterData());

            var publicationsInaRow = Publications.Select(publication => publication[0]).ToList();
            PublicationRights.Add(publicationsInaRow);
            var pubCount = publicationsInaRow.Count;

            PublicationRights[0].Add("Groups");

            var rightsTypes = Enum.GetValues(typeof (Rights));
            var rightsTypesInaRow = (from object type in rightsTypes select type.ToString()).ToList();
                // add rights across the top right of the table

            PublicationRights[0].InsertRange(pubCount + 1, rightsTypesInaRow);

            var rightsTypeCount = rightsTypes.Length;

            foreach (XElement publicationElement in bluePrintXml.Nodes())
            {
                var pubID = publicationElement.Attribute("ID");
                if (pubID == null) continue;
                var publication = _core.Read(pubID.Value, DefaultReadOptions) as PublicationData;
                if (publication == null) continue;
                var pubTitle = publication.Title;
                var publicationAclList = publication.AccessControlList.AccessControlEntries;
                var j = publicationsInaRow.IndexOf(pubTitle);

                foreach (var acl in publicationAclList)
                {
                    var outputRow = new List<string>(pubCount + rightsTypeCount);
                    for (var i = 0; i < pubCount + rightsTypeCount + 1; i++)
                    {
                        outputRow.Add("");
                    }
                    var allowedRights = acl.AllowedRights.ToString();
                    outputRow[j] = "x";

                    outputRow[pubCount] = acl.Trustee.Title;
                    var rights = allowedRights.Split(',');

                    foreach (var right in rights)
                    {
                        var lookup = rightsTypesInaRow.IndexOf(right.Trim());
                        outputRow[pubCount + 1 + lookup] = "x";
                    }

                    PublicationRights.Add(outputRow);
                }
            }
        }

        private static void GetGroupScopeFromCore()
        {
            var groupsFilterData = new GroupsFilterData();
            var groupSystemWideList = _core.GetSystemWideList(groupsFilterData);

            List<string> publicationsInaRow = Publications.Select(publication => publication[0]).ToList();
            publicationsInaRow.Insert(0, "Groups \n  Publications");
            publicationsInaRow.Insert(1, "All Publications (setting, not an actual Group)");

            Groups.Add(publicationsInaRow);

            foreach (GroupData groupData in groupSystemWideList) // loop groups
            {
                foreach (var groupMembershipData in groupData.GroupMemberships) // loop group membership scope
                {
                    var groupMembership = groupMembershipData.Group.Title;
                    var scopeCount = groupMembershipData.Scope.Count();

                    if (scopeCount == 0)
                    {
                        Groups.Add(new List<string> {groupData.Title, groupMembership});
                    }

                    foreach (LinkWithIsEditableToRepositoryData scope in groupMembershipData.Scope)
                        // multiple scopes possible
                    {
                        var row = new List<string>();

                        row.Add(groupData.Title);

                        var j = publicationsInaRow.IndexOf(scope.Title); // get Publications to match against

                        row.Add(" "); // left padding

                        row.AddRange(Publications.Select(t => ""));
                        row.Insert(j, groupMembership);
                        Groups.Add(row);
                    }
                }
            }
        }

        private static void GetFolderPermissionsFromCore(OrganizationalItemItemsFilterData orgFilter,
                                                         string publicationID, int orgItemArrayPos)
        {
            var row = new List<string>();
            var groups = _core.GetSystemWideList(new GroupsFilterData());
            var users = _core.GetSystemWideList(new UsersFilterData());

            row = GetOrgItemHeaderRow(publicationID, groups, users, row);

            Folders[orgItemArrayPos].Add(row);

            var publication = (PublicationData) _core.Read(publicationID, new ReadOptions());
            var rootFolderId = publication.RootFolder.IdRef;
            var rootSgId = publication.RootStructureGroup.IdRef;

            ProcessSubOrgItems(orgFilter, rootFolderId, 0, orgItemArrayPos);
            ProcessSubOrgItems(orgFilter, rootSgId, 0, orgItemArrayPos);
        }

        private static List<string> GetOrgItemHeaderRow(string publicationID,
                                                       IdentifiableObjectData[] groups, IdentifiableObjectData[] users,
                                                       List<string> row)
        {
            row.Add("Publication: " + _core.Read(publicationID, DefaultReadOptions).Title); // 0
            row.Add(""); // 1
            row.Add("");
            row.Add("");
            row.Add("");
            row.Add("");
            row.Add(""); // 6
            row.Add("Localize here (typically \"no\")");
            row.Add("Set permissions here (item not shared from above)? If \"no\", ignore these settings.");
            row.Add("Inherit Security Settings from Parent (ignore on shared items)?");
            row.Add("Linked Schema (default components created in this orgItem)");

            var i = 0;
            foreach (GroupData group in groups) // loop groups
            {
                row.Add(group.Title);

                if (!groupsHashtable.ContainsKey(group.Title))
                {
                    groupsHashtable.Add(group.Title, i);
                }

                i++;
            }

            foreach (UserData user in users) // loop users
            {
                row.Add(user.Title);

                if (!groupsHashtable.ContainsKey(user.Title))
                {
                    groupsHashtable.Add(user.Title, i);
                }
                i++;
            }


            return row;
        }

        private static void ProcessSubOrgItems(OrganizationalItemItemsFilterData orgItemfilter,
                                              string tcmid, int indentLevel, int orgItemArrayPos)
        {
            if (indentLevel > 5) return;
            indentLevel++;
            var orgItem = _core.Read(tcmid, DefaultReadOptions) as OrganizationalItemData;
            SaveOrgItemPermissions(orgItem, indentLevel, orgItemArrayPos, 5);

            var subItem = _core.GetListXml(tcmid, orgItemfilter);

            foreach (XElement item in subItem.Nodes())
            {
                var idAttribute = item.Attribute("ID");
                if (idAttribute == null) continue;
                ProcessSubOrgItems(orgItemfilter, idAttribute.Value, indentLevel, orgItemArrayPos);
            }
        }

        private static void SaveOrgItemPermissions(OrganizationalItemData orgItem, int indent, int folderArrayPos, int maxIndent)
        {
            var acls = orgItem.AccessControlList.AccessControlEntries;

            var row = new List<string>();

            for (var i = 0; i < maxIndent + 5; i++)
            {
                row.Add("");
            }

            _count = Folders[folderArrayPos][0].Count;
            var width = _count - maxIndent - 5;
            for (var i = 0; i < width; i++)
            {
                row.Add("");
            }

            row[indent] = orgItem.Title;
            row[maxIndent + 2] = orgItem.BluePrintInfo.IsLocalized.GetValueOrDefault() ? "Yes" : "No";
            row[maxIndent + 3] = orgItem.BluePrintInfo.IsShared.GetValueOrDefault() ? "No" : "Yes";
            row[maxIndent + 4] = orgItem.IsPermissionsInheritanceRoot.GetValueOrDefault() ? "No" : "Yes";
            
            if (orgItem is FolderData)
            {
                var folder = orgItem as FolderData;
                row[maxIndent + 5] = folder.LinkedSchema.Title;
            }
            else
            {
                row[maxIndent + 5] = "n/a";
            }

            foreach (var acl in acls)
            {
                string shortPermissions;

                switch (acl.AllowedPermissions.ToString())
                {
                    case "All":
                        shortPermissions = "All";
                        break;
                    case "Read":
                        shortPermissions = "r";
                        break;
                    case "Read, Write":
                        shortPermissions = "w";
                        break;
                    case "Read, Write, Delete":
                        shortPermissions = "d";
                        break;
                    case "Read, Write, Localize":
                        shortPermissions = "l";
                        break;
                    case "None":
                        shortPermissions = "";
                        break;
                    default:
                        shortPermissions = "n/a";
                        break;
                }



                var j = 0;
                if (groupsHashtable.ContainsKey(acl.Trustee.Title)) j = (int) groupsHashtable[acl.Trustee.Title] + 11;


                if (
                    (!orgItem.BluePrintInfo.IsShared.GetValueOrDefault() && orgItem.IsPermissionsInheritanceRoot.GetValueOrDefault()) ||
                    (orgItem.BluePrintInfo.IsShared.GetValueOrDefault() && orgItem.BluePrintInfo.IsLocalized.GetValueOrDefault() && orgItem.IsPermissionsInheritanceRoot.GetValueOrDefault())                    
                   )
                {
                    row[0] = "Set here";
                    
                }

                else // don't set here
                {
                    row[0] = "Don't set here";

                    shortPermissions = "<" + shortPermissions + ">";
                }

                //row[maxIndent + 2] = orgItem.BluePrintInfo.IsLocalized.GetValueOrDefault() ? "Yes" : "No";
                //row[maxIndent + 3] = orgItem.BluePrintInfo.IsShared.GetValueOrDefault() ? "No" : "Yes";
                //row[maxIndent + 4] = orgItem.IsPermissionsInheritanceRoot.GetValueOrDefault() ? "No" : "Yes";

                row[j] = shortPermissions;

            }
            Folders[folderArrayPos].Add(row);
        }
    }
}

