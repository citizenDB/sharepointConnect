using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;

namespace sharepointConnect
{
    public class Action
    {
        private string clientContext { get; set; }
        private string siteCollection { get; set; }
        private string siteLibrary { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Action"/> class.
        /// </summary>
        public Action(string clientContext, string siteCollection, string siteLibrary)
        {
            this.clientContext = clientContext;
            this.siteCollection = siteCollection;
            this.siteLibrary = siteLibrary;
        }

        public class Item
        {
            public int id { get; set; }
            public string documentID { get; set; }
            public string title { get; set; }
            public string filetype { get; set; }
            public string created { get; set; }
            public string modified { get; set; }
            public string url { get; set; }
        }

        Helper helper = new Helper();

        /// <summary>
        /// This function first uploads a file to the App_Data.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="columnName"></param>
        /// <param name="columnValue"></param>
        /// <param name="checkin"></param>
        /// <returns>
        /// Fields of the uploaded document
        /// </returns>
        public object upload(HttpPostedFileBase file, string columnName, string columnValue, bool checkin = false)
        {
            
            string directory = System.AppDomain.CurrentDomain.BaseDirectory + "/App_Data/";

            /// Get file name, size and MIMEType
            int fileSize = file.ContentLength;
            string mimeType = file.ContentType;
            System.IO.Stream fileContent = file.InputStream;

            /// Upload file to temp folder on server
            var fileName = Path.GetFileName(file.FileName);
            var path = Path.Combine(directory, fileName);
            file.SaveAs(path);

            using (ClientContext ctx = new ClientContext(clientContext))
            {
                try
                {
                    /// Upload the file to Sharepoint
                    List documentLibrary = ctx.Web.Lists.GetByTitle(siteLibrary);
                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Content = helper.ReadFile(path);

                    /// Overwrite existing file
                    newFile.Overwrite = true;
                    newFile.Url = siteCollection + fileName.Trim();
                    Microsoft.SharePoint.Client.File uploadFile = documentLibrary.RootFolder.Files.Add(newFile);

                    /// Get Sharepoint specific meta data
                    uploadFile.ListItemAllFields["Title"] = fileName;
                    uploadFile.ListItemAllFields[columnName] = columnValue;
                    uploadFile.ListItemAllFields.Update();

                    // Check-in file
                    if (checkin)
                    {
                        uploadFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    }

                    ctx.Load(uploadFile, f => f.ListItemAllFields);
                    ctx.ExecuteQuery();

                    /// Create list with Sharepoint fields
                    var item = uploadFile.ListItemAllFields;

                    var list = new List<Item>();

                    list.Add(new Item()
                    {
                        id = uploadFile.ListItemAllFields.Id,
                        documentID = (string)item["_dlc_DocId"],
                        title = Convert.ToString(item["Title"]),
                        url = helper.hostURL(clientContext) + Convert.ToString(item["FileRef"]),
                        filetype = Convert.ToString(item["File_x0020_Type"]),
                        created = "",
                        modified = "",
                    });

                    return list;
                }
                catch (Exception ex)
                {
                    return "[]";
                }
                finally
                {
                    /// Aways delete the file from local web server
                    System.IO.File.Delete(path);
                }
            };

        }

        /// <summary>
        /// This function deletes a file from Sharepoint based on the document ID / Library Name
        /// </summary>
        /// <param name="id"></param>
        /// <returns>
        /// Returns the ID of a deleted document to use in the front-end (json)
        /// </returns>
        public string delete(string id, string libraryName)
        {
            using (ClientContext ctx = new ClientContext(clientContext))
            {
                string spLibraryName = libraryName;

                /// Delete document
                var list = ctx.Web.Lists.GetByTitle(spLibraryName);
                var listItem = list.GetItemById(id);
                listItem.DeleteObject();
                ctx.ExecuteQuery();
            }

            return id;
        }

        /// <summary>
        /// This function gets the documents from Sharepoint based on the library name, Fieldref and Lookup value
        /// </summary>
        /// <param name="spLibrary"></param>
        /// <param name="columnName"></param>
        /// <param name="columnValue"></param>
        /// <param name="rowLimit"></param>
        /// <returns>List of documentswith fields</returns>
        public object get(string spLibrary, string columnName, string columnValue, int rowLimit = 100)
        {
            ClientContext ctx = new ClientContext(clientContext);

            // Library Name
            List oList = ctx.Web.Lists.GetByTitle(spLibrary);

            CamlQuery camlQuery = new CamlQuery();

            /// Sharepoint CAML: Lookup keyword in specific column
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='" + columnName + "'/>" +
                                "<Value Type='Lookup'>" + columnValue + "</Value>" +
                                "</Eq></Where></Query>" +
                                "<RowLimit>" + rowLimit + "</RowLimit></View>";

            ListItemCollection collListItem = oList.GetItems(camlQuery);


            /// Create list with Sharepoint fields
            if (collListItem != null)
            {

                var list = new List<Item>();
                var items = oList.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();

                foreach (var col in items)
                {

                    list.Add(new Item()
                    {
                        id = Convert.ToInt32(col["ID"]),
                        documentID = Convert.ToString(col["_dlc_DocId"]),
                        title = Convert.ToString(col["Title"]),
                        url = helper.hostURL(clientContext) + Convert.ToString(col["FileRef"]),
                        filetype = Convert.ToString(col["File_x0020_Type"]),
                        created = Convert.ToString(col["Created"]),
                        modified = Convert.ToString(col["Modified"]),
                    });
                }

                return list;

            }
            else {
                return "[]";
            }
        }

    }
}