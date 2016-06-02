# sharepointConnect
C# Class Library to interact with a Sharepoint Library or List

## Prerequisites

This project contains references to:
 - Microsoft.SharePoint.Client.dll
 - Microsoft.SharePoint.Client.Runtime.dll

These files can be found on a Sharepoint (2013) server in the folder C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\ or can be downloaded from from https://www.microsoft.com/en-us/download/details.aspx?id=35585

## ASP.net MVC Example

Add sharepointConnect to your existing project
 
```
using sharepointConnect;

namespace mvcDemo
{

	public Action spConnect = new Action(
		"http://sharepoint-url/sites/sales",            /// Sharepoint URL
		"http://sharepoint-url/sites/sales/documents",  /// Site collection
		"Sales Documents");                             /// library or list name
```

Upload file to a Sharepoint library

```
	[HttpPost]
	public ActionResult Upload(string column_name, string column_value, bool checkIn)
	{
		HttpPostedFileBase file = Request.Files[0];
		var result = spConnect.upload(file, 
						column_name, // Library column name (e.g. sales_id)
						column_value, // Library column value (e.g. 2202)
						checkIn); // Directly check-in the file

        	return Json(result, JsonRequestBehavior.AllowGet);
	}

```

Delete file from a Sharepoint library

```
        [HttpPost]
        public ActionResult Delete(string id)
        {
           /// id = spActionA.delete(id, "Sales Documents"); // id and name of the library

            id = spConnect.delete(id, "Test");

            return Json(new { id = id });
        }
```

Get a list of all documents of a Sharepoint library

```
        public JsonResult getDocuments()
        {
            var result = spConnect.get("IT Documents",		/// Library name 
					"ContentType",  	/// Column name (e.g. Content Type)
					"Sales Reports", 	/// Column value (e.g. Sales Reports)
					10);			/// number of items to return

            return Json(result, JsonRequestBehavior.AllowGet);
        }
```
