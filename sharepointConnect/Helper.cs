using System;
using System.IO;

namespace sharepointConnect
{
    class Helper
    {
        /// <summary>
        /// This function reads the data of the uploaded file
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>
        /// Uploaded data
        /// </returns>
        public byte[] ReadFile(string filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            int length = Convert.ToInt32(fs.Length);
            byte[] data = new byte[length];
            fs.Read(data, 0, length);
            fs.Close();
            return data;
        }

        /// <summary>
        /// This function gets the Sharepoint hostname
        /// </summary>
        //// <param name="clientContext"></param>
        /// <returns>
        /// Sharepoint hostname
        /// </returns>
        public string hostURL(string clientContext)
        {
            Uri uri = new Uri(clientContext);
            string pathQuery = uri.PathAndQuery;
            string host = uri.ToString().Replace(pathQuery, "");
            return host;
        }

    }
}
