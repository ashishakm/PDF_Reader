using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PDF_Demo.View
{
    /// <summary>
    /// Summary description for FileUpload1
    /// </summary>
    public class FileUpload1 : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            context.Response.Write("Excel file created");
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}