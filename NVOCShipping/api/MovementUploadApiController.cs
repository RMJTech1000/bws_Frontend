using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Configuration;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using DataManager;
using DataTier;

namespace NVOCShipping.api
{
    public class MovementUploadApiController : ApiController
    {

        [ActionName("MovementExcelUpload")]
        public List<MyExcelContainers> MovementExcelUpload(MyExcelContainers Data)
        {
            MovementUploadManager cm = new MovementUploadManager();
            List<MyExcelContainers> st = cm.MovementExcelUpload(Data);
            return st;
        }

        [ActionName("MovementUploadView")]
        public List<MyExcelContainers> MovementUploadView(MyExcelContainers Data)
        {
            MovementUploadManager cm = new MovementUploadManager();
            List<MyExcelContainers> st = cm.MovementUploadView(Data);
            return st;
        }

        [ActionName("MovementUploadLogDtls")]
        public List<MyExcelContainerslog> MovementUploadLogDtls(MyExcelContainerslog Data)
        {
            MovementUploadManager cm = new MovementUploadManager();
            List<MyExcelContainerslog> st = cm.MovementUploadLogDtls(Data);
            return st;
        }
    }
}
