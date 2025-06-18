using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace NVOCShipping.Models
{
    public class Products
    {
        [DisplayName("Product Name")]
        public string ProductName { get; set; }

        [DisplayName("Product Desc")]
        public string ProductDesc { get; set; }

        [DisplayName("Product Image")]
        public string ProductImg { get; set; }

        public HttpPostedFileBase UploadFile { get; set; }
    }
}