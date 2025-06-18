using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataTier;

namespace NVOCShipping.Models
{
    public class MasterModel
    {
       public Masters MasterDetails { get; set; }
    }
    public class Masters
    {
        public int ID { get; set; }
        
    }


    public class MasterDepo
    {
        public string ID { get; set; }
    }

   
}