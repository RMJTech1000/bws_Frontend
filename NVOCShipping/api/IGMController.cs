using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using DataManager;
using DataTier;

namespace NVOCShipping.api
{
    public class IGMController : ApiController
    {
        [ActionName("PortManifest")]
        public List<MyIGM> PortManifest(MyIGM Data)
        {

            IGMManager IgmMange = new IGMManager();
            List<MyIGM> st = IgmMange.ManifestPortMaster(Data);
            return st;
        }

        [ActionName("GeoLocationManifest")]
        public List<MyIGM> GeoLocationManifest(MyIGM Data)
        {

            IGMManager IgmMange = new IGMManager();
            List<MyIGM> st = IgmMange.ManifestGeoLocationMaster(Data);
            return st;
        }
    }
}