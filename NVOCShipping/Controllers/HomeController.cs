using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;


namespace NVOCShipping.Controllers
{

    public class HomeController : Controller
    {

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult test()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        #region anand
        public ActionResult Reg()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult Dashboard()
        {
            ViewBag.Message = "Your Dashboard.";

            return View();
        }
        public ActionResult Location()
        {
            ViewBag.Message = "Your Dashboard.";

            return View();
        }
        public ActionResult Country(string id)
        {

            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Country.";

            return View(BusData);
        }
        public ActionResult CountryView()
        {
            ViewBag.Message = "Your Country.";

            return View();
        }
        public ActionResult Currency(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Currency.";

            return View(BusData);

        }
        public ActionResult CurrencyView()
        {
            ViewBag.Message = "Your Currency.";

            return View();
        }
        public ActionResult Depot(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Depot.";

            return View(BusData);
        }
        public ActionResult DepotView()
        {
            ViewBag.Message = "Your DepotView.";

            return View();
        }
        public ActionResult Customer()
        {
            ViewBag.Message = "Your Customer.";

            return View();
        }

        public ActionResult Terminal(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Terminal.";

            return View(BusData);
        }

        public ActionResult TerminalView()
        {
            ViewBag.Message = "Your TerminalView.";

            return View();
        }

        public ActionResult State()
        {
            ViewBag.Message = "Your State.";

            return View();
        }
        public ActionResult StateView()
        {
            ViewBag.Message = "Your StateView.";

            return View();
        }

        public ActionResult NotesandClausesView()
        {
            ViewBag.Message = "Your StateView.";

            return View();
        }
        public ActionResult NotesandClauses()
        {
            ViewBag.Message = "Your StateView.";

            return View();
        }

        #endregion

        #region Ganesh
        public ActionResult City(string id)
        {

            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your City.";

            return View(BusData);
        }
        public ActionResult CityView()
        {
            ViewBag.Message = "Your CityView.";

            return View();
        }
        public ActionResult Port(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Port.";


            return View(BusData);

        }

        public ActionResult PortView()
        {
            ViewBag.Message = "Your Port.";

            return View();
        }
        public ActionResult MainPort()
        {
            ViewBag.Message = "Your Port.";

            return View();
        }
        public ActionResult MainPortView()
        {
            ViewBag.Message = "Your Port.";

            return View();
        }
        public ActionResult CargoPackage(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Cargo Package.";


            return View(BusData);

        }

        public ActionResult CargoPackageView()
        {
            ViewBag.Message = "Your Port.";

            return View();
        }

        public ActionResult Commodity(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Commodity.";


            return View(BusData);

        }

        public ActionResult CommodityView()
        {
            ViewBag.Message = "Your Commodity.";

            return View();
        }

        public ActionResult ExchangeRate(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your ExchangeRate.";


            return View(BusData);

        }

        public ActionResult ExchangeRateView()
        {
            ViewBag.Message = "Your ExchangeRate.";

            return View();
        }

        public ActionResult GeoLocationView()
        {
            ViewBag.Message = "Your GeoLocation.";

            return View();
        }

        public ActionResult GeoLocation()
        {
            ViewBag.Message = "Your GeoLocation.";

            return View();
        }
        #endregion


    }
}