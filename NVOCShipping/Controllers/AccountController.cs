using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class AccountController : Controller
    {
        // GET: Account
        public ActionResult Receipt()
        {
            return View();
        }
        public ActionResult ReceiptView()
        {
            return View();
        }

        public ActionResult InvoiceSetOff()
        {
            return View();
        }
        public ActionResult InvoiceSetoffView()
        {
            return View();
        }
        public ActionResult ReceiptCancel()
        {
            return View();
        }
        public ActionResult ReceiptCancelView()
        {
            return View();
        }

        public ActionResult PaymentView()
        {
            return View();
        }
        public ActionResult Payment()
        {
            return View();
        }
        public ActionResult JournalVoucherView()
        {
            return View();
        }
        public ActionResult JournalVoucher()
        {
            return View();
        }


        public ActionResult VendorSetoffView()
        {

            return View();
        }

        public ActionResult VendorSetoff()
        {

            return View();
        }
        public ActionResult ReceivableMatching()
        {

            return View();
        }
        public ActionResult ReceivableUnmatching()
        {

            return View();
        }
        public ActionResult ReceivableMatchingView()
        {

            return View();
        }
        public ActionResult ReceivableUnmatchingView()
        {

            return View();
        }
    }
}