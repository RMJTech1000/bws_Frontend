using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DataManager;
using System.ComponentModel;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net;

namespace NVOCShipping.Controllers
{
    public class CANEmailController : Controller
    {
        // GET: CANEmail
        public ActionResult Index()
        {
            return View();
        }
    }
}