using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;

namespace NVOCShipping.Controllers
{
    public class UsersController : Controller
    {
        // GET: Users
        public ActionResult user()
        {
            return View();
        }
        public ActionResult userview()
        {
            return View();
        }

        public ActionResult ControlParameter()
        { 
            return View();
        }
    }
}