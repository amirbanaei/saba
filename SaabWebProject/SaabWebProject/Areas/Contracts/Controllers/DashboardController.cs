using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class DashboardController : Controller
    {
        // GET: Contracts/Dashboard
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
    }
}