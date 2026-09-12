using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.FieldMonitoring.Controllers
{
    public class FieldMonitoringController : Controller
    {
        // GET: FieldMonitoring/FieldMonitoring
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
    }
}