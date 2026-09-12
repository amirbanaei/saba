using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Committees
{
    public class AdministrativeHealthController : Controller 
    {
        // GET: Committees/AdministrativeHealth
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult _Mosavabat_Erja_Shode()
        {
            return View("~/Areas/Committees/Views/AdministrativeHealth/_Mosavabat_Erja_Shode.cshtml");

        }
    }
}