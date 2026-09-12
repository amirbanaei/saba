using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.ManagementAccounting.Controllers
{
    public class CashDistributionController : Controller
    {
        // GET: ManagementAccounting/CashDistribution
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult ChooseMoalefe()
        {
            return View("~/Areas/ManagementAccounting/Views/CashDistribution/ChooseMoalefe.cshtml");
        }
    }
}
