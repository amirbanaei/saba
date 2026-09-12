using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.AlarmsAndLog.Controllers
{
    public class PeymanAlarmController : Controller
    {
        // GET: AlarmsAndLog/Peymanc 
        [AuthorizeAAA]
        public ActionResult Enheraf_Az_Bodje()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanAlarm/Enheraf_Az_Bodje.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Etmama_Mablaq_Qarardad()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanAlarm/Etmama_Mablaq_Qarardad.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Navaghes_Sorat_Vaiyat()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanAlarm/Navaghes_Sorat_Vaiyat.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Kasri_Madarek()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanAlarm/Kasri_Madarek.cshtml");
        }
    }
}