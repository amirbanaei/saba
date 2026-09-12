using SaabWebProject.Models.StaticClasses;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.AlarmsAndLog.Controllers
{
    public class PeymanKarController : Controller
    {
        // GET: AlarmsAndLog/PeymanKar
        [AuthorizeAAA]
        public ActionResult EnherafAzShakhes()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/EnherafAzShakhes.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult zaman_gharardad_khateme_yafte()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/zaman_gharardad_khateme_yafte.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult sorat_Vaziyat_Sabt_Nashode()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/sorat_Vaziyat_Sabt_Nashode.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult sorat_Vaziyat_Taed_Nashode()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/sorat_Vaziyat_Taed_Nashode.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Navaghes_lebas_va_Kafsh_Kar()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/Navaghes_lebas_va_Kafsh_Kar.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Navaghes_Vasile_Naghliye()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/Navaghes_Vasile_Naghliye.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Karkard_Kamtar_Az_Estandard()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/Karkard_Kamtar_Az_Estandard.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Project_Kalaye_Baz()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/Project_Kalaye_Baz.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult randeman()
        {
            return View("~/Areas/AlarmsAndLog/Views/PeymanKar/randeman.cshtml");

        }
    }
}