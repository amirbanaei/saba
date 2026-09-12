using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Documents.Controllers
{
    public class DocumentController : Controller
    {
        // GET: Documents/Document
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

    }
}