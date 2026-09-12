using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Educations.Controllers
{
    public class EducationController : Controller
    {
        // GET: Educations/Education
        [AuthorizeAAA] 
        public ActionResult Index()
        {
            return View();   
        } 
    }
}