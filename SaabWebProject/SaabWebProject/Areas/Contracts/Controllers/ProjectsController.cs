using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Controllers
{
    public class ProjectsController : Controller
    {
        #region متغیر ها
        Models.Repositories.tbProjectsRepository rep_project = new Models.Repositories.tbProjectsRepository();
        #endregion
        #region سازنده ها
        public ProjectsController()
        {

        }
        #endregion
        #region صفحات

        [AuthorizeAAA]
        public ActionResult _PartialViewListProjects()
        {
            return View("~/Areas/Contracts/Views/Projects/_PartialViewListProjects.cshtml", rep_project.Update());
        }
        #endregion

    }
}