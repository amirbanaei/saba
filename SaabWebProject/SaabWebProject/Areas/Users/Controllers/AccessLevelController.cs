using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Microsoft.Ajax.Utilities;
using PagedList;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.Users;
using SaabWebProject.Utility;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Areas.Users.Controllers
{
    public class AccessLevelController : Controller
    {
        private BusinessSideRepository rep_bussiness;
        private tbUsersRepository rep_users;
        private SaabEntities db;

        public AccessLevelController()
        {
            db = new SaabEntities();
            rep_bussiness = new BusinessSideRepository(db);
            rep_users = new tbUsersRepository(db);
        }

        [AuthorizeAAA]
        public ActionResult _TreeViewAccessLevels_Users(int Id = 0)
        {
            if (Id != 0)
            {
                ViewBag.Id = Id;
            }
            ViewBag.url = Url.Action("_SetAccessLevelForUsers","AccessLevel");
            return PartialView("_TreeViewAccessLevels", rep_bussiness.GetAllActionsForUsers(Id));
        }

        [AuthorizeAAA]
        public ActionResult _TreeViewAccessLevels_Positions(int Id = 0)
        {
            if (Id != 0)
            {
                ViewBag.Id = Id;
            }
            ViewBag.url = Url.Action("_SetAccessLevelForPositions","AccessLevel");
            return PartialView("_TreeViewAccessLevels", rep_bussiness.GetAllActionsForPositions(Id));
        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _SetAccessLevelForUsers(List<string> Actions, int Id)
        {
            return Content(rep_users.SetAccessLevel(Actions, Id).ToString());
        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _SetAccessLevelForPositions(List<string> Actions, int Id)
        {
            return Content(rep_bussiness.SetAccessLevel(Actions, Id).ToString());
        }
    }
}