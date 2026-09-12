using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Utility;
 
namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class CategoriesController : Controller
    {
        tbCategoriesRepository tbCategoriesRepository;
        private SaabEntities db;
        public CategoriesController() 
        { 
            db = new SaabEntities();
            tbCategoriesRepository = new tbCategoriesRepository(db);
        }
        // GET: Contracts/Categories
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View(); 
        }
        [AuthorizeAAA]
        public ActionResult GetAllCategories()
        {
            return PartialView("_GetAllListCategory", tbCategoriesRepository.Update());
        }
        [AuthorizeAAA]
        [HttpGet]
        public ActionResult _CreateCategory()
        {
            return PartialView("_CreateCategory");
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _CreateCategory(tbCategories category)
        {
            if(category.Category_Name !=null )
            {

            
            var a = db.tbCategories.Where(p => p.Category_Name == category.Category_Name).FirstOrDefault();
            category.Status = true;
            if (a == null)
            {
               
                var res = tbCategoriesRepository.Create(category);
                if (res == "True")
                {
                    TempData["Success"] = true;
                }
                else
                {
                    TempData["Success"] = false;
                }
            }

            else
            {
                TempData["Success"] = false;
            }
            }
            else
            {
                TempData["Success"] = false;
            }
            return View("~/Areas/BaseInformation/Views/Moalefe/Index.cshtml");
        }

        [AuthorizeAAA]

        public ActionResult _EditCategory(int Id)
        {
            var res = tbCategoriesRepository.Find(Id);
            return PartialView("_EditCategory",res);
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _EditCategory(tbCategories category)
        {
            category.Status = true;
            var res = tbCategoriesRepository.Update(category);
            if (res == "True")
            {
                TempData["Success"] = true;
            }
            else if(res == "False")
            {
                TempData["Success"] = false;
            }
            return View("~/Areas/BaseInformation/Views/Moalefe/Index.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult DisableCategory(int Id)
        {
            if (Id != 0)
            {
                var res = tbCategoriesRepository.Disable(Id);
                if (res)
                {
                    TempData["Success"] = true;
                }
                else if (res == false)
                {
                    TempData["Success"] = false;
                }
                
            }
            return View("~/Areas/BaseInformation/Views/Moalefe/Index.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult _GetAllListCategory_Select()
        {
            return PartialView("_GetAllListCategory_Select", tbCategoriesRepository.GetAllActiveCategory());
        }
    }
}