using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.BaseInformation;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.BaseInformation.Controllers
{

    public class MoalefeController : Controller
    {
        SaabEntities db;
        tbUnitParameterRepository unitRepo;
        public MoalefeController()
        {
            db = new SaabEntities();
            unitRepo = new tbUnitParameterRepository(db);
        }
        // GET: BaseInformation/Moalefe
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View("~/Areas/BaseInformation/Views/Moalefe/Index.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult CreateVahed(string VahedValue)
        {
            if(VahedValue.Trim()=="")
            {
                TempData["response"] = "NULL";
                return RedirectToAction("Index", "Moalefe");
                
            }
            tbUnitParameter entity = new tbUnitParameter
            {
                unt_Name=VahedValue,
                Status=true
            };
            var Create = unitRepo.Create(entity);
            if(Create=="True")
            {
                TempData["response"] = "True";
                return RedirectToAction("Index", "Moalefe");

            }
            else
            {
                TempData["response"] = "False";
                return RedirectToAction("Index", "Moalefe");

            }
        }
        [AuthorizeAAA]
        public ActionResult Coefficient_pyman()
        {
            return View();
        }

        [AuthorizeAAA]
        public ActionResult _VahedListTable()
        {
            var Model = unitRepo.Update();
            if(Model!=null)
            {
                return PartialView("_VahedListTable", Model);
            }
            else
            {
                return PartialView("_VahdeListTable", new tbUnitParameter());
            }
            
        }
        [AuthorizeAAA]
        public ActionResult DisableUnit(int Id)
        {
            if (Id != 0)
            {
                var res = unitRepo.Disable(Id);
                if (res)
                {
                    TempData["Success"] = true;
                }
                else if (res == false)
                {
                    TempData["Success"] = false;
                }

            }
            return RedirectToAction("Index");
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _EditVahed(tbUnitParameter unit)
        {
            unit.Status = true;
            var res = unitRepo.Update(unit);
            if (res == "True")
            {
                TempData["Success"] = true;
            }
            else if (res == "False")
            {
                TempData["Success"] = false;
            }
            return RedirectToAction("Index");
        }
        [AuthorizeAAA]
        public ActionResult _EditVahed(int Id)
        {
            var res = unitRepo.Find(Id);
            return PartialView("_EditVahed", res);
        }
        [AuthorizeAAA]
        public ActionResult _ListVahedForSelect()
        {
            return PartialView("_ListVahedForSelect", unitRepo.ListActiveUnit());
        }
        public ActionResult view_zarib_caran()
        {
            return PartialView("view_zarib_caran");
        }
        public ActionResult view_zarib_caran_list()
        {
            return PartialView("view_zarib_caran_list", db.tbCaranSettings_zarib.ToList());
        }
        public ActionResult create_sabt1(int citys=0,float codPersonal=0,int value=0)
        {
            try
            {
                tbCaranSettings_zarib tbCaranSettings_zarib = new tbCaranSettings_zarib();
                var find = db.tbCaranSettings_zarib.Where(s => s.FK_city == citys && s.FK_moalfeh == value).FirstOrDefault();
                if (find != null)
                {
                    find.valu = (float)codPersonal;
                    db.SaveChanges();


                }
                else
                {
                    tbCaranSettings_zarib.FK_city = citys;
                    tbCaranSettings_zarib.FK_moalfeh = value;
                    tbCaranSettings_zarib.valu = (float)codPersonal ;
                    db.tbCaranSettings_zarib.Add(tbCaranSettings_zarib);
                    db.SaveChanges();


                }
                return Content("TRue");
            }
            catch (Exception ex)
            {
                return Content(ex.Message);

            }


        }
        public ActionResult create_sabt(int citys = 0, decimal codPersonal = 0, int value = 0)
        {
            try
            {
                //float finalValue = (float)Math.Round(codPersonal, 2);

                var find = db.tbCaranSettings_zarib
                    .FirstOrDefault(s => s.FK_city == citys && s.FK_moalfeh == value);

                if (find != null)
                {
                    find.value_decimal = codPersonal;
                }
                else
                {
                    tbCaranSettings_zarib item = new tbCaranSettings_zarib();
                    item.FK_city = citys;
                    item.FK_moalfeh = value;
                    item.value_decimal = codPersonal;
                    db.tbCaranSettings_zarib.Add(item);
                }

                db.SaveChanges();
                return Content("TRue");
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
        }

        public ActionResult _ListVahedForSelect222()
        {
            return PartialView("_ListVahedForSelect222", unitRepo.ListActiveUnit());
        }
    }
}