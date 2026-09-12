using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Setting;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Setting.Controllers
{

    public class SaleController : Controller
    {
        //private SaabEntities db;

        //public SaleController(SaabEntities _db)
        //{
        //    db = _db;            
        //}

        SettingRepository AdminRepository;
        SellerSettingRepository SellerSettingRepository;
        tbSetting_Sell_CreateRepository tbSetting_Sell_CreateRepository;
        public SaleController()
        {
            AdminRepository = new SettingRepository();
            SellerSettingRepository = new SellerSettingRepository();
            tbSetting_Sell_CreateRepository = new tbSetting_Sell_CreateRepository();
        }

        // GET: Setting/Sale
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// تعریف خریدار
        /// </summary>
        /// <returns></returns>     
        /// 
        [AuthorizeAAA]
        public ActionResult _Definition_Buyer()
        {
            return PartialView();
        }

        /// <summary>
        /// تعریف فروشنده
        /// </summary>
        /// <returns></returns>    
        /// 
        [AuthorizeAAA]
        public ActionResult _Definition_Seller()
        {

            return PartialView(AdminRepository.Find(1));
        }

        [AuthorizeAAA]
        public ActionResult _Definition_Seller2()
        {
            return View("~/Areas/Setting/Views/Sale/_Definition_Seller2.cshtml");
        }

        public ActionResult _Definition_Seller_List()
        {
            return View("~/Areas/Setting/Views/Sale/_Definition_Seller_List.cshtml");
        }

        public ActionResult _Definition_Seller_Update()
        {
            return View("~/Areas/Setting/Views/Sale/_Definition_Seller_Update.cshtml");
        }

        public string UploadAttachmentFile(string UserName = "", string UserTitle = "", string Password1 = "", string Password2 = "", string NameSoftwar = ""
            , string VersionSoftwar = "", DateTime DateSell = default, DateTime DateSuport = default,
     IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {


            tbSetting_Sell_Create obj = new tbSetting_Sell_Create();

            if (UserName == "")
            {
                obj.UserName = null;
            }
            else
            {
                obj.UserName = UserName;
            }

            if (UserTitle == "")
            {
                obj.UserTitle = null;
            }
            else
            {
                obj.UserTitle = UserTitle;
            }

            obj.Password1 = Password1;
            obj.Password2 = Password2;
            obj.NameSoftwar = NameSoftwar;
            obj.VersionSoftwar = VersionSoftwar;
            obj.DateSell = DateSell;
            obj.DateSuport = DateSuport;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Setting/Images/AdminImages/" + filename));
                        obj.UserImage = filename;

                        return tbSetting_Sell_CreateRepository.Create(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return tbSetting_Sell_CreateRepository.Create(obj).ToString();
            }
        }
        public string EditeUploadAttachmentFile(string UserName = "", string UserTitle = "", string Password1 = "", string Password2 = "", string NameSoftwar = ""
            , string VersionSoftwar = "", DateTime DateSell = default, DateTime DateSuport = default, int ID = 0,
     IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {


            tbSetting_Sell_Create obj = tbSetting_Sell_CreateRepository.Find(ID);

            if (UserName == "")
            {
                obj.UserName = null;
            }
            else
            {
                obj.UserName = UserName;
            }

            if (UserTitle == "")
            {
                obj.UserTitle = null;
            }
            else
            {
                obj.UserTitle = UserTitle;
            }

            obj.Password1 = Password1;
            obj.Password2 = Password2;
            obj.NameSoftwar = NameSoftwar;
            obj.VersionSoftwar = VersionSoftwar;
            obj.DateSell = DateSell;
            obj.DateSuport = DateSuport;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Setting/Images/AdminImages/" + filename));
                        obj.UserImage = filename;

                        return tbSetting_Sell_CreateRepository.Update2(obj).ToString();

                    }
                }
                return "True";
            }
            else
            {

                return tbSetting_Sell_CreateRepository.Update2(obj).ToString();
            }
        }

        /// <summary>
        /// تعریف مدیر سیستم
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _Definition_System_Administrator()
        {
            return PartialView();
        }


        /// <summary>
        /// دامنه فعالیت محصول
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _Scope_Product_Activity()
        {
            return PartialView();
        }

        #region رویداد ها
        /// <summary>
        /// تعریف فروشنده
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _Definition_Seller_Changes(string username)
        {
            tb_Admin Admin = new tb_Admin();
            Admin = AdminRepository.Find(1);

            if (Request.Files.Count > 0)
            {
                HttpFileCollectionBase files = Request.Files;
                string fname;
                HttpPostedFileBase file = files[0];
                fname = file.FileName;
                Admin.Picture = Guid.NewGuid().ToString() + Path.GetExtension(fname);
                file.SaveAs(Server.MapPath("~/Areas/Setting/Images/AdminImages/" + Admin.Picture));
            }
            if (username != null)
            {
                Admin.userName = username;
            }

            AdminRepository.Update(Admin);
            return Content("1");
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _Definition_Seller_ChangesPass1(string pass1)
        {
            tb_Admin Admin = new tb_Admin();
            Admin = AdminRepository.Find(1);

            if (pass1 != null)
            {
                Admin.Pass_1 = pass1;
            }

            AdminRepository.Update(Admin);
            return Content("1");
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _Definition_Seller_ChangesPass2(string pass2)
        {
            tb_Admin Admin = new tb_Admin();
            Admin = AdminRepository.Find(1);

            if (pass2 != null)
            {
                Admin.Pass_2 = pass2;
            }

            AdminRepository.Update(Admin);
            return Content("1");
        }
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult SubmitCard3(string peymanCount, string rightUsersCount, string legalUsersCount, DateTime termination_date, string work_geography, string Maximum_inactivity_time, string MacAddress, string Latitude, string Longitude)
        {
            tb_SellerSetting Seller = new tb_SellerSetting();
            Seller.Count_Peyman = Int32.Parse(peymanCount);
            Seller.Count_RealUsers = Int32.Parse(rightUsersCount);
            Seller.Count_LegalUsers = Int32.Parse(legalUsersCount);
            Seller.SoftwareExpirationDate = termination_date;
            Seller.SoftwareWorkGeography = work_geography;
            Seller.MaxUserInactivityTime = Int32.Parse(Maximum_inactivity_time);
            Seller.MacAddress = MacAddress;
            Seller.SoftwareWorkGIS_X = Latitude;
            Seller.SoftwareWorkGIS_Y = Longitude;

            SellerSettingRepository.Create(Seller);

            return Content("1");
        }
        #endregion
    }
}