using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.Repositories.WareHouse;
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Utility;
using Stimulsoft.Report;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Areas.WareHouse.Controllers
{
    public class WareHouseController : Controller
    {
        #region تعریف متغیر ها
        SaabEntities db = new SaabEntities();
        tbCommodityProjectRepository CommodityProjectTable;
        tbPeymanContractsRepository rep_peyman;
        tbUsersRepository rep_users;
        SaabEntities Context = new SaabEntities();
        #endregion

        #region سازنده ها
        public WareHouseController()
        {
            CommodityProjectTable = new tbCommodityProjectRepository();
            rep_peyman = new tbPeymanContractsRepository();
            rep_users = new tbUsersRepository();
        }
        #endregion


        #region صفحه ها
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

        [AuthorizeAAA]
        public ActionResult Tree()
        {
            StringBuilder builder = new StringBuilder();
            string parent = "";
            parent = "{\"key\":" + 1 + "," + "\"name\":\"" + "انبار 1" + "\"" + "," + "\"title\":" + "\"" + "1" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 2 + "," + "\"name\":\"" + "انبار 2" + "\"" + "," + "\"title\":" + "\"" + "انبار 2" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "1" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 3 + "," + "\"name\":\"" + "انبار 3" + "\"" + "," + "\"title\":" + "\"" + "انبار 3" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "1" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 4 + "," + "\"name\":\"" + "انبار 4" + "\"" + "," + "\"title\":" + "\"" + "انبار 4" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "1" + "},";
            builder.Append(parent);
            if (builder.Length > 0)
            {
                builder.Replace(",", "", (builder.Length - 1), 1);
            }

            return View(builder);

        }

        [AuthorizeAAA]
        public ActionResult TreeGrouping()
        {
            StringBuilder builder = new StringBuilder();
            string parent = "";
            parent = "{\"key\":" + 1 + "," + "\"name\":\"" + "دسته بندی" + "\"" + "," + "\"title\":" + "\"" + "1" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 2 + "," + "\"name\":\"" + "ابزار آلات" + "\"" + "," + "\"title\":" + "\"" + "انبار 2" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "1" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 3 + "," + "\"name\":\"" + "رول پلاک" + "\"" + "," + "\"title\":" + "\"" + "انبار 3" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "2" + "},";
            builder.Append(parent);

            parent = "{\"key\":" + 4 + "," + "\"name\":\"" + "دریافتی" + "\"" + "," + "\"title\":" + "\"" + "انبار 4" + "\"" + "," + "\"pic\":" + "\"" + "paszamine_com_1701_Full_HD.jpg" + "\"" + "," + "\"parent" + "\":" + "1" + "},";
            builder.Append(parent);
            if (builder.Length > 0)
            {
                builder.Replace(",", "", (builder.Length - 1), 1);
            }

            return View("~/Areas/WareHouse/Views/WareHouse/Tree.cshtml", builder);

        }


        public ActionResult Product_definition()
        {
            return PartialView();
        }


        /// <summary>
        /// پارشیال ویو لیست پروژه ها
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult TableProjects()
        {
            var list = CommodityProjectTable.Listt().OrderByDescending(p => p.Id).ToList();

            return PartialView("~/Areas/WareHouse/Views/WareHouse/_TableProjects.cshtml", list);
        }

        [AuthorizeAAA]
        public ActionResult _ManageProjects(string id)
        {
            tbCommodityProject st = CommodityProjectTable.Listt().Where(u => u.ProjectCode == id).First();

            return PartialView(st);
        }

        [AuthorizeAAA]
        public ActionResult _GetPeymanContractsSelectEdit(string id)
        {
            tbCommodityProject st = CommodityProjectTable.Listt().Where(u => u.ProjectCode == id).First();

            ViewBag.OptSelected = st.tbPeymanContracts.pec_Title;
            var res = rep_peyman.Update();
            return PartialView(res);
        }

        public string AcceptProj(int userContractID)
        {
            try
            {
                tbProduct_project_Accept by = new tbProduct_project_Accept();

                by.FK_Project = userContractID;
                by.Accept = true;

                var User = new tbUsers();
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    int userid = 0;
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    using (SaabEntities db = new SaabEntities())
                    {
                        User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    }
                    if (User != null)
                    {

                        userid = User.usr_ID;

                    }
                }
                var e = db.tbCommodityProject.Where(p => p.Id == userContractID).FirstOrDefault();

                var exist3 = db.tbReffrenceAccept.Where(p => p.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == 6 && p.FK_PeymanID == e.FK_PeymanContracts && p.tbReffrenceSaveLevel.tbReffrenceSave.FK_PeymanID == e.FK_PeymanContracts).FirstOrDefault();
                if (exist3 != null)
                {
                    var exist_us = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist3.ID).OrderByDescending(p => p.ID).FirstOrDefault();
                    if (exist_us != null)
                    {
                        var exist_final = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).OrderByDescending(p => p.ID).Select(p => p.FK_User).FirstOrDefault();
                        if (exist_final == User.usr_ID)
                        {
                            by.Final_Accept = true;
                        }
                    }
                }
                else
                {
                    by.Final_Accept = false;
                }


                var c = db.tbCommodityProject.Where(s => s.Id == userContractID).FirstOrDefault();

                if (c != null)
                {
                    by.Year = c.DateOfDefinitionCommodityProject.GetShamsYear();
                    by.Month = c.DateOfDefinitionCommodityProject.GetShamsiMonth();

                    // Check if a record with the same project, month, and year already exists
                    var existingRecord = db.tbProduct_project_Accept
                        .Where(s => s.FK_Project == userContractID && s.Month == by.Month && s.Year == by.Year && s.Final_Accept == true)
                        .FirstOrDefault();

                    if (existingRecord != null)
                    {
                        return "True2";
                    }

                    db.tbProduct_project_Accept.Add(by); // Add the new entity to the context
                    db.SaveChanges(); // Save changes to the database

                    return "True";
                }
                else
                {
                    return "False"; // Handle the case where tbCommodityProject record is not found
                }
            }
            catch (DbUpdateException ex)
            {
                // Log or handle the exception
                // Access the inner exception details for more information
                var innerException = ex.InnerException;
                while (innerException != null)
                {
                    // Log or inspect inner exception details
                    innerException = innerException.InnerException;
                }

                return "Error: " + ex.Message; // Or return an appropriate error message
            }
        }



        public ActionResult Acceptview_Proje()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository rep_accept = new SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository();
                        var result = rep_accept.ListtForproje(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            baste_id.Add(item.FK_PeymanID);
                        }
                        return View(NotSubmitedMoalefe(baste_id));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }

        public ActionResult Acceptview_Sorat()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository rep_accept = new SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository();
                        var result = rep_accept.ListtProjSorat(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            baste_id.Add(item.FK_ReffrenceSaveLevel);
                        }
                        return View(NotSubmitedMoalefe2(baste_id));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }


        public List<tbsaveSoratBasteh> NotSubmitedMoalefe2(List<int?> Bastehlst)
        {

            return db.tbsaveSoratBasteh.Where(p => Bastehlst.Contains((int)p.fk_Basteh) && p.ISsubmit == false).ToList();

        }
        public List<tbsaveSoratBasteh> NotSubmitedMoalefe22(List<int?> Bastehlst, int Month, int Year)
        {

            return db.tbsaveSoratBasteh.Where(p => Bastehlst.Contains((int)p.fk_Basteh)&&p.tbSoratSavefromExcelMoalfe.Any(s=>s.Month==Month&&s.Year==Year) && p.ISsubmit == false).ToList();

        }
        public ActionResult Listview_Sorat()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository rep_accept = new SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository();
                        var result = rep_accept.ListtProjSorat(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            baste_id.Add(item.FK_ReffrenceSaveLevel);
                        }
                        return View(NotSubmitedMoalefe2(baste_id));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }
        public ActionResult TableFordetaiusr(int idd)
        {
            var svdfunc = db.tbsaveSoratBasteh.Find(idd);
            //var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            //if (cookie_user != null)
            //{
            //    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //    using (SaabEntities db = new SaabEntities())
            //    {
            //        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
            //        if (User != null)
            //        {

            //            if (db.tbFucntionExcel.Where(p => p.FK_UserSubmit == User.usr_ID).FirstOrDefault() != null)
            //            {
            //                svdfunc.OnlineUserUploadedFile = true;
            //            }
            //            else
            //            {
            //                svdfunc.OnlineUserUploadedFile = false;
            //            }
            //        }
            //    }
            //}

            //var Basteh = svdfunc.svdfunc_BastehID;
            //var Peyman = svdfunc.svdfunc_pymnID;
            //var result = functionexcelRepo.Update().Where(p => p.FK_Basteh == Basteh && p.FK_Peyman == Peyman && p.Funcexcl_DateTime >= svdfunc.svdfunc_FromDate && p.Funcexcl_DateTime <= svdfunc.svdfunc_ToDate).ToList();

            return PartialView(svdfunc);
        }



        public List<tbCommodityProject> NotSubmitedMoalefe(List<int?> Bastehlst)
        {

            return db.tbCommodityProject.Where(p => Bastehlst.Contains((int)p.FK_PeymanContracts)).ToList();

        }

        public ActionResult tBAcceptList(int Year = 0, int Month = 0, int Peyman_ID = 0)
        {
            List<tbCommodityProject> usr2 = new List<tbCommodityProject>();

            var x2 = db.tbCommodityProject

.Select(s => s.DateOfDefinitionCommodityProject)
.ToList();



            var ex = x2
                .Where(date => /*date?.GetShamsiMonth() >= month  && */date.GetShamsYear() == Year && date.GetShamsiMonth() == Month)
            .ToList();
            foreach (var item in ex)
            {
                var c = db.tbCommodityProject.Where(p => p.DateOfDefinitionCommodityProject == item && p.FK_PeymanContracts == Peyman_ID && p.P_Active == true).FirstOrDefault();
                usr2.Add(c);
            }
            return View(usr2);

        }

        #endregion

        #region رویداد ها



        public string UploadAttachmentFile(int FK_PeymanContracts = 0, int Totalnumber = 0, bool P_Active = true,

   int City_Id = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, string ProjectTitle = "", string ProjectCode = "", DateTime DateOfDefinitionCommodityProject = default,
   IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {


            tbCommodityProject obj = new tbCommodityProject();



            obj.City_Id = City_Id;







            obj.Totalnumber = Totalnumber;





            obj.P_Active = P_Active;




            obj.FK_PeymanContracts = FK_PeymanContracts;


            obj.DateOfDefinitionCommodityProject = DateOfDefinitionCommodityProject;
            obj.ProjectCode = ProjectCode;
            obj.ProjectTitle = ProjectTitle;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/WareHouse/Contents/WareHouseAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;

                        return CommodityProjectTable.Create(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return CommodityProjectTable.Create(obj).ToString();
            }
        }

        public string UploadAttachmentFile_edit(int FK_PeymanContracts = 0, int Totalnumber = 0, bool P_Active = true, string File_SystemName = "", string File_RealName = "", int Id = 0,

   int City_Id = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, string ProjectTitle = "", string ProjectCode = "", DateTime DateOfDefinitionCommodityProject = default,
   IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {


            tbCommodityProject obj = new tbCommodityProject();



            obj.City_Id = City_Id;
            obj.Id = Id;







            obj.Totalnumber = Totalnumber;





            obj.P_Active = P_Active;




            obj.FK_PeymanContracts = FK_PeymanContracts;


            obj.DateOfDefinitionCommodityProject = DateOfDefinitionCommodityProject;
            obj.ProjectCode = ProjectCode;
            obj.ProjectTitle = ProjectTitle;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/WareHouse/Contents/WareHouseAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;

                        return CommodityProjectTable.Update(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {
                obj.File_SystemName = File_SystemName;
                obj.File_RealName = File_RealName;

                return CommodityProjectTable.Update(obj).ToString();
            }
        }









        [AuthorizeAAA]
        public ActionResult SubmitCommodityProject(tbCommodityProject zt, string[] Peymans)
        {
            zt.P_Active = true;
            for (int i = 0; i < Peymans.Length; i++)
            {
                zt.FK_PeymanContracts = Int32.Parse(Peymans[i]);
                CommodityProjectTable.Create(zt);
            }
            CommodityProjectTable.SaveChanges();

            return Content("1");
        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult _ManageProjects(tbCommodityProject st, string ProjectCode1)
        {
            tbCommodityProject st1 = CommodityProjectTable.Listt().Where(u => u.ProjectCode == ProjectCode1).First();
            st.Id = st1.Id;
            CommodityProjectTable.Update(st);
            CommodityProjectTable.SaveChanges();
            return Content("1");
        }

        [AuthorizeAAA]
        public ActionResult _GetCities(string id)
        {
            tbCommodityProject st1 = CommodityProjectTable.Listt().Where(u => u.ProjectCode == id).First();

            var obj = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
            var City = obj.Where(p => p.cityId == st1.City_Id).Select(p => p.cityName).FirstOrDefault();

            return Content(City);
        }

        [AuthorizeAAA]
        public ActionResult DisableProject(int id)
        {
            tbCommodityProject st = CommodityProjectTable.Find(id);
            st.P_Active = false;
            CommodityProjectTable.Update(st);
            CommodityProjectTable.SaveChanges();
            return Content("1");
        }

        [AuthorizeAAA]
        public ActionResult EnableProject(string id)
        {
            tbCommodityProject st = CommodityProjectTable.Find(Int32.Parse(id));
            st.P_Active = true;
            CommodityProjectTable.Update(st);
            CommodityProjectTable.SaveChanges();
            return Content("1");
        }

        #endregion
    }
}