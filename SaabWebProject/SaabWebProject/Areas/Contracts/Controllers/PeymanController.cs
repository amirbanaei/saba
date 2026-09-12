using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.DomainModels;
using System.Globalization;
using SaabWebProject.Utility;
using SaabWebProject.Models.Repositories.Contracts;
using Microsoft.Ajax.Utilities;
using SaabWebProject.Models.Utilitis;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using ExcelLibrary.BinaryFileFormat;
using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class PeymanController : Controller
    {
        #region variables
        tbPeymanContractsRepository rep_peyman;
        tbProjectsRepository rep_project = new tbProjectsRepository();
        tbPeymanElhaghieRepository rep_Elhaghie = new tbPeymanElhaghieRepository();
        tbPeymanContractsZaribRepository rep_Zaribha = new tbPeymanContractsZaribRepository();
        tbRial_restrictions_pymnRepository rep_Rial_restrictions = new tbRial_restrictions_pymnRepository();
        PersianCalendar p = new PersianCalendar();
        SaabEntities db = new SaabEntities();

        #endregion

        #region سازنده ها
        public PeymanController()
        {
            rep_peyman = new tbPeymanContractsRepository();
        }


        #endregion

        #region Pages

        [AuthorizeAAA]
        public ActionResult ManagePeymans()
        {
            return View("~/Areas/Contracts/Views/Peyman/ManagePeymans.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult _PartialListPeymans()
        {
            return View("~/Areas/Contracts/Views/Peyman/_PartialListPeymans.cshtml", rep_peyman.Update().OrderByDescending(p => p.pec_ID).ToList());
        }
        [AuthorizeAAA]
        public ActionResult _PartialUpdatePeymans(int id)
        {
            var result = rep_peyman.Find(id);
            result.pec_StartTime_day = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[2]);
            result.pec_StartTime_month = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[1]);
            result.pec_StartTime_year = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[0]);
            result.pec_ENDTime_day = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[2]);
            result.pec_ENDTime_month = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[1]);
            result.pec_ENDTime_year = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[0]);
            foreach (var item in result.tbPeymanElhaghie.ToList())
            {
                if (item.EndTime != null)
                {

                    item.EndTime_day = Convert.ToInt32(item.ShmasiENDTime.Split('/')[2]);
                    item.EndTime_month = Convert.ToInt32(item.ShmasiENDTime.Split('/')[1]);
                    item.EndTime_year = Convert.ToInt32(item.ShmasiENDTime.Split('/')[0]);
                }
                else
                {
                    item.EndTime_day = 0;
                    item.EndTime_month = 0;
                    item.EndTime_year = 0;
                }
            }
            return View("~/Areas/Contracts/Views/Peyman/_PartialUpdatePeymans - Copy.cshtml", result);

        }
        [AuthorizeAAA]
        public ActionResult _PartialCreatePeyman()
        {
            return View("~/Areas/Contracts/Views/Peyman/_PartialCreatePeyman.cshtml", new tbPeymanContracts());

        }
        [AuthorizeAAA]

        public ActionResult _GetAllListPeymans_Select()
        {
            var res = rep_peyman.Update();
            return PartialView(res);
        }
        public ActionResult _GetAllListPeymans_Selecthameh()
        {
            var res = rep_peyman.Update();
            return PartialView(res);
        }


        public ActionResult _GetAllListcitys_Select(int id=0)
        {
            List<tbpeymancities> list = new List<tbpeymancities>();
            if(id != 0)
            {
                list = db.tbpeymancities.Where(p => p.FK_PYMN == id).ToList();


            }
            else
            {
                list = db.tbpeymancities.ToList();

            }
            //var res = rep_peyman.Update();
            return PartialView(list);
        }
        public ActionResult _GetAllListPeymans_Select2()
        {
            var res = rep_peyman.Update();
            return PartialView(res);
        }
        public ActionResult listpysel()
        {
            var t = db.tbPeymanContracts.Where(p=>p.Inactive!=true).ToList();
            return PartialView(t);

        }
      

        [AuthorizeAAA]
        public ActionResult _GetAllListPeymans_SingleSelect()
        {
            var res = rep_peyman.Update();
            return PartialView(res);
        }
        public ActionResult _GetAllListPeymans_SingleSelec2()
        {
            List<tbReffrenceSave> returnList = new List<tbReffrenceSave>();
            List<int> Pymn = new List<int>();
            List<tbPeymanContracts> Pymn2 = new List<tbPeymanContracts>();

            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                var today = DateTime.Now.GetShamsiDayOfMonth();
                var levels = db.tbReffrenceSaveLevel.Where(p => p.StartDaysInMonth <= today && p.Deleted != true && today <= (p.StartDaysInMonth + p.DeadLine)).ToList();
                foreach (var lvl in levels)
                {
                    var baste_duration = lvl.tbReffrenceSave.Duration;
                    if (baste_duration == 1)
                    {
                        var lew = lvl.ID;
                        var l = db.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userid && p.FK_LevelID == lew).FirstOrDefault();
                        if (l != null)
                        {
                            if (lvl.tbReffrenceSave.IsFor == 6)
                            {
                                returnList.Add(lvl.tbReffrenceSave);
                            }
                        }
                    }
                }
            }


            foreach (var x in returnList)
            {
                Pymn.Add((int)x.FK_PeymanID);
            }
            foreach (var item in Pymn)
            {
                var c = db.tbPeymanContracts.Where(p => p.pec_ID == item && p.Inactive != true).FirstOrDefault();
                Pymn2.Add(c);
            }
            return PartialView("_GetAllListPeymans_SingleSelect", Pymn2);
        }


        public ActionResult Deletcityvalue(int ID)
        {
            try
            {
                if (Delete(ID))
                {
                    return Content("True");
                }
                else
                {
                    return Content("False");
                }
            }
            catch (Exception)
            {

                return Content("False");
            }

        }

        public bool Delete(int ID)
        {
            try
            {
                var yy = db.tblink_moalfe_city_valu.Where(p => p.ID == ID).FirstOrDefault();
                db.tblink_moalfe_city_valu.Remove(db.tblink_moalfe_city_valu.Where(p=>p.ID==ID).FirstOrDefault());
                //db.tblink_moalfe_city.Remove(db.tblink_moalfe_city.Where(p => p.ID == yy.FK_tblink_moalfe_city).FirstOrDefault());

                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }

        }

        public ActionResult viewforedit(int id= 0)
        {
            return View(db.tblink_moalfe_city_valu.Where(p => p.ID == id).FirstOrDefault());
        }
        public string edit(float Creditorcredit=0,int Determining_creditline_ID=0)
        {
            var yy = db.tblink_moalfe_city_valu.Where(p => p.ID == Determining_creditline_ID).FirstOrDefault();
            if(yy  != null)
            {
                yy.value =(float)Creditorcredit;
                db.SaveChanges();
            }
            return "True";
        }


        public ActionResult _GetAllListPeymans_SingleSelecforSoratsabt2()
        {
            List<tbReffrenceSave> returnList = new List<tbReffrenceSave>();
            List<int> Pymn = new List<int>();
            List<tbPeymanContracts> Pymn2 = new List<tbPeymanContracts>();

            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                var today = DateTime.Now.GetShamsiDayOfMonth();
                var levels = db.tbReffrenceSaveLevel.Where(p => p.StartDaysInMonth <= today && p.Deleted != true && today <= (p.StartDaysInMonth + p.DeadLine) && p.Deleted != true).ToList();
                foreach (var lvl in levels)
                {
                    var baste_duration = lvl.tbReffrenceSave.Duration;
                    if (baste_duration == 1)
                    {
                        var lew = lvl.ID;
                        var l = db.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userid && p.FK_LevelID == lew).FirstOrDefault();
                        if (l != null)
                        {
                            if (lvl.tbReffrenceSave.IsFor == 5)
                            {
                                returnList.Add(lvl.tbReffrenceSave);
                            }
                        }
                    }
                }
            }



            return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_SingleSelecforSoratsabt2.cshtml", returnList);
        }
        public ActionResult _GetAllListPeymans_SingleSelecforSoratsabt()
        {
            List<tbReffrenceSave> returnList = new List<tbReffrenceSave>();
            List<int> Pymn = new List<int>();
            List<tbPeymanContracts> Pymn2 = new List<tbPeymanContracts>();

            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                var today = DateTime.Now.GetShamsiDayOfMonth();
                var levels = db.tbReffrenceSaveLevel.Where(p => p.StartDaysInMonth <= today && p.Deleted != true && today <= (p.StartDaysInMonth + p.DeadLine)&&p.Deleted!=true).ToList();
                foreach (var lvl in levels)
                {
                    var baste_duration = lvl.tbReffrenceSave.Duration;
                    if (baste_duration == 1)
                    {
                        var lew = lvl.ID;
                        var l = db.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userid && p.FK_LevelID == lew).FirstOrDefault();
                        if (l != null)
                        {
                            if (lvl.tbReffrenceSave.IsFor == 5)
                            {
                                returnList.Add(lvl.tbReffrenceSave);
                            }
                        }
                    }
                }
            }



            return PartialView(returnList);
        }




        public ActionResult GetAllCompanies()
        {


            var returnList = db.tbPeymanContracts.ToList();

            return PartialView(returnList);
        }

        public ActionResult GetAllCompanies345()
        {
            int userid = 0;
            List<tbPeymanContracts> pymn = new List<tbPeymanContracts>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                var finbesines = db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == userid).FirstOrDefault();
                if (finbesines != null && finbesines.FK_up_ID != 9 && finbesines.FK_up_ID != 2175)
                {
                    pymn = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == userid && p.Status == true).Select(s => s.tbPeymanContracts).ToList();
                }
                else
                {
                    pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();
                }
            }
            else
            {
                pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();

            }



            return PartialView("~/Areas/Contracts/Views/Peyman/GetAllCompanies.cshtml", pymn);

        }










        public ActionResult GetAllCompaniespardaght()
        {
            int userid = 0;
            List<tbPeymanContracts> pymn = new List<tbPeymanContracts>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                var finbesines = db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == userid).FirstOrDefault();
                if (finbesines != null && finbesines.FK_up_ID != 9 && finbesines.FK_up_ID != 2175)
                {
                    pymn = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == userid && p.Status == true).Select(s => s.tbPeymanContracts).ToList();
                }
                else
                {
                    pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();
                }
            }
            else
            {
                pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();

            }



            return PartialView("~/Areas/Contracts/Views/Peyman/GetAllCompaniespardaght.cshtml", pymn);

        }




        public ActionResult GetAllCompaniespardaghtsayeer()
        {
            int userid = 0;
            List<tbPeymanContracts> pymn = new List<tbPeymanContracts>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                var finbesines = db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == userid).FirstOrDefault();
                if (finbesines != null && finbesines.FK_up_ID != 9 && finbesines.FK_up_ID != 2175)
                {
                    pymn = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == userid && p.Status == true).Select(s => s.tbPeymanContracts).ToList();
                }
                else
                {
                    pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();
                }
            }
            else
            {
                pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();

            }



            return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select2.cshtml", pymn);

        }










        public ActionResult GetAllCompanies2()
        {


            var returnList = db.tbPeymanContracts.ToList();

            return PartialView(returnList);
        }

        public ActionResult GetAllCompanies233()
        {
            int userid = 0;
            List<tbPeymanContracts> pymn = new List<tbPeymanContracts>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                var finbesines = db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == userid).FirstOrDefault();
               if (finbesines!=null && finbesines.FK_up_ID!=9&& finbesines.FK_up_ID != 2175)
                {
                    pymn=db.Link_User_And_Peyman.Where(p=>p.FK_User_ID== userid &&p.Status==true).Select(s=>s.tbPeymanContracts).ToList();
                }
                else
                {
                    pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();
                }
            }
            else
            {
                pymn = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();

            }



            return PartialView("~/Areas/Contracts/Views/Peyman/GetAllCompanies2.cshtml",pymn);
        }


        public ActionResult _GetAllListPeymans_SingleSelecforFishsabt()
        {
            List<tbReffrenceSave> returnList = new List<tbReffrenceSave>();
            List<int> Pymn = new List<int>();
            List<tbPeymanContracts> Pymn2 = new List<tbPeymanContracts>();

            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                var today = DateTime.Now.GetShamsiDayOfMonth();
                var levels = db.tbReffrenceSaveLevel.Where(p => p.StartDaysInMonth <= today && p.Deleted != true && today <= (p.StartDaysInMonth + p.DeadLine)).ToList();
                foreach (var lvl in levels)
                {
                    var baste_duration = lvl.tbReffrenceSave.Duration;
                    if (baste_duration == 1)
                    {
                        var lew = lvl.ID;
                        var l = db.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userid && p.FK_LevelID == lew).FirstOrDefault();
                        if (l != null)
                        {
                            if (lvl.tbReffrenceSave.IsFor == 7)
                            {
                                returnList.Add(lvl.tbReffrenceSave);
                            }
                        }
                    }
                }
            }



            return PartialView("_GetAllListPeymans_SingleSelecforSoratsabt", returnList);
        }

        public ActionResult _GetAllListPeymans_SingleSelecforFishsabt3()
        {
            List<tbReffrenceSave> returnList = new List<tbReffrenceSave>();
            List<int> Pymn = new List<int>();
            List<tbPeymanContracts> Pymn2 = new List<tbPeymanContracts>();

            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                userid = User.usr_ID;
                List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                var today = DateTime.Now.GetShamsiDayOfMonth();
                var levels = db.tbReffrenceSaveLevel.Where(p => p.StartDaysInMonth <= today && p.Deleted != true && today <= (p.StartDaysInMonth + p.DeadLine)).ToList();
                foreach (var lvl in levels)
                {
                    var baste_duration = lvl.tbReffrenceSave.Duration;
                    if (baste_duration == 1)
                    {
                        var lew = lvl.ID;
                        var l = db.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userid && p.FK_LevelID == lew).FirstOrDefault();
                        if (l != null)
                        {
                            if (lvl.tbReffrenceSave.IsFor == 7)
                            {
                                returnList.Add(lvl.tbReffrenceSave);
                            }
                        }
                    }
                }
            }



            return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_SingleSelecforFishsabt3.cshtml", returnList);
        }




        public ActionResult _GetAllListPeymansForThisUser_SingleSelect()
        {
            int? personal = 0;
            var Model = new List<tbPeymanContracts>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
            personal = user.usr_ID;
            var tarafdovom = db.tbCompanies.Where(p => p.FK_ManagerID == personal).FirstOrDefault();
            if (tarafdovom != null)
            {

                Model = db.tbPeymanContracts.Where(p => p.FK_UserTarafDovvom == tarafdovom.ID && p.Inactive != true).ToList();
            }
            else
            {
                Model = db.tbPeymanContracts.Where(p=>p.Inactive!=true).ToList();
            }
            return PartialView("_GetAllListPeymans_SingleSelect", Model);
        }
        [AuthorizeAAA]
        public ActionResult _SetPeymansPrice()
        {
            return PartialView("~/Areas/Contracts/Views/Peyman/_SetPeymansPrice.cshtml", new tbPeymanContractPrice());
        }
        /// <summary>
        /// صفحه نمایش جزئیات پیمان
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialDetailPeyman(int id)
        {
            var result = rep_peyman.Find(id);
            result.pec_StartTime_day = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[2]);
            result.pec_StartTime_month = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[1]);
            result.pec_StartTime_year = Convert.ToInt32(result.pec_ShmasiStartTime.Split('/')[0]);
            result.pec_ENDTime_day = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[2]);
            result.pec_ENDTime_month = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[1]);
            result.pec_ENDTime_year = Convert.ToInt32(result.pec_ShmasiENDTime.Split('/')[0]);
            foreach (var item in result.tbPeymanElhaghie.ToList())
            {
                if (item.EndTime != null)
                {
                    item.EndTime_day = Convert.ToInt32(item.ShmasiENDTime.Split('/')[2]);
                    item.EndTime_month = Convert.ToInt32(item.ShmasiENDTime.Split('/')[1]);
                    item.EndTime_year = Convert.ToInt32(item.ShmasiENDTime.Split('/')[0]);
                }
                else
                {
                    item.EndTime_day = 0;
                    item.EndTime_month = 0;
                    item.EndTime_year = 0;
                }

            }
            return PartialView("~/Areas/Contracts/Views/Peyman/_PartialDetailPeyman.cshtml", result);
        }
        /// <summary>
        /// تعیین حدود انتظار
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialHadeEntezar(int id)
        {
            var result = rep_peyman.Find(id);

            return PartialView("~/Areas/Contracts/Views/Peyman/_PartialHadeEntezar.cshtml", result);
        }
        [AuthorizeAAA]
        public ActionResult _PartialPey4in(int id)
        {
            var result = rep_peyman.Find(id);
            //var find = db.tbRial_restrictions_pymn.Where(p => p.Fk_pymn == id).FirstOrDefault();
            

            return PartialView("~/Areas/Contracts/Views/Peyman/_PartialPey4in.cshtml", result);
        }

        public ActionResult soratvaziatmafad(int id)
        {
            var result = rep_peyman.Find(id);

            return PartialView("~/Areas/Contracts/Views/Peyman/soratvaziatmafad.cshtml", result);
        }

        public ActionResult barrsiview(int id=0,int numbersorat=0,float hosin=0,float beme=0,int mostrak=0,float jari=0)
        {
            var sor = db.tbSoratvaziat.Where(p => p.number_sorat <= numbersorat && p.Fk_pymn == id).ToList();
            
            long finalsorat = 0;
            float sarm = 100 - jari;
            foreach (var item in sor)
            {
                var numberen = sor.Where(p => p.Finalvaluesorat != null).FirstOrDefault();
                if(numberen != null)
                {
                    finalsorat += (long)numberen.Finalvaluesorat;
                }
            }
            tbpeymaninformatio infor=new tbpeymaninformatio();
            var fin = db.tbpeymaninformatio.Where(p => p.Fk_pymn == id && p.numbersorat == numbersorat).FirstOrDefault();
            if (fin != null)
            {
                infor = fin;
                
            }
            else
            {
                infor.jari = finalsorat*jari;
                infor.hosin = finalsorat*hosin;
                infor.sarmayi = sarm * finalsorat;
                infor.jari=finalsorat* beme;
                infor.Fk_pymn = id;
                infor.numbersorat= numbersorat;
                infor.moshtrakin = mostrak;
                infor.bemeh=beme*finalsorat;
            }


            return PartialView("~/Areas/Contracts/Views/Peyman/barrsiview.cshtml", infor);
        }


        public string UploadRestrictions_ABCD(float restrictions_A = 0, float restrictions_B = 0, float restrictions_C = 0, float restrictions_D = 0,int Fk_pymn = 0,float mahdodatmogym_B=0)
        {
            var find = db.tbRial_restrictions_pymn.Where(p => p.Fk_pymn == Fk_pymn).FirstOrDefault();

            if(find != null)
            {
                find.restrictions_A = restrictions_A;
                find.restrictions_B = restrictions_B;
                find.restrictions_C = restrictions_C;
                find.restrictions_D = restrictions_D;
                find.mahdodatmogym_B = mahdodatmogym_B;

                db.SaveChanges();
                return "True";
            }
            else
            {
                tbRial_restrictions_pymn obj = new tbRial_restrictions_pymn();

                obj.restrictions_A = restrictions_A;
                obj.restrictions_B = restrictions_B;
                obj.restrictions_C = restrictions_C;
                obj.restrictions_D = restrictions_D;
                obj.mahdodatmogym_B = mahdodatmogym_B;

                obj.Fk_pymn = Fk_pymn;
                return rep_Rial_restrictions.Create(obj);

            }
        

        }




        [AuthorizeAAA]
        public ActionResult _PartialPey2in(int id)
        {
            var tt = db.tbPeymanContracts.Where(p => p.pec_ID == id && p.Inactive != true).FirstOrDefault();

            // داده‌های tt متناسب با مدل ویو هستند
            return PartialView("~/Areas/Contracts/Views/Peyman/_PartialPey2in.cshtml", tt);
        }



        [AuthorizeAAA]
        public ActionResult _PartialzaribHa(int id)
        {
            using (var db = new SaabEntities())
            {
                var peyman = db.tbPeymanContracts.Find(id);
                var CreditIndicators = db.tbCreditIndicators.ToList();
                var ZaribHa = db.tbPeymanContractsZarib.Where(p => p.FK_Peyman == id).ToList();
                foreach (var item in CreditIndicators)
                {
                    if (ZaribHa.Where(p => p.FK_CreditIndicators == item.ID).FirstOrDefault() == null)
                    {
                        tbPeymanContractsZarib obj = new tbPeymanContractsZarib(); // نماینده مقیم
                        obj.FK_CreditIndicators = item.ID;
                        obj.FK_Peyman = id;
                        obj.tbCreditIndicators = item;
                        obj.tbPeymanContracts = peyman;
                        obj.Type = 0;
                        obj.Value = null;
                        ZaribHa.Add(obj);

                        tbPeymanContractsZarib obj1 = new tbPeymanContractsZarib(); // کادر اداری
                        obj1.FK_CreditIndicators = item.ID;
                        obj1.FK_Peyman = id;
                        obj1.tbCreditIndicators = item;
                        obj1.tbPeymanContracts = peyman;
                        obj1.Type = 1;
                        obj1.Value = null;
                        ZaribHa.Add(obj1);
                    }
                }
                return PartialView("~/Areas/Contracts/Views/Peyman/_PartialzaribHa.cshtml", ZaribHa);
            }
        }

        [AuthorizeAAA]
        public ActionResult _GetAllListPeymans_FilterSelect(List<tbPeymanContracts> Peymans)
        {
            return PartialView(Peymans);
        }
        #endregion

        #region Functions
        [AuthorizeAAA]
        public string Peyman_Create(tbPeymanContracts Filters)
        {
            Filters.pec_EndTime = p.ToDateTime(Filters.pec_ENDTime_year, Filters.pec_ENDTime_month, Filters.pec_ENDTime_day, 0, 0, 0, 0).Date;
            Filters.pec_StartTime = p.ToDateTime(Filters.pec_StartTime_year, Filters.pec_StartTime_month, Filters.pec_StartTime_day, 0, 0, 0, 0).Date;
            foreach (var item in Filters.tbPeymanContractPrice)
            {
                item.DateOccured = DateTime.Now;
                item.IsElhaghie = false;
            }
            return rep_peyman.Create(Filters);
        }
        [AuthorizeAAA]
        public string Peyman_Update(IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,
            int pec_ID = 0, string pec_Title = "", bool pec_TypeOfContract = false, string pec_Price = "", int FK_KarfarmaID = 0,
            DateTime pec_StartTime = default,
            DateTime pec_EndTime = default, string pec_ContractNumber = "", int pec_ProjectCode = 0, bool pec_IsElhaghie = false, string pec_BriefTitle = "")
        {
            if (pec_ID <= 0 || pec_Title.Trim() == null || pec_TypeOfContract == null || Convert.ToDouble(pec_Price) < 0 || FK_KarfarmaID < 0
                || pec_StartTime == null || pec_EndTime == null || pec_ContractNumber.Trim() == null || pec_ProjectCode < 0)
            {
                return "false_INotValid";
            }

            tbPeymanContracts obj_peyman = new tbPeymanContracts();
            obj_peyman.pec_ID = pec_ID;
            obj_peyman.pec_Title = pec_Title;
            obj_peyman.pec_TypeOfContract = pec_TypeOfContract;
            obj_peyman.pec_Price = pec_Price;
            obj_peyman.FK_KarfarmaID = FK_KarfarmaID;
            obj_peyman.pec_StartTime = pec_StartTime;
            obj_peyman.pec_EndTime = pec_EndTime;
            obj_peyman.pec_BriefTitle = pec_BriefTitle;
            obj_peyman.pec_ContractNumber = pec_ContractNumber;
            obj_peyman.pec_ProjectCode = pec_ProjectCode;
            obj_peyman.pec_IsElhaghie = pec_IsElhaghie;

            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + "." + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/Peymans/" + filename));
                        obj_peyman.pec_FileSystemName = filename;
                        obj_peyman.pec_FileName = file.FileName;
                    }
                }
            }
            if (obj_peyman.pec_IsElhaghie == true)
            {
                obj_peyman.FK_PecFatherElhaghie = obj_peyman.pec_ID;
                obj_peyman.pec_ID = 0;
                return rep_peyman.Create(obj_peyman);

            }
            else
            {
                obj_peyman.FK_PecFatherElhaghie = null;
                return rep_peyman.Update(obj_peyman);
            }
        }

        [AuthorizeAAA]
        public string Peyman_UploadFile2(int PeymanID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            if (PeymanID == 0)
            {
                return "False";
            }
            string message = "";
            if (files != null)
            {
                foreach (var file in files)
                {
                    int maxFileSize = 5 * 1024 * 1024; // 5MB
                    string extension = Path.GetExtension(file.FileName).ToLower();
                    string[] allowedExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".pdf", ".zip" };

                    if (!allowedExtensions.Contains(extension))
                    {
                        message = "فرمت فایل " + file.FileName + " معتبر نیست. فقط فرمت‌های jpg, jpeg, png, gif, bmp, pdf, zip مجاز هستند.";
                        return message;
                    }

                    if (file.ContentLength > maxFileSize)
                    {
                        message = "حجم فایل " + file.FileName + " بیشتر از حد مجاز (5MB) است.";
                        return message; // کل عملیات متوقف
                    }
                    var find = db.madarek_sabtdoreh_4
             .FirstOrDefault(p => p.ID == PeymanID);
                    if (file.ContentLength > 0&& find!=null)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Suggestion/Contents/" + filename));
                        var contract = rep_peyman.Find(PeymanID);
                        contract.pec_FileSystemName = filename;
                        contract.pec_FileName = file.FileName;
                        find.timesabt = DateTime.Now.Date;
                        find.System_FIle = filename;
                        find.Realname_namefile = file.FileName;
                        db.SaveChanges();
                    }
                    else
                    {
                        return "True";
                    }
                }
                return "True";
            }
            else
            {
                return "True";
            }
        }

        public string Peyman_UploadFile(int PeymanID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            if (PeymanID == 0)
            {
                return "False";
            }
            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/Peymans/" + filename));
                        var contract = rep_peyman.Find(PeymanID);

                        contract.pec_FileSystemName = filename;
                        contract.pec_FileName = file.FileName;
                        return rep_peyman.Update(contract);
                    }
                    else
                    {
                        return "True";
                    }
                }
                return "True";
            }
            else
            {
                return "True";
            }
        }
        public string UploadFilehogogi(int ID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {

            if (ID == 0)
            {
                return "False";
            }
            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/Peymans/" + filename));
                        var peyman = Find(ID);
                        peyman.FileSystemName = filename;
                        peyman.FileName = file.FileName;
                        db.SaveChanges();

                        return "True";
                    }
                    else
                    {
                        return "True";
                    }
                }
                return "True";
            }
            else
            {
                return "True";
            }
        }

      
        public tbCompanies Find(int ID)
        {
            try
            {
                return db.tbCompanies.Find(ID);
            }
            catch
            {
                return null;
            }
        }
        [AuthorizeAAA]
        public string UploadFile(int ID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {

            if (ID == 0)
            {
                return "False";
            }
            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/Peymans/" + filename));
                        var peyman = rep_peyman.Find(ID);
                        peyman.pec_FileSystemName = filename;
                        peyman.pec_FileName = file.FileName;
                        return rep_peyman.Update(peyman).ToString();
                    }
                    else
                    {
                        return "True";
                    }
                }
                return "True";
            }
            else
            {
                return "True";
            }
        }

        [AuthorizeAAA]
        private List<tbPeymanContracts> List()
        {
            return rep_peyman.Update();
        }


        [AuthorizeAAA]
        private List<tbPeymanContracts> Details(int peymanID)
        {
            return rep_peyman.Listt(peymanID);
        }


        [AuthorizeAAA]
        public string SaveElhaghie22(tbPeymanElhaghie Filters)
        {
            if (Filters.EndTime_day == 0)
            {
                Filters.EndTime = null;
            }
            else
            {
                Filters.EndTime = p.ToDateTime(Filters.EndTime_year, Filters.EndTime_month, Filters.EndTime_day, 0, 0, 0, 0).Date;


            }
            return rep_Elhaghie.Create(Filters);
        }
        [AuthorizeAAA]
    public string SaveElhaghie(tbPeymanContracts Filters)
    {
        
        return rep_peyman.Update(Filters);
    }

    [AuthorizeAAA]
    public bool Peyman_Delete(int ID)
    {
        return rep_peyman.Disable(ID);//deleted
    }

    public bool Peyman_Delete2(int ID)
        {
            var delet = db.tbPeymanContracts.Where(p => p.pec_ID == ID && p.Inactive != true).FirstOrDefault();
            if (delet != null)
            {
                delet.Inactive = true;
                db.SaveChanges();            }
            return true; 
        }
        public ActionResult AcceptPy(int FK_Pymn = 0, int Month = 0, int Year = 0)
    {

        var exist = db.tbReffrenceSave.Where(p => p.FK_PeymanID == FK_Pymn && p.IsFor == 1).ToList();
        List<tbReffrenceSaveLevel> List = new List<tbReffrenceSaveLevel>();
        List<tbMoalefeDastmozdiValueFromExcel> ListForfin = new List<tbMoalefeDastmozdiValueFromExcel>();


        foreach (var item in exist)
        {
            var sh = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == item.ID && p.Deleted != true && p.Data_time_create.GetShamsiMonth() == Month && p.Data_time_create.GetShamsYear() == Year).FirstOrDefault();
            List.Add(sh);
        }
        foreach (var item in List)
        {
            var sh = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == item.ID && p.svdfunc_IsSubmmit == true && p.Final_Accept == true).FirstOrDefault();
            var ex = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == sh.svdfunc_ID && p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).FirstOrDefault();
            ListForfin.Add(ex);
        }
        if (ListForfin.Count == List.Count)
        {
            var matchedRows = db.tbMoalefeValuePishkhan
                          .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-تایید عملکرد پیمان" && p.mlfval_FKPeyman == FK_Pymn && p.mlfval_Month == Month && p.mlfval_Year == Year)
                          .FirstOrDefault();

            matchedRows.mlfval_Value = "1";

            db.SaveChanges();

            return Content("true");

        }
        else
        {
            return Content("jdjds");
        }









    }

    public bool SaveZaribHa(List<tbPeymanContractsZarib> list_zaribha)
    {
        using (var db = new SaabEntities())
        {
            try
            {
                foreach (var item in list_zaribha)
                {
                    var temp = db.tbPeymanContractsZarib.Where(p => p.FK_Peyman == item.FK_Peyman && p.FK_CreditIndicators == item.FK_CreditIndicators && p.Type == item.Type).FirstOrDefault();
                    if (temp == null)
                    {
                        rep_Zaribha.Create(item);
                    }
                    else
                    {
                        item.ID = temp.ID;
                        rep_Zaribha.Update(item);

                    }
                }
                return true;

            }
            catch (Exception)
            {

                return false;

            }

        }

    }

        public bool downloadPeymanFile(int ID = 0)
        {
            return true;
        }

    #endregion


    }
}