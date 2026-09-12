using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Windows.Navigation;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class FunctionsController : Controller
    {
        #region  صفحه اصلی
        tbMoalefeDastmozdiValueFromExcelRepository MoalefeExcelRepo;
        RegistrationAndConfirmationProceduresController registerController;
        tbFunctionExcelRepository functionexcelRepo;
        tbSavedFunctionsRepositories savedfunctionRepo;

        SaabEntities db;

        // GET: Contracts/Functions

        public FunctionsController()
        {
            db = new SaabEntities();
            MoalefeExcelRepo = new tbMoalefeDastmozdiValueFromExcelRepository(db);
            registerController = new RegistrationAndConfirmationProceduresController();
            functionexcelRepo = new tbFunctionExcelRepository(db);
            savedfunctionRepo = new tbSavedFunctionsRepositories(db);
        }
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult _ShowOnePersonModal()
        {
            return View();
        }



        #endregion  


        #region تب کارکرد تجهیزات

        /// <summary>
        /// تب کارکرد تجهیزات
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _RegisterOprationEquipment()
        {
            return PartialView();
        }

        /*------------------------------------------------[1402/08/20]-|KH|-*/
        /// ثبت تکی
        public ActionResult _SingleImportEquipment()
        {
            return View();
        }
        /*-------------------------------------------------------------|KH|-*/

        #endregion

        #region تب کارکرد پروژه های کالایی

        /// <summary>
        /// تب کارکرد پروژه های کالایی
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _RegisterOprationProject()
        {
            return PartialView();
        }
        //public ActionResult ShowMoalefeTable()
        //{

        //    return null;


        //}

        [AuthorizeAAA]
        public ActionResult ShowMoalefeTable(string usersId, int Basteh, DateTime FromDate, DateTime ToDate)
        {
            List<int> Ids = usersId.Split(',').Select(int.Parse).ToList();
            var listMoalefe = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);
            var ListIDtbMoalefeexcel = db.tbSavedFunctions.Where(p => p.svdfunc_FromDate >= FromDate && p.svdfunc_FromDate <= ToDate && p.svdfunc_ToDate >= FromDate && p.svdfunc_ToDate <= ToDate).Select(p => p.svdfunc_ID).FirstOrDefault();
            var Model = MoalefeExcelRepo.Update().Where(p => Ids.Contains((int)p.MoalfeVal_FKUser) && listMoalefe.Contains((int)p.MoalfeVal_FKMoalafeDastmozdi) && p.FK_SavedFunctionsID == ListIDtbMoalefeexcel).ToList();
            // .Where(p=>p.MoalfeVal_FromDate>=FromDate && p.MoalfeVal_FromDate<=ToDate && p.MoalfeVal_ToDate>=FromDate && p.MoalfeVal_ToDate<=FromDate).ToList();

            return View("~/Areas/Contracts/Views/Functions/_ShowMoalefeTable.cshtml", Model);
            //return View();


        }

        [AuthorizeAAA]
        public ActionResult _ShowMoalefeModal()
        {
            return View("~/Areas/Contracts/Views/Functions/_ShowMoalefe.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult ShowUsersFunctions(int year_pY=0,int month_pY=0,int svdfunc_BastehID=0, int fkbast=0)
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
                        var result = rep_accept.Listt(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            if(item.FK_ReffrenceSaveLevel == fkbast)
                            {
                                baste_id.Add(item.FK_ReffrenceSaveLevel);

                            }
                            //if()
                        }
                        return View(NotSubmitedMoalefe222(baste_id,month_pY,year_pY));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }






        public async Task<string> deletebast(int ID = 0)
        {
            var find = await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == ID).ToListAsync();
            if (find.Count != 0)
            {
                db.tbMoalefeDastmozdiValueFromExcel.RemoveRange(find);
                await db.SaveChangesAsync();
            }

            var findfunc =await db.tbFucntionExcel.Where(p => p.FK_SavedFunction == ID).ToListAsync();

            if (findfunc.Count != 0)
            {
                db.tbFucntionExcel.RemoveRange(findfunc);
                await db.SaveChangesAsync();
            }

            var findrelatin = await db.tbSavedFunctions.Where(p => p.svdfunc_ID == ID).FirstOrDefaultAsync();
            if (findrelatin != null)
            {
                db.tbSavedFunctions.Remove(findrelatin);
                await db.SaveChangesAsync();
            }
            return "True";

        }


        public async Task<string> deletebasttool(int ID = 0)
        {
            var find = await db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == ID).ToListAsync();
            if (find.Count != 0)
            {
                db.tbEquipmentMoalefeValue.RemoveRange(find);
                await db.SaveChangesAsync();
            }

            //var findfunc = await db.tbFucntionExcel.Where(p => p.FK_SavedFunction == ID).ToListAsync();

            //if (findfunc.Count != 0)
            //{
            //    db.tbFucntionExcel.RemoveRange(findfunc);
            //    await db.SaveChangesAsync();
            //}

            var findrelatin = await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == ID).FirstOrDefaultAsync();
            if (findrelatin != null)
            {
                db.tbEquipmentMoalefeValueReffrenceSave.Remove(findrelatin);
                await db.SaveChangesAsync();
            }
            return "True";
        }

        public async Task<string> deletebastkar(int ID = 0)
        {
            var find = await db.tbMoalefeValueFish.Where(p => p.FK_EXCel == ID).ToListAsync();
            if (find.Count != 0)
            {
                db.tbMoalefeValueFish.RemoveRange(find);
                await db.SaveChangesAsync();
            }

            //var findfunc = await db.tbFucntionExcel.Where(p => p.FK_SavedFunction == ID).ToListAsync();

            //if (findfunc.Count != 0)
            //{
            //    db.tbFucntionExcel.RemoveRange(findfunc);
            //    await db.SaveChangesAsync();
            //}

            var findrelatin = await db.tbmoalfefishexcel.Where(p => p.ID == ID).FirstOrDefaultAsync();
            if (findrelatin != null)
            {
                db.tbmoalfefishexcel.Remove(findrelatin);
                await db.SaveChangesAsync();
            }
            return "True";
        }





        [AuthorizeAAA]
        public List<tbSavedFunctions> NotSubmitedMoalefe(List<int?> Bastehlst)
        {
            List<tbSavedFunctions> list = new List<tbSavedFunctions>();

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);

                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                var savfin=db.tbSavedFunctions.ToList();
                var gg = db.tbReffrenceAccept.ToList();
                var gg2 = db.tbReffrenceAcceptLevel.ToList();
                var usrfin = db.tbReffrenceAcceptLevelUsers.ToList();

                foreach (var item in Bastehlst)
                {
                    var fin = savfin.Where(p => p.svdfunc_BastehID == item).ToList();
                    var fkb = gg.Where(p => p.FK_ReffrenceSaveLevel == item).FirstOrDefault();

                    if (fkb != null)
                    { var ty = gg2.Where(p => p.FK_ReffrenceAccept == fkb.ID).FirstOrDefault();
                        if (ty != null)
                        {
                            var fi = usrfin.Where(p => p.FK_ReffrenceAcceptLevel == ty.ID).OrderByDescending(s=>s.ID).FirstOrDefault();
                            foreach(var tyuu in fin)
                            {
                                if (fi != null && fi.FK_User == User.usr_ID)
                                {
                                    if (fin != null)
                                    {
                                        if (tyuu.svdfunc_ID == 2242)
                                        {

                                        }
                                        if (tyuu.Final_Accept == true)
                                        {
                                            list.Add(tyuu);
                                        }
                                        else if (tyuu.svdfunc_IsSubmmit == false)
                                        {
                                            list.Add(tyuu);
                                        }
                                    }
                                }
                                else
                                {
                                    if (tyuu.svdfunc_IsSubmmit == false)
                                    {
                                        list.Add(tyuu);

                                    }
                                }
                            }
                         
                        }

                    }
                    
                } }
            return list;
        }




        public List<tbSavedFunctions> NotSubmitedMoalefe222(List<int?> Bastehlst,int month_pY, int year_pY)
        {
            List<tbSavedFunctions> list = new List<tbSavedFunctions>();

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);

                var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();

                foreach (var item in Bastehlst)
                {
                    var fin = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == item && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month_pY && s.MoalfeVal_Year == year_pY)).ToList();
                    var fkb = db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == item).FirstOrDefault();

                    if (fkb != null)
                    {
                        var ty = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == fkb.ID).FirstOrDefault();
                        if (ty != null)
                        {
                            var fi = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == ty.ID).OrderByDescending(s => s.ID).FirstOrDefault();
                            if (fi != null && fi.FK_User == User.usr_ID)
                            {
                                foreach( var t in fin)
                                {
                                    if (t != null)
                                    {
                                        if (t.Final_Accept == true)
                                        {
                                            list.Add(t);
                                        }
                                        else if (t.svdfunc_IsSubmmit == false)
                                        {
                                            list.Add(t);
                                        }
                                    }

                                }

                            }
                          
                        }

                    }

                }
            }
            return list;
        }






        public ActionResult ShowUsersFuctionforPy()
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
                        var result = rep_accept.Listt(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            if (item.FK_ReffrenceSaveLevel == 92)
                            {

                            }
                            baste_id.Add(item.FK_ReffrenceSaveLevel);
                        }
                        return View(NotSubmitedMoalefe(baste_id));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }


        #endregion


        #region تب تایید کارکرد کاربران
      //  public ActionResult Pyman_forAccept()
      //  {
      //      List<tbReffrenceSaveLevel> save3 = new List<tbReffrenceSaveLevel>();
        
      //          var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
      //          if (cookie_user != null)
      //          {
      //              var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
      //              var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();

      //              if (User != null)
      //              {
      //                  var exist_usr = db.tbReffrenceAcceptLevelUsers
      //                      .Where(p => p.FK_User == User.usr_ID)
      //                      .Select(p => p.FK_ReffrenceAcceptLevel)
      //                      .ToList();

      //                  List<tbReffrenceSaveLevel> save = new List<tbReffrenceSaveLevel>();
      //                  List<tbReffrenceAccept> save2 = new List<tbReffrenceAccept>();

      //                  List<int> ID = new List<int>();
      //                  List<int> ID2 = new List<int>();

      //                  foreach (var item in exist_usr)
      //                  {
      //                      var exist_pymn = db.tbReffrenceAcceptLevel
      //                          .Where(p => p.ID == item)
      //                          .Select(p => p.FK_ReffrenceAccept)
      //                          .FirstOrDefault();

      //                      if (exist_pymn != null)
      //                      {
      //                          ID.Add(exist_pymn.Value);
      //                      }
      //                  }

      //                  foreach (var item in ID)
      //                  {
      //                      var exist_py = db.tbReffrenceAccept
      //                          .Where(p => p.ID == item)
      //                          .FirstOrDefault();

      //                      if (exist_py != null)
      //                      {
      //                          save2.Add(exist_py);
      //                      }
      //                  }

      //                  foreach (var item in save2)
      //                  {
      //                      var exist_py = db.tbReffrenceSaveLevel
      //                          .Where(p => p.ID == item.FK_ReffrenceSaveLevel)
      //                          .FirstOrDefault();

      //                      if (exist_py != null)
      //                      {
      //                          save.Add(exist_py);
      //                      }
      //                  }

      //                  foreach (var item in save)
      //                  {
      //                      var exist_py = db.tbReffrenceSaveLevel
      //.Where(p => p.FK_RRSave == p.tbReffrenceSave.ID && p.tbReffrenceSave.IsFor == 1)
      //.FirstOrDefault();


      //                      if (exist_py != null)
      //                      {
      //                          save3.Add(exist_py);
      //                      }
      //                  }

      //                  // Materialize the query results before disposing of the context
      //                  return View("~/Areas/Contracts/Views/Functions/Pyman_forAccept.cshtml", save3);
      //              }
      //          }

      //          return View();
            
      //  }





        public ActionResult Accerpt_Baste(int peyman_id = 0)
        {
            var exist = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == peyman_id && p.Deleted != true).FirstOrDefault();
            return View("~/Areas/Contracts/Views/Functions/Accerpt_Baste.cshtml", exist);
        }

        public ActionResult view_for_Accept()
        {
            return View();
        }







        [AuthorizeAAA]
        public ActionResult _SubmitFunction(int idd/*int Basteh,int Peyman,DateTime FromDate,DateTime ToDate*/)
        {
            var svdfunc = savedfunctionRepo.Find(idd);
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {

                        if (db.tbFucntionExcel.Where(p => p.FK_UserSubmit == User.usr_ID).FirstOrDefault() != null)
                        {
                            svdfunc.OnlineUserUploadedFile = true;
                        }
                        else
                        {
                            svdfunc.OnlineUserUploadedFile = false;
                        }
                    }
                }
            }

            var Basteh = svdfunc.svdfunc_BastehID;
            var Peyman = svdfunc.svdfunc_pymnID;
            var result = functionexcelRepo.Update().Where(p => p.FK_Basteh == Basteh && p.FK_Peyman == Peyman && p.Funcexcl_DateTime >= svdfunc.svdfunc_FromDate && p.Funcexcl_DateTime <= svdfunc.svdfunc_ToDate).ToList();

            return PartialView("_SubmitFunction", svdfunc);

        }
        public ActionResult showtablesabt()
        {
            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid = User.usr_ID;
                        List<int?> id = db.tbReffrenceSaveLevelUser
                            .Where(p => p.FK_UserID == userid)
                            .Select(p => p.FK_LevelID)
                            .ToList();
                        List<int> idrrsv = new List<int>();
                        List<int> pymn = new List<int>();
                        foreach (var a in id)
                        {
                            var x = db.tbReffrenceSaveLevel.Where(p => p.ID == a && p.Deleted != true).FirstOrDefault();
                            idrrsv.Add(x?.FK_RRSave ?? 0);
                        }
                        foreach (var a in idrrsv)
                        {
                            var x = db.tbReffrenceSave.Where(p => p.ID == a).FirstOrDefault();
                            pymn.Add(x?.FK_PeymanID ?? 0);
                        }
                     
                        return View(NotSubmitedMoalefe2(pymn));
                    }
                }
            }

            return View(new List<tbSavedFunctions>());
        }





        public ActionResult showtablesabtinsabtandedit()
        {
            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid = User.usr_ID;
                        List<int?> id = db.tbReffrenceSaveLevelUser
                            .Where(p => p.FK_UserID == userid)
                            .Select(p => p.FK_LevelID)
                            .ToList();
                        List<int> idrrsv = new List<int>();
                        List<int> pymn = new List<int>();

                        foreach (var a in id)
                        {
                            var x = db.tbReffrenceSaveLevel.Where(p => p.ID == a && p.Deleted != true).FirstOrDefault();
                            idrrsv.Add(x?.FK_RRSave ?? 0);
                        }
                        foreach (var a in idrrsv)
                        {
                            var x = db.tbReffrenceSave.Where(p => p.ID == a).FirstOrDefault();
                            pymn.Add(x?.FK_PeymanID ?? 0);
                        }

                        return View(NotSubmitedMoalefe23(pymn));
                    }
                }
            }

            return View(new List<tbSavedFunctions>());
        }
        public ActionResult LIstSoratSabt(int fk_basteh=0,int Month=0, int Year=0)
        {
            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid = User.usr_ID;
                        List<int?> id = db.tbReffrenceSaveLevelUser
                            .Where(p => p.FK_UserID == userid)
                            .Select(p => p.FK_LevelID)
                            .ToList();
                        List<int> idrrsv = new List<int>();
                        List<int> pymn = new List<int>();

                        foreach (var a in id)
                        {
                            var x = db.tbReffrenceSaveLevel.Where(p => p.ID == a && p.Deleted != true).FirstOrDefault();
                            idrrsv.Add(x?.FK_RRSave ?? 0);
                        }
                        foreach (var a in idrrsv)
                        {
                            var x = db.tbReffrenceSave.Where(p => p.ID == a&&p.FK_PeymanID== fk_basteh && p.IsFor==5).FirstOrDefault();
                            pymn.Add(x?.FK_PeymanID ?? 0);
                        }

                        return View(PYMNSorat(pymn, Month, Year));
                    }
                }
            }

            return View(new List<tbsaveSoratBasteh>());

        }

        public ActionResult ListFIsh()
        {
            int userid = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid = User.usr_ID;
                        List<int?> id = db.tbReffrenceSaveLevelUser
                            .Where(p => p.FK_UserID == userid)
                            .Select(p => p.FK_LevelID)
                            .ToList();
                        List<int> idrrsv = new List<int>();
                        List<int> pymn = new List<int>();

                        foreach (var a in id)
                        {
                            var x = db.tbReffrenceSaveLevel.Where(p => p.ID == a && p.Deleted != true).FirstOrDefault();
                            idrrsv.Add(x?.FK_RRSave ?? 0);
                        }
                        foreach (var a in idrrsv)
                        {
                            var x = db.tbReffrenceSave.Where(p => p.ID == a && p.IsFor == 7).FirstOrDefault();
                            pymn.Add(x?.FK_PeymanID ?? 0);
                        }

                        return View(PymnFish(pymn));
                    }
                }
            }

            return View(new List<tbmoalfefishexcel>());

        }




        public ActionResult FilterSabt_Usr(int fk_py = 0, int fk_basteh = 0, int Month = 0, int Year = 0)
        {
            List<tbSavedFunctions> id = new List<tbSavedFunctions>();

            var x = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == fk_py && p.svdfunc_BastehID == fk_basteh && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Month == Month && m.MoalfeVal_Year == Year)).ToList();
            if (x.Count!=0)
            {
           
               
                    foreach (var item in x)
                    {

                        id.Add(item);


                    }
                    return View(id);

                
              
            }

            return View(new List<tbSavedFunctions>());

        }








        public ActionResult _SubmitFunctioninsabtandedit(int idd/*int Basteh,int Peyman,DateTime FromDate,DateTime ToDate*/)
        {
            var svdfunc = savedfunctionRepo.Find(idd);
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {

                        if (db.tbFucntionExcel.Where(p => p.FK_UserSubmit == User.usr_ID).FirstOrDefault() != null)
                        {
                            svdfunc.OnlineUserUploadedFile = true;
                        }
                        else
                        {
                            svdfunc.OnlineUserUploadedFile = false;
                        }
                    }
                }
            }

            var Basteh = svdfunc.svdfunc_BastehID;
            var Peyman = svdfunc.svdfunc_pymnID;
            var result = functionexcelRepo.Update().Where(p => p.FK_Basteh == Basteh && p.FK_Peyman == Peyman && p.Funcexcl_DateTime >= svdfunc.svdfunc_FromDate && p.Funcexcl_DateTime <= svdfunc.svdfunc_ToDate).ToList();

            return PartialView("_SubmitFunctioninsabtandedit", svdfunc);

        }
    public ActionResult TableFordetaiusr(int idd)
        {
            var svdfunc = savedfunctionRepo.Find(idd);
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {

                        if (db.tbFucntionExcel.Where(p => p.FK_UserSubmit == User.usr_ID).FirstOrDefault() != null)
                        {
                            svdfunc.OnlineUserUploadedFile = true;
                        }
                        else
                        {
                            svdfunc.OnlineUserUploadedFile = false;
                        }
                    }
                }
            }

            var Basteh = svdfunc.svdfunc_BastehID;
            var Peyman = svdfunc.svdfunc_pymnID;
            var result = functionexcelRepo.Update().Where(p => p.FK_Basteh == Basteh && p.FK_Peyman == Peyman && p.Funcexcl_DateTime >= svdfunc.svdfunc_FromDate && p.Funcexcl_DateTime <= svdfunc.svdfunc_ToDate).ToList();

            return PartialView( svdfunc);
        }



        public ActionResult pavastperson(IEnumerable<HttpPostedFileBase> files = null, int id = 0)
        {
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();

            // Retrieve the record by ID
            var record = db.tbSavedFunctions.Find(id);

            if (record != null)
            {
                if (files != null)
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));

                            // Save the values to the existing record
                            record.svdfunc_FileSystemName = filename;
                            record.svdfunc_FileName = file.FileName;
                        }
                    }

                    // Save changes to the database
                    db.SaveChanges();
                    return Content("True");

                }
            }
            return View();
        }

        public List<tbSavedFunctions> NotSubmitedMoalefe2(List<int> inputList)
        {
           


            return db.tbSavedFunctions.Where(p => inputList.Contains((int)p.svdfunc_pymnID)).ToList();

        }

        public List<tbSavedFunctions> NotSubmitedMoalefe23(List<int> inputList)
        {
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);


            var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
            if (User != null)
            {
                userid = User.usr_ID;
            }


            return db.tbSavedFunctions.Where(p => inputList.Contains((int)p.svdfunc_pymnID) && p.svdfunc_UserSaveID == userid).ToList();

        }
        public List<tbsaveSoratBasteh> PYMNSorat(List<int> inputList, int Month,int Year)
        {
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);


            var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
            if (User != null)
            {
                userid = User.usr_ID;
            }


            return db.tbsaveSoratBasteh.Where(p => inputList.Contains((int)p.fk_pymn)&&p.tbSoratSavefromExcelMoalfe.Any(s=>s.Month==Month&&s.Year==Year)).ToList();

        }


        public List<tbmoalfefishexcel> PymnFish(List<int> inputList)
        {
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);


            var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
            if (User != null)
            {
                userid = User.usr_ID;
            }


            return db.tbmoalfefishexcel.Where(p => inputList.Contains((int)p.Fk_Pymn) ).ToList();

        }





        public ActionResult showFunctionperson()
        {
            return View();
        }
        public ActionResult showtaki()
        {
            return View();
        }

        [AuthorizeAAA]
        public string UploadAcceptFile(int SaveFunction_ID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            if (SaveFunction_ID != 0)
            {
                var sv = db.tbSavedFunctions.Find(SaveFunction_ID);
                if (sv != null)
                {
                    tbFucntionExcel obj = new tbFucntionExcel();
                    obj.FK_Basteh = sv.svdfunc_BastehID;
                    obj.FK_Peyman = sv.svdfunc_pymnID;
                    obj.FK_SavedFunction = SaveFunction_ID;
                    var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                    if (cookie_user != null)
                    {
                        var nt1 = Utility.Base64.Base64Decode(cookie_user.Value);
                        using (SaabEntities db = new SaabEntities())
                        {
                            var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nt1).FirstOrDefault();
                            if (User != null)
                            {
                                obj.FK_UserSubmit = User.usr_ID;

                            }
                        }
                    }
                    
                    obj.Funcexcl_DateTime = DateTime.Now;

                    if (files != null)
                    {
                        foreach (var file in files)
                        {

                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + "_Taeed." + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                obj.Funcexcl_ExcelName = filename;
                                return functionexcelRepo.Create(obj).ToString();
                            }
                        }
                    }
                }
            }
            return "false";

        }




        public string UploadAcceptFileForaccept(IEnumerable<HttpPostedFileBase> files = null, int svdfunc_BastehID=0, int svdfunc_pymnID=0,int year_pY=0,int month_pY=0)
        {
            var latestSaveFunction = db.tbSavedFunctions.OrderByDescending(sf => sf.svdfunc_ID).FirstOrDefault();
            var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == svdfunc_BastehID &&
     p.svdfunc_pymnID == svdfunc_pymnID &&
     p.tbMoalefeDastmozdiValueFromExcel
         .Any(m => m.MoalfeVal_Year == year_pY && m.MoalfeVal_Month == month_pY)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();
            if (sabtt.svdfunc_ID != 0)
            {
                var sv = db.tbSavedFunctions.Find(sabtt.svdfunc_ID);
                if (sv != null)
                {
                    tbFucntionExcel obj = new tbFucntionExcel();
                    obj.FK_Basteh = sv.svdfunc_BastehID;
                    obj.FK_Peyman = sv.svdfunc_pymnID;
                    obj.FK_SavedFunction = sabtt.svdfunc_ID;

                    // حالا می‌توانید obj را به دیتابیس ذخیره کنید
                    
                    var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                    if (cookie_user != null)
                    {
                        var nt1 = Utility.Base64.Base64Decode(cookie_user.Value);
                        using (SaabEntities db = new SaabEntities())
                        {
                            var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nt1).FirstOrDefault();


                            var exist=db.tbReffrenceAccept.Where(p=>p.FK_ReffrenceSaveLevel== svdfunc_BastehID&&p.FK_PeymanID== svdfunc_pymnID).FirstOrDefault();
                            if (exist != null)
                            {
                                var exist_us = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist.ID).OrderByDescending(p => p.ID).FirstOrDefault();
                                if (exist_us != null)
                                {
                                    var exist_final = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).Select(p=> p.FK_User).FirstOrDefault();
                                    if (exist_final == User.usr_ID)
                                    {



                                        var exxistpy = db.tbReffrenceSaveLevel.Where(p => p.ID == svdfunc_BastehID && p.Deleted != true).Select(p => p.FK_RRSave).FirstOrDefault();
                                        if (exxistpy != null)
                                        {
                                            var existpy = db.tbReffrenceSave.Where(s => s.ID == exxistpy).Select(s => s.IsFor).FirstOrDefault();
                                            if (existpy ==1)
                                            {
                                                var matchedRows = db.tbMoalefeValuePishkhan
                              .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-تایید عملکرد پیمان")
                              .ToList();

                                                foreach (var row in matchedRows)
                                                {
                                                    row.mlfval_Value = "1";
                                                }

                                                db.SaveChanges();
                                            }
                                        }
                                      
                                        // در اینجا شرط برقرار است

                                        // دریافت شئی متناظر با exist_final از جدول tbSavedFunctions

                                        var savedFunction = db.tbSavedFunctions.Where(sf => sf.svdfunc_pymnID == svdfunc_pymnID&&sf.svdfunc_BastehID== svdfunc_BastehID&&sf.For_Accept==true).OrderByDescending(p => p.svdfunc_ID).FirstOrDefault();
                                        if (savedFunction != null)
                                        {
                                            savedFunction.Final_Accept = true;


                                            


                                                // ذخیره تغییرات در دیتابیس
                                                db.SaveChanges();
                                            
                                        }
                                        var sabtfunction = db.tbSavedFunctions.Where(sf => sf.svdfunc_pymnID == svdfunc_pymnID && sf.svdfunc_BastehID == svdfunc_BastehID   && sf.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year_pY) &&
                 sf.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Month == month_pY)).ToList();

                                        if (sabtfunction != null)
                                        { 
                                        foreach(var item in sabtfunction)
                                            {
                                                item.svdfunc_IsSubmmit = true;
                                                db.SaveChanges();
                                            }
                                        
                                        }

                                        }

                                }
                            }
                            if (User != null)
                            {
                                obj.FK_UserSubmit = User.usr_ID;

                            }
                        }
                    }

                    obj.Funcexcl_DateTime = DateTime.Now;

                    if (files != null)
                    {
                        foreach (var file in files)
                        {

                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + "_Taeed." + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                obj.Funcexcl_ExcelName = filename;
                                return functionexcelRepo.Create(obj).ToString();
                            }
                        }
                    }
                }
            }
            return "false";

        }


        #endregion

    }
} 
