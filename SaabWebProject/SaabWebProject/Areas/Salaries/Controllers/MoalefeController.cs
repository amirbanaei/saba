using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.Salaries.MoalefeKarkardi;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Areas.Contracts.Models.Classes;
using SaabWebProject.Utility;
using Telerik.Web.Spreadsheet;
using static SaabWebProject.Areas.Users.Controllers.MessageBoxController;
using SaabWebProject.Models.Repositories.Contracts;
using Syncfusion.XlsIO;
using SaabWebProject.Areas.Contracts.Controllers;
using SaabWebProject.Models.Classes;
using SaabWebProject.Models.Functions.Salaries.Functions;

using SaabWebProject.Models.ViewModels.Contracts.Function;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using static Stimulsoft.Report.StiRecentConnections;
using Stimulsoft.Blockly.Model;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using System.Diagnostics.Metrics;
using Org.BouncyCastle.Asn1.Ocsp;
using Syncfusion.XlsIO.Implementation;
using ExcelLibrary.BinaryFileFormat;
using OfficeOpenXml;
using System.Web.UI.WebControls;
using SaabWebProject.Models.Repositories.Salaries.Formula;
using System.Reflection;
using Stimulsoft.Controls.Wpf.ControlsV3;
using System.Runtime.ConstrainedExecution;
using System.Threading.Tasks;
using System.Threading;
using SaabWebProject.Models.ViewModels.Statements;
using System.Configuration;
using System.Data.SqlClient;
using Stimulsoft.Report.Chart;
using Stimulsoft.Controls.Win.DotNetBar;
using System.Globalization;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Ajax.Utilities;
using System.Security.Cryptography;
using static SaabWebProject.Models.ViewModels.Contracts.Function.FunctionModel;
using System.Text;
using System.Data.Entity.Core.EntityClient;
using Microsoft.SqlServer.Management.Smo;
using DocumentFormat.OpenXml.InkML;
using Microsoft.SqlServer.Management.Smo.Agent;
using System.Windows.Media.Animation;
// Add the appropriate namespaces related to Excel or worksheets


// ReSharper disable All

namespace SaabWebProject.Areas.Salaries.Controllers
{
    public class MoalefeController : Controller
    {
        ZaribForFishRepository ZaribforFish = new ZaribForFishRepository();

        #region متغییر

        SaabEntities db;
        tbUsersRepository userRepo;
        tbCategoriesRepository CatRepo;
        tbContractMoalefeDastMozdiRepository MoalefeRepo;
        tbMoalefeDastmozdiValueFromExcelRepository moalefeexcelRepo;
        tbSoratSavefromExcelMoalfeRepository moalefeexcelRepoSorat = new tbSoratSavefromExcelMoalfeRepository();

        tblink_moalfe_city_valuRepository link_Maolf_city_valu = new tblink_moalfe_city_valuRepository();

        tblink_moalfe_cityRepository link_Maolf_city = new tblink_moalfe_cityRepository();
        tbsaveSoratBastehRepository moaleSorat = new tbsaveSoratBastehRepository();


        tbSavedFunctionsRepositories savedfunctionRepo;
        private tbContractMoalefeDastMozdiRepository rep_moalefeDastmozdi;
        FileContentResult file;
        RegistrationAndConfirmationProceduresController registerController;
        LogFunction logfunc;
        tbFunctionExcelRepository functionexcelRepo; 
        #endregion

        #region سازنده

        public MoalefeController()
        {
            db = new SaabEntities();
            userRepo = new tbUsersRepository(db);
            MoalefeRepo = new tbContractMoalefeDastMozdiRepository(db);
            moalefeexcelRepo = new tbMoalefeDastmozdiValueFromExcelRepository(db);
            savedfunctionRepo = new tbSavedFunctionsRepositories(db);
            rep_moalefeDastmozdi = new tbContractMoalefeDastMozdiRepository(db);
            CatRepo = new tbCategoriesRepository(db);
            registerController = new RegistrationAndConfirmationProceduresController();
            logfunc = new LogFunction(db);
            functionexcelRepo = new tbFunctionExcelRepository(db);
        }
        //private static Timer _timer;

        //protected override void OnActionExecuting(ActionExecutingContext filterContext)
        //{
        //    StartTimer();
        //    base.OnActionExecuting(filterContext);
        //}

        #endregion

        #region صفحات
        [AuthorizeAAA]
        public ActionResult Create()
        {
            return View();
        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult SaveSetting(List<settingview> model)
        {
            try
            {
                var counter = 0;
                foreach (var item in model)
                {
                    int id;
                    if (int.TryParse(item.Moalefe_ID, out id))
                    {
                        var moalefe = db.tbContractMoalefeDastmozdi.Where(p => p.md_ID == id).FirstOrDefault();
                        if (moalefe != null)
                        {
                            if (moalefe.SettingVisableShowMoalefe.Count() == 1)
                            {
                                SettingVisableShowMoalefe setting = moalefe.SettingVisableShowMoalefe.ToList().LastOrDefault();
                                setting.Status = item.Status;
                                setting.Moalefe_ID = id.ToString();
                                setting.FK_Moalefe_ID = id;
                            }
                            else
                            {
                                SettingVisableShowMoalefe setting = new SettingVisableShowMoalefe();
                                setting.Status = item.Status;
                                setting.Moalefe_ID = id.ToString();
                                setting.FK_Moalefe_ID = id;
                                moalefe.SettingVisableShowMoalefe.Add(setting);
                            }
                        }
                    }
                    else
                    {
                        var static_moalefe = db.SettingVisableShowMoalefe.Where(p => p.Moalefe_ID == item.Moalefe_ID).FirstOrDefault();
                        if (static_moalefe != null)
                        {
                            static_moalefe.Status = item.Status;
                        }
                    }

                    counter++;
                }

                return Content(System.Convert.ToBoolean(db.SaveChanges()).ToString());
            }
            catch (Exception e)
            {
                return Content("false");
            }
        }
        [AuthorizeAAA]
        public ActionResult SettingMoalefeForFormula()
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            var ListMoalefe = db.SettingVisableShowMoalefe.ToList();
            foreach (var item in ListMoalefe)
            {
                switch (item.Moalefe_ID)
                {
                    case "usc_CountOfChild":
                        {
                            myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild", visable = item.Status });
                            break;
                        }
                    case "usc_TypeOfContract":
                        {
                            myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", visable = item.Status });
                            break;
                        }
                    case "usc_StartTime":
                        {
                            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime", visable = item.Status });
                            break;
                        }
                    case "usc_EndTime":
                        {
                            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime", visable = item.Status });
                            break;
                        }
                    case "usc_DurationTime":
                        {
                            myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime", visable = item.Status });
                            break;
                        }
                }
            }


            foreach (var item in rep_moalefeDastmozdi.Update())
            {
                string type = "قراردادی";
                switch (item.md_Type)
                {
                    case 1:
                        type = "دستمزدی";
                        break;
                    case 2:
                        type = "کارکردی";
                        break;
                    case 3:
                        type = "سایر";
                        break;
                    default:
                        type = "قراردادی";
                        break;
                }

                var setting = item.SettingVisableShowMoalefe.ToList().LastOrDefault();
                var a = new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type };
                if (setting != null)
                {
                    a.visable = (bool)setting.Status;
                }

                myList.Add(a);
            }

            return View(myList);
        }

        [AuthorizeAAA]
        public ActionResult ManageMoalefeKarkardi()
        {
            UserAndMoalefeVM Model = new UserAndMoalefeVM();
            var usrlst = userRepo.Update();
            var moalefekarkardi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 2).ToList();
            Model.MaolefeKarkardilst = moalefekarkardi;
            Model.Userlst = usrlst;
            return View(Model);
        }

        [AuthorizeAAA]
        public ActionResult ShowSalaryForUsers()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult CitiesZarib()
        {
            return View();
        }

        #endregion


        #region توابع 











        //private static Timer _timer;

        //public static void StartTimer()
        //{
        //    // ایجاد یک تایمر که هر 3 ساعت یکبار اجرا شود
        //    _timer = new Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromHours(3));
        //}


        [AuthorizeAAA]
        [HttpGet]




        public ActionResult sabtpersonalabzar(int ID)
        {
            var OutPutFile = sabtabzar2(ID);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];




            return File(stream.ToArray(), mimeType, "Moalefe" + extension);

        }
        public Workbook sabtabzar2(int ID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/personalabzar.xlsx"));
            Row Row;
            //var MoalefeKarkardi = db.tbEquipmentSpecifications.Where(p => p.FK_EqpBunch_ID == ID).ToList();

            //var count = 10;
            //foreach (var item in MoalefeKarkardi)
            //{
            //    Row = new Row() { Height = 20, Index = 0 };
            //    Row.AddCells(new List<Cell>()
            //    {
            //        new Cell()
            //        {
            //            Value = item.sp_SpecTitle,
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = count
            //        },
            //    });
            //    count++;
            //    Moalefeexcelfile.Sheets[0].AddRow(Row);
            //}

            //باید از کامنت خارج شود
            //List<string> MoalefeKarkardi2 = new List<string>();
            //foreach (var item in MoalefeID)
            //{
            //    MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            //}

            var count2 = 3;
            foreach (var item in db.tbEquipmentBunch.Where(p => p.FK_EqpgrpID == 4).ToList())
            {
                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Eqpbnch_Name,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count2
                    },
                });
                count2++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }


            int counter = 1;
            //List<tbUsers> UserList = new List<tbUsers>();
            //foreach (var item in usrID)
            //{
            //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    UserList.Add(user);
            //}

            foreach (var item in db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == ID && p.Status == true).ToList())
            {
                //نام ، نام خانوادگی و کد پرسنلی
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = item.tbUsers.usr_Name,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 0
                },
                new Cell()
                {
                    Value = item.tbUsers.usr_Family,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
                new Cell()
                {
                    Value = item.tbUsers.usr_Personal_ID,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 2
                }
            });
                    //مقادیر مولفه ها
                    //int index = 3;
                    //foreach (var item2 in MoalefeKarkardi)
                    //{
                    //    var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                    //    && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate && p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                    //    Row.AddCells(new List<Cell>()
                    //    {
                    //         new Cell()
                    //    {
                    //        Value = moalefevalue,
                    //        FontFamily = "B Nazanin",
                    //        Bold = false,
                    //        Enable = true,
                    //        Wrap = false,
                    //        FontSize = 12,
                    //        Italic = false,
                    //        Underline = false,
                    //        Index = index
                    //    }


                    //     });
                    //    index++;
                    //}
                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }




            return Moalefeexcelfile;

        }



        public ActionResult sabtabzarmashin(int ID)
        {
            var OutPutFile = sabtabzar(ID);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];




            return File(stream.ToArray(), mimeType, "equpa" + extension);

        }


        public ActionResult sabtabzarmashinorabzar(int ID)
        {
            var OutPutFile = sabtabzarmashinorabzar2(ID);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];




            return File(stream.ToArray(), mimeType, "abzarr" + extension);

        }


        public Workbook sabtabzar(int ID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/MASHUNABZAR.xlsx"));
            Row Row;
            var MoalefeKarkardi = db.tbEquipmentSpecifications.Where(p => p.FK_EqpBunch_ID == ID).ToList();

            var count = 14;
            var count1 = 1;

            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.sp_SpecTitle,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            foreach (var item in db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList())
            {
                Row = new Row() { Height = 20, Index = count1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.FullName,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 0
                    },
                         new Cell()
                    {
                        Value = item.usr_Personal_ID,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 1
                    },
                });
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            return Moalefeexcelfile;

        }
        public Workbook sabtabzarmashinorabzar2(int ID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzarexcel2.xlsx"));
            Row Row;
            var MoalefeKarkardi = db.tbEquipmentBunch.Where(p => p.FK_EqpgrpID == ID).ToList();

            var count = 14;
            var count1 = 1;
            var cont = 1;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = cont };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Eqpbnch_Name,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 0
                    },
                });
                cont++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            return Moalefeexcelfile;

        }
        public ActionResult ExportMoalefeExcel(int year, int month, DateTime FromDate, DateTime ToDate, string PeymanID, int Basteh)
        {
            var pymn = PeymanID.Split(',');
            int number = int.Parse(PeymanID);

            var find = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == number && p.Final_Accept == true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year && m.MoalfeVal_Month == month)).FirstOrDefault();

            if (find != null)
            {
                return Content("False");
            }

            else
            {






                List<int?> listpymn = new List<int?>();
                for (int i = 0; i < pymn.Count(); i++)
                {
                    listpymn.Add(System.Convert.ToInt32(pymn[i]));
                }
                var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToList();
                var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);
                var OutPutFile = SetDataExcel_Moalefe2(year, month, usrID, MoalefeID, FromDate, ToDate, Basteh, number);
                string extension = ".xlsx";

                var stream = new MemoryStream();
                OutPutFile.Save(stream, extension);
                var mimeType = MimeTypes.ByExtension[extension];




                return File(stream.ToArray(), mimeType, "Moalefe" + extension);
            }
        }



        public ActionResult ExportMoalefeExceledit(int year, int month,int  id)
        {
            //var pymn = PeymanID.Split(',');
            //int number = int.Parse(PeymanID);

            //var find = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == number && p.Final_Accept == true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year && m.MoalfeVal_Month == month)).FirstOrDefault();

            //if (find != null)
            //{
            //    return Content("False");
            //}

           






               
                var OutPutFile = SetDataExcel_Moalefe2edit(year, month,id);
                string extension = ".xlsx";

                var stream = new MemoryStream();
                OutPutFile.Save(stream, extension);
                var mimeType = MimeTypes.ByExtension[extension];




                return File(stream.ToArray(), mimeType, "sabtandtaeed" + extension);
            
        }

        public async Task< ActionResult> ExportMoalefeExceleditersal(int year, int month, int id)
        {
            //var pymn = PeymanID.Split(',');
            //int number = int.Parse(PeymanID);

            //var find = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == number && p.Final_Accept == true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year && m.MoalfeVal_Month == month)).FirstOrDefault();

            //if (find != null)
            //{
            //    return Content("False");
            //}
            int usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {

                    var User = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefaultAsync();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                    }
                    }
            }

                    var tbmarahesabt4 = db.tbmarahesabt4.FirstOrDefault(p => p.ID == id);
            if (tbmarahesabt4 == null)
            {
                return HttpNotFound(); // اگر tbmarahesabt4 پیدا نشد، خطا بدهد
            }

            if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
                .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 1) == true)
            {

                var x = await db.tbSavedFunctions
                      .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && p.svdfunc_UserSaveID == usr&&p.del!=true&&p.FK_tbmarahesabt4==tbmarahesabt4.ID
                                  && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year))
                      .FirstOrDefaultAsync();
                var x1 = await db.tbSavedFunctions
                    .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true && p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && p.FK_tbmarahesabt4 == tbmarahesabt4.ID
                                && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year)).OrderByDescending(p => p.svdfunc_ID)
                    .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.submitt = true;
                    db.SaveChanges();
                }
                else if (x1 != null)
                {
                    var find = await db.tbMoalefeDastmozdiValueFromExcel
                  .Where(s => s.FK_SavedFunctionsID == x1.svdfunc_ID)
                  .ToListAsync();

                    tbSavedFunctions tbSavedFunctions = new tbSavedFunctions
                    {
                        FK_Basteh = x1.FK_Basteh,
                        city = x1.city,
                        FK_tbmarahesabt4 = x1.FK_tbmarahesabt4,
                        Final_Accept = x1.Final_Accept,
                        FK_tbSavedFunctions = x1.FK_tbSavedFunctions,
                        svdfunc_UserSaveID = usr,
                        svdfunc_SavedDateTime = DateTime.Now,
                        svdfunc_pymnID = x1.svdfunc_pymnID,
                        svdfunc_FileSystemName = x1.svdfunc_FileSystemName,

                        svdfunc_FileName = x1.svdfunc_FileName,
                        For_Accept = x1.For_Accept,
                        Datatmie = DateTime.Now,
                        submitt = true,


                        // انتقال سایر فیلدهای x1 در صورت نیاز
                    };

                    db.tbSavedFunctions.Add(tbSavedFunctions);
                    await db.SaveChangesAsync();

                    tbSavedFunctions.submitt = true;
                    tbSavedFunctions.svdfunc_UserSaveID = usr;

                    await db.SaveChangesAsync();

                    List<tbMoalefeDastmozdiValueFromExcel> tbMoalefeDastmozdiValueFromExcellist = new List<tbMoalefeDastmozdiValueFromExcel>();

                    foreach (var t in find)
                    {
                        tbMoalefeDastmozdiValueFromExcel tbMoalefeDastmozdiValueFromExcel = new tbMoalefeDastmozdiValueFromExcel
                        {
                            FK_SavedFunctionsID = tbSavedFunctions.svdfunc_ID,
                            MoalfeVal_Month = t.MoalfeVal_Month,
                            MoalfeVal_Year = t.MoalfeVal_Year,
                            MoalfeVal_FKUser=t.MoalfeVal_FKUser,
                            MoalfeVal_FKMoalafeDastmozdi=t.MoalfeVal_FKMoalafeDastmozdi,
                            MoalfeVal_Value=t.MoalfeVal_Value,
                            // انتقال سایر مقادیر t به شیء جدید
                        };
                        tbMoalefeDastmozdiValueFromExcellist.Add(tbMoalefeDastmozdiValueFromExcel);
                    }

                    db.tbMoalefeDastmozdiValueFromExcel.AddRange(tbMoalefeDastmozdiValueFromExcellist);
                    await db.SaveChangesAsync();

                }
                else
                {
                    return Content("بسته ای در این ماه و سال ثبت نشده است ");

                }
            }
            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
              .Any(s => s.FK_taghiz != null) == true)
            {
                var x = await db.tbEquipmentMoalefeValueReffrenceSave
                      .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.FK_User == usr && p.del != true && p.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                                  && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year))
                      .FirstOrDefaultAsync();
                var x1 = await db.tbEquipmentMoalefeValueReffrenceSave
                   .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true &&p.FK_pymn== tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn&&p.city== tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                               && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year)).OrderByDescending(p=>p.ID)
                   .FirstOrDefaultAsync();
                var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
                var sabt = tbmarhesabt3.marahelsabt;
                var taeed = tbmarhesabt3.taeed;
                int vaziat = 0;
                int vaziat2 = 0;
                if (tbmarahesabt4.sabt == true)
                {
                    vaziat = 1;
                    vaziat2 = (int)tbmarahesabt4.number;

                }
                else if (tbmarahesabt4.taeed == true)
                {
                    vaziat = 2;
                    vaziat2 = (int)tbmarahesabt4.number;
                }

                bool foraccept = false;
                bool Final_Accept = false;

                if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                {
                    Final_Accept = true;
                    foraccept = true;

                }
                else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                {
                    Final_Accept = false;
                    foraccept = true;

                }
                else
                {
                    Final_Accept = false;
                    foraccept = false;
                };






                if (x != null)
                {
                    x.submitt = true;
                    db.SaveChanges();
                }
              
                else if (x1 != null)
                {
                    var find = await db.tbEquipmentMoalefeValue
                  .Where(s => s.FK_tbEquipmentMoalefeValueReffrenceSave == x1.ID)
                  .ToListAsync();
                    tbEquipmentMoalefeValueReffrenceSave tbEquipmentMoalefeValueReffrenceSave = new tbEquipmentMoalefeValueReffrenceSave
                    {
                        FK_Basteh = x1.FK_Basteh,
                        city = x1.city,
                        FK_tbmarahesabt4 = tbmarahesabt4.ID,
                        Final_Accept = Final_Accept,
                        FK_tbEquipmentMoalefeValueReffrenceSave = x1.FK_tbEquipmentMoalefeValueReffrenceSave,
                        FK_User = usr,
                        DateTime = DateTime.Now,
                        FK_pymn = x1.FK_pymn,
                        File_SystemName = x1.File_SystemName,
                   
                    FileName = x1.FileName,
                        For_Acceot = foraccept,
                        Datatmie = DateTime.Now,
                        submitt = true,

                        
                        // انتقال سایر فیلدهای x1 در صورت نیاز
                    };

                    db.tbEquipmentMoalefeValueReffrenceSave.Add(tbEquipmentMoalefeValueReffrenceSave);
                    await db.SaveChangesAsync();


               

                    List<tbEquipmentMoalefeValue> tbEquipmentMoalefeValuelist = new List<tbEquipmentMoalefeValue>();

                    var newValues = find.Select(t => new tbEquipmentMoalefeValue
                    {
                        FK_tbEquipmentMoalefeValueReffrenceSave = tbEquipmentMoalefeValueReffrenceSave.ID,
                        Month = t.Month,
                        Year = t.Year,
                        FK_Equipment = t.FK_Equipment,
                        Fk_pymn = t.Fk_pymn,
                        Fk_user = t.Fk_user,
                        CountDays = t.CountDays
                    }).ToList();

                    if (newValues.Any())
                    {
                        db.tbEquipmentMoalefeValue.AddRange(newValues);
                        await db.SaveChangesAsync();
                    }


                    //foreach (var t in find)
                    //{
                    //    tbEquipmentMoalefeValue tbEquipmentMoalefeValue = new tbEquipmentMoalefeValue
                    //    {
                    //        FK_tbEquipmentMoalefeValueReffrenceSave = tbEquipmentMoalefeValueReffrenceSave.ID,
                    //        Month = t.Month,
                    //        Year = t.Year,
                    //        FK_Equipment = t.FK_Equipment,
                    //        Fk_pymn = t.Fk_pymn,
                    //        Fk_user = t.Fk_user,
                    //        CountDays = t.CountDays,

                    //        // انتقال سایر مقادیر t به شیء جدید
                    //    };
                    //    tbEquipmentMoalefeValuelist.Add(tbEquipmentMoalefeValue);
                    //}

                    //db.tbEquipmentMoalefeValue.AddRange(tbEquipmentMoalefeValuelist);
                    //await db.SaveChangesAsync();
                }
                else
                {
                    return Content("بسته ای در این ماه و سال ثبت نشده است ");

                }
            }

            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
               .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 2) == true)
            {
                var x = await db.tbmoalfefishexcel
                    .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && p.FK_tbmarahesabt4 == id && p.usr_sabt==usr && p.del != true
                                && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year))
                    .FirstOrDefaultAsync();
                var x1 = await db.tbmoalfefishexcel
                   .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && p.del != true
                               && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year)).OrderByDescending(p => p.ID)
                   .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.submitt = true;
                    db.SaveChanges();
                }
                else if (x1 != null)
                {
                    var find = await db.tbMoalefeValueFish
                  .Where(s => s.FK_EXCel == x1.ID)
                  .ToListAsync();
                    tbmoalfefishexcel tbmoalfefishexcel = new tbmoalfefishexcel
                    {
                        FK_Basteh1 = x1.FK_Basteh1,
                        city = x1.city,
                        FK_tbmarahesabt4 = x1.FK_tbmarahesabt4,
                        Final_accept = x1.Final_accept,
                        FK_tbmoalfefishexcel = x1.FK_tbmoalfefishexcel,
                        usr_sabt = usr,
                        Fk_Pymn = x1.Fk_Pymn,
                        Filsesystem = x1.Filsesystem,

                        FileName = x1.FileName,
                        Datatmie = DateTime.Now,
                        submitt = true,


                        // انتقال سایر فیلدهای x1 در صورت نیاز
                    };

                    db.tbmoalfefishexcel.Add(tbmoalfefishexcel);
                    await db.SaveChangesAsync();




                    List<tbMoalefeValueFish> tbMoalefeValueFishlist = new List<tbMoalefeValueFish>();

                    foreach (var t in find)
                    {
                        tbMoalefeValueFish tbMoalefeValueFish = new tbMoalefeValueFish
                        {
                            FK_EXCel = tbmoalfefishexcel.ID,
                            mlfvlfsh_Month = t.mlfvlfsh_Month,
                            mlfvlfsh_Year = t.mlfvlfsh_Year,
                            FK_Moalefe = t.FK_Moalefe,
                            mlfvlfsh_Value = t.mlfvlfsh_Value,
                            FK_User = t.FK_User,
                            Finalaccept = t.Finalaccept,

                            // انتقال سایر مقادیر t به شیء جدید
                        };
                        tbMoalefeValueFishlist.Add(tbMoalefeValueFish);
                    }

                    db.tbMoalefeValueFish.AddRange(tbMoalefeValueFishlist);
                    await db.SaveChangesAsync();

                    x1.submitt = true;
                    db.SaveChanges();
                }
                else
                {
                    return Content("بسته ای در این ماه و سال ثبت نشده است ");

                }
            }


            else if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 4))
            {

                var x = await db.tbsaveSoratBasteh
                  .Where(p => p.FK_Bastehedit == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.Fk_usrsave == usr && p.del != true
                              && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == month && s.Year == year))
                  .FirstOrDefaultAsync();
                var x1 = await db.tbsaveSoratBasteh
                   .Where(p => p.FK_Bastehedit == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true
                               && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == month && s.Year == year)).OrderByDescending(p => p.ID)
                   .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.submitt = true;
                    db.SaveChanges();
                }
                else if (x1 != null)
                {
                    var find = await db.tbSoratSavefromExcelMoalfe
                  .Where(s => s.FK_SavedFunctionsID == x1.ID)
                  .ToListAsync();
                    tbsaveSoratBasteh tbmoalfefishexcel = new tbsaveSoratBasteh
                    {
                        FK_Bastehedit = x1.FK_Bastehedit,
                        city = x1.city,
                        FK_tbmarahesabt4 = x1.FK_tbmarahesabt4,
                        Final_Accept = x1.Final_Accept,
                        Fk_usrsave = usr,
                        fk_pymn = x1.fk_pymn,
                        FileSystemName = x1.FileSystemName,

                        FileName = x1.FileName,
                        Data_time_edit = DateTime.Now,
                        submitt = true,


                        // انتقال سایر فیلدهای x1 در صورت نیاز
                    };

                    db.tbsaveSoratBasteh.Add(tbmoalfefishexcel);
                    await db.SaveChangesAsync();




                    List<tbSoratSavefromExcelMoalfe> tbMoalefeValueFishlist = new List<tbSoratSavefromExcelMoalfe>();

                    foreach (var t in find)
                    {
                        tbSoratSavefromExcelMoalfe tbMoalefeValueFish = new tbSoratSavefromExcelMoalfe
                        {
                            FK_SavedFunctionsID = tbmoalfefishexcel.ID,
                            Month = t.Month,
                            Year = t.Year,
                            FK_MoalfeDastmozdi = t.FK_MoalfeDastmozdi,
                            valuenergh = t.valuenergh,
                            FK_city = t.FK_city,
                            FK_price = t.FK_price,
                            Value=t.Value,

                            // انتقال سایر مقادیر t به شیء جدید
                        };
                        tbMoalefeValueFishlist.Add(tbMoalefeValueFish);
                    }

                    db.tbSoratSavefromExcelMoalfe.AddRange(tbMoalefeValueFishlist);
                    await db.SaveChangesAsync();

                    x1.submitt = true;
                    db.SaveChanges();
                }
                else
                {
                    return Content("بسته ای در این ماه و سال ثبت نشده است ");

                }

            }

            return Content("True");

            //    var OutPutFile = SetDataExcel_Moalefe2edit(year, month, id);
            //string extension = ".xlsx";

            //var stream = new MemoryStream();
            //OutPutFile.Save(stream, extension);
            //var mimeType = MimeTypes.ByExtension[extension];




            //return File(stream.ToArray(), mimeType, "Moalefe" + extension);

        }

        public async Task<ActionResult> ExportMoalefeExceleditersaldel(int year, int month, int id)
        {
            //var pymn = PeymanID.Split(',');
            //int number = int.Parse(PeymanID);

            //var find = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == number && p.Final_Accept == true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year && m.MoalfeVal_Month == month)).FirstOrDefault();

            //if (find != null)
            //{
            //    return Content("False");
            //}
            int usr = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {

                    var User = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefaultAsync();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                    }
                }
            }
            var tbmarahesabt41111 =await db.tbmarahesabt4.Where(s=>s.del!=true).ToListAsync();
            var tbSavedFunctions = await db.tbSavedFunctions.ToListAsync();
            var tbmoalfefishexcel = await db.tbmoalfefishexcel.ToListAsync();

            var tbmarahesabt4 = db.tbmarahesabt4.FirstOrDefault(p => p.ID == id);
            if (tbmarahesabt4 == null)
            {
                return HttpNotFound(); // اگر tbmarahesabt4 پیدا نشد، خطا بدهد
            }

            if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
                .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 1) == true)
            {
                if (tbmarahesabt4.taeed == true)
                {
                    int vaz = 0;

                    List<tbmarahesabt4> tbmarahesabt4143 = new List<tbmarahesabt4>();
                    List<tbmarahesabt4> tbmarahesabt421 = new List<tbmarahesabt4>();
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true ).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true&&p.number> tbmarahesabt4.number).ToList())
                    {
                        var x1 =  tbSavedFunctions
                                                 .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.svdfunc_UserSaveID == ty1.FK_usr&&p.FK_tbmarahesabt4==ty1.ID
                                                             && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year)&&p.del!=true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Canttaeed");

                        }

                    }

                    var x = await db.tbSavedFunctions
                                         .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.svdfunc_UserSaveID == usr
                                                     && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.del != true)
                                         .FirstOrDefaultAsync();
                    if (x != null)
                    {
                        x.del = true;
                        x.Final_Accept = null;
                        x.Datatmie = DateTime.Now;

                        db.SaveChanges();
                    }
                }
                else if (tbmarahesabt4.sabt == true)
                {
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true).ToList())
                    {
                        var x1 = tbSavedFunctions
                                                 .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.svdfunc_UserSaveID == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Canttaeed");

                        }

                    }


                    var x = await db.tbSavedFunctions
                                         .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.svdfunc_UserSaveID == usr
                                                     && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.del != true)
                                         .FirstOrDefaultAsync();
                    if (x != null)
                    {
                        x.del = true;
                        x.Final_Accept = null;
                        x.Datatmie = DateTime.Now;
                        db.SaveChanges();
                    }
                }
                else if (tbmarahesabt4.control == true)
                {
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true&&p.number> tbmarahesabt4.number).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    var x = await db.tbSavedFunctions
                                         .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.svdfunc_UserSaveID == usr
                                                     && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.del != true)
                                         .FirstOrDefaultAsync();
                    if (x != null)
                    {
                        x.del = true;
                        x.Final_Accept = null;

                        db.SaveChanges();
                    }
                }
            }
            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
              .Any(s => s.FK_taghiz != null) == true)
            {
                var x = await db.tbEquipmentMoalefeValueReffrenceSave
                      .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST&&p.FK_tbmarahesabt4== id && p.FK_User == usr
                                  && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year))
                      .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.Datatmie= DateTime.Now;
                    x.DateTime = DateTime.Now;

                    x.del = true;
                    x.Final_Accept = null;
                    db.SaveChanges();
                   
                }

            }

            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
               .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 2) == true)
            {


                if (tbmarahesabt4.taeed == true)
                {
                    int vaz = 0;

                    List<tbmarahesabt4> tbmarahesabt4143 = new List<tbmarahesabt4>();
                    List<tbmarahesabt4> tbmarahesabt421 = new List<tbmarahesabt4>();
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true && p.number > tbmarahesabt4.number).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                            && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Canttaeed");

                        }

                    }

                   
                }
                else if (tbmarahesabt4.sabt == true)
                {
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Canttaeed");

                        }

                    }


                   
                }
                else if (tbmarahesabt4.control == true)
                {
                    foreach (var ty1 in tbmarahesabt41111.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true && p.number > tbmarahesabt4.number).ToList())
                    {
                        var x1 = tbmoalfefishexcel
                                                 .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == ty1.FK_usr && p.FK_tbmarahesabt4 == ty1.ID
                                                             && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.del != true)
                                                 .FirstOrDefault();
                        if (x1 != null)
                        {
                            return Content("Cantcontrol");

                        }

                    }
                    
                }





                var x = await db.tbmoalfefishexcel
                    .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.usr_sabt == usr && p.FK_tbmarahesabt4 == id
                                && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year)&&p.del!=true)
                    .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.del = true;
                    x.Datatmie =  DateTime.Now;
                    
                    db.SaveChanges();
                    if (x != null)
                    {
                        var listtbMoalefeValueFish = await db.tbMoalefeValueFish.Where(p => p.FK_EXCel == x.ID).ToListAsync();
                        foreach(var it in listtbMoalefeValueFish)
                        {
                            it.Finalaccept = null;
                            await db.SaveChangesAsync();

                        }
                    }
                }
            }



            else if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 4))
            {

                var x = await db.tbsaveSoratBasteh
                    .Where(p => p.FK_Bastehedit == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_tbmarahesabt4 == id && p.Fk_usrsave == usr
                                && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == month && s.Year == year))
                    .FirstOrDefaultAsync();
                if (x != null)
                {
                    x.Data_time_edit = DateTime.Now;

                    x.del = true;
                    x.Final_Accept = null;
                    db.SaveChanges();

                }

            }
            return Content("True");

            //    var OutPutFile = SetDataExcel_Moalefe2edit(year, month, id);
            //string extension = ".xlsx";

            //var stream = new MemoryStream();
            //OutPutFile.Save(stream, extension);
            //var mimeType = MimeTypes.ByExtension[extension];




            //return File(stream.ToArray(), mimeType, "Moalefe" + extension);

        }



        public ActionResult ExceleForsorat(string PeymanID, int Basteh, int year, int month)
        {
            var pymn = PeymanID.Split(',');
            int number = int.Parse(PeymanID);

            var find = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == number && p.Final_Accept == true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year && m.MoalfeVal_Month == month)).FirstOrDefault();







            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.tbpeymancities.Where(p => listpymn.Contains(p.FK_PYMN)).Select(p => p.FK_City).ToList();
            var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);
            var OutPutFile = SetDataExcel_MoalefeForsorat(year, month, usrID, MoalefeID, Basteh, number);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];




            return File(stream.ToArray(), mimeType, "Soratsabt" + extension);

        }
        public ActionResult ExportMoalefeExcelForAccept(int year, int month, string PeymanID, int Basteh)
        {
            var pymn = PeymanID.Split(',');
            int number = int.Parse(PeymanID);
            var us = db.tbReffrenceSave
    .Where(p => p.FK_PeymanID == number)
    .SelectMany(s => s.tbReffrenceSaveLevel)
    .Where(s => s.ID == Basteh)
    .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
     .Where(p => p.FK_RRSave == us.FK_RRSave)
     .OrderByDescending(p => p.ID)
     .FirstOrDefault();

            var rus3 = db.tbReffrenceSaveLevelUser
          .Where(p => p.FK_LevelID == us2.ID)
          .OrderByDescending(p => p.ID)
          .Select(p => p.FK_UserID)
          .FirstOrDefault();
            var us4 = -1;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        us4 = User.usr_ID;
                    }
                }
            }










            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToList();
            var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);
            var OutPutFile = SetDataExcel_MoalefeForAccept(year, month, usrID, MoalefeID, Basteh, number);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];




            return File(stream.ToArray(), mimeType, "Moalefe" + extension);
        }






        public async Task<ActionResult> ExportMoalefeExcel3(int id)



        {
            //این تابع و ایجکسش درست کار میکنه ؟
            //بله 
            var tt = db.tbSavedFunctions.Where(p => p.svdfunc_ID == id).FirstOrDefault();
            if (tt == null)
            {
                // انجام کار مناسب برای مواجهه با عدم یافتن رکورد
                return Content("sdkjljl"); // یا هر نوع پاسخی که برنامه‌ی شما نیاز دارد
            }
            var idpy = tt.svdfunc_pymnID;
            //var pymn = idpy.ToString().Split(',');

            var Baste = tt.svdfunc_BastehID;


            var user2 = tt.svdfunc_UserSaveID;

            var id2 = db.tbSavedFunctions
       .Where(p => p.svdfunc_pymnID == idpy && p.svdfunc_UserSaveID == user2)
       .FirstOrDefault();

            var id3 = id2.svdfunc_ID;
            var month2 = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == id).FirstOrDefault();
            var year2 = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == id).FirstOrDefault();
            int month = (int)month2.MoalfeVal_Month;
            int year = (int)month2.MoalfeVal_Year;
            var from = db.tbSavedFunctions.Where(p => p.svdfunc_ID == id).FirstOrDefault();
            DateTime FromDate = (DateTime)from.svdfunc_FromDate;
            DateTime ToDate = (DateTime)from.svdfunc_ToDate;










            List<int?> listpymn = new List<int?>();

            var usrID = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == idpy && p.Status == true).Select(s => s.FK_User_ID).ToList();

            var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel((int)Baste);
            var OutPutFile = await SetDataExcel_Moalefe(year, month, usrID, MoalefeID, FromDate, ToDate, (int)Baste, id);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];



            //var segment = FileName.Split('.');
            //string file_type = segment[segment.Length - 1];
            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + extension).ToString();

            //file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));

            // Save the values to the existing record
            //record.svdfunc_FileSystemName = filename;
            //record.svdfunc_FileName = file.FileName;
            var directoryPath = Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/");
            var fileName = filename; // نام فایل
            var path = Path.Combine(directoryPath, fileName);

            // ذخیره فایل Excel
            OutPutFile.Save(path);
            var t = db.tbSavedFunctions.Where(p => p.svdfunc_ID == id).FirstOrDefault();
            t.svdfunc_FileSystemNameExcel = filename;
            t.svdfunc_FileNameExcel = filename;
            db.SaveChanges();

            return File(stream.ToArray(), mimeType, "Moalefe" + extension);

        }




















        public async Task<ActionResult> ExportMoalefeExcel2(int? idd, string PeymanID, string Baste2, int id)
        {
            //این تابع و ایجکسش
            //
            //درست کار میکنه ؟
            //بله 
            var idpy = db.tbPeymanContracts.FirstOrDefault(p => p.pec_Title == PeymanID && p.Inactive != true).pec_ID;
            var pymn = idpy.ToString().Split(',');

            var Baste = db.tbReffrenceSaveLevel.FirstOrDefault(p => p.Title == Baste2 && p.Deleted == null && p.Deleted != true && p.tbReffrenceSave.FK_PeymanID == idpy).ID;


            var user2 = idd;

            var id2 = db.tbSavedFunctions
       .Where(p => p.svdfunc_pymnID == idpy && p.svdfunc_UserSaveID == user2)
       .FirstOrDefault();

            var id3 = id2.svdfunc_ID;
            var month2 =await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == id).FirstOrDefaultAsync();
            var year2 = await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == id).FirstOrDefaultAsync();
            int month = (int)month2.MoalfeVal_Month;
            int year = (int)month2.MoalfeVal_Year;
            var from = db.tbSavedFunctions.Where(p => p.svdfunc_ID == id).FirstOrDefault();
            DateTime FromDate = (DateTime)from.svdfunc_FromDate;
            DateTime ToDate = (DateTime)from.svdfunc_ToDate;










            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = await db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToListAsync();

            var MoalefeID =registerController.GetAllMoalefeInReffrenceSaveLevel(Baste);
            var OutPutFile =await SetDataExcel_Moalefe25(year, month, usrID, MoalefeID, FromDate, ToDate, Baste, id);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            ////var directoryPath = Server.MapPath("~/Content/ExcelFiles/");
            ////var fileName = "Moalefe1400.xlsx"; // نام فایل
            ////var path = Path.Combine(directoryPath, fileName);

            //// ذخیره فایل Excel
            //OutPutFile.Save(path);


            return File(stream.ToArray(), mimeType, "Karkard" + extension);
        }


        //public tbMoalefeDastmozdiValueFromExcel GetAll()
        //{
            
        //}

        public async Task<Workbook> SetDataExcel_Moalefe25(int year, int month, IEnumerable<int?> usrID, IEnumerable<int> MoalefeID, DateTime FromDate, DateTime ToDate, int Baste, int id)
        {
            //GetAll Table
            var All_tbMoalefeDastmozdiValueFromExcel = await db.tbMoalefeDastmozdiValueFromExcel.Where(p=>p.FK_SavedFunctionsID!=null).ToListAsync();


            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Moalefe.xlsx"));
            Row Row;

            // Fill year
            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = year,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        },
    });
            Moalefeexcelfile.Sheets[0].AddRow(Row);

            // Fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = month,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        },
    });
            Moalefeexcelfile.Sheets[0].AddRow(Row);

            // Fetch MoalefeKarkardi titles
            var MoalefeKarkardi = await db.tbContractMoalefeDastmozdi
                                        .Where(p => MoalefeID.Contains(p.md_ID))
                                        .Select(p => p.md_Title)
                                        .ToListAsync();

            // Populate MoalefeKarkardi into Excel rows
            var count343 = 3;
            var count34334 = 3;

            foreach (var item in MoalefeID)
            {
                Row = new Row() { Height = 20, Index = 1 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count34334
            },
        });
                count34334++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            var count = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count343
            },
        });
                count343++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            //Dictionary<List<string>,List<string>,List<string>, List<float>> moalefeValuesDict = new Dictionary<List<string>, List<string>, List<string>, List<float>>();

            // Fetch user details
            var UserList = await db.tbUsers.Where(p => usrID.Contains(p.usr_ID)).ToListAsync();
            //foreach (var item in UserList)
            //{

            //    foreach (var moalefeTitle in MoalefeKarkardi)
            //    {
            //        var moalefeValue = await db.tbMoalefeDastmozdiValueFromExcel
            //                                   .Where(p => p.tbContractMoalefeDastmozdi.md_Title == moalefeTitle &&
            //                                               p.MoalfeVal_FKUser == item.usr_ID &&
            //                                               p.tbSavedFunctions.svdfunc_ID == id &&
            //                                               p.tbSavedFunctions.svdfunc_BastehID == Baste)
            //                                   .Select(p => p.MoalfeVal_Value)
            //                                   .FirstOrDefaultAsync();

            //    }

            //}
            //// Populate user data into Excel rows
            //foreach (var item in UserList)
            //{
            //    var tytyy = 0;
            //    var allMoalefeTitles = MoalefeKarkardi.ToList();

            //    // اجرای یک بار کوئری برای همه‌ی عناوین
            //    var moslfetitles = All_tbMoalefeDastmozdiValueFromExcel
            //        .Where(p => allMoalefeTitles.Any(title => p.tbContractMoalefeDastmozdi.md_Title.Contains(title)) &&
            //                    p.MoalfeVal_FKUser == item.usr_ID &&
            //                    p.tbSavedFunctions.svdfunc_ID == id &&
            //                    p.tbSavedFunctions.svdfunc_BastehID == Baste)
            //        .Select(p => p.MoalfeVal_Value)
            //        .ToList();

            //    foreach (var moalefeTitle in MoalefeKarkardi)
            //    {
            //        // اکنون می‌توانید از moslfetitles استفاده کنید بدون اجرای مجدد کوئری
            //        var moslfetitle = moslfetitles.FirstOrDefault(title => title.ToString().Contains(moalefeTitle));
            //        // دستورات بعدی
            //    }
            //}





            // جمع‌آوری همه‌ی عناوین moalefe
     






            foreach (var item in UserList)
            {
                Row = new Row() { Height = 20, Index = count };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.usr_Name,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 0
            },
            new Cell()
            {
                Value = item.usr_Family,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 1
            },
            new Cell()
            {
                Value = item.usr_Personal_ID,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 2
            }
        });
                List<double?> users = new List<double?>();

                // Fetch Moalefe values for each user
                int index = 3;
                //foreach (var moalefeTitle in MoalefeKarkardi)
                //{
                //    users = await db.tbMoalefeDastmozdiValueFromExcel
                //                               .Where(p => p.tbContractMoalefeDastmozdi.md_Title == moalefeTitle &&
                //                                           p.MoalfeVal_FKUser == item.usr_ID &&
                //                                           p.tbSavedFunctions.svdfunc_ID == id &&
                //                                           p.tbSavedFunctions.svdfunc_BastehID == Baste)
                //                               .Select(p => p.MoalfeVal_Value)
                //                               .ToListAsync();

                //}
                foreach (var moalefeTitle in MoalefeKarkardi) {
                    var moslfetitle = All_tbMoalefeDastmozdiValueFromExcel
                                               .Where(p => p.tbContractMoalefeDastmozdi.md_Title .Contains(moalefeTitle) &&
                                                           p.MoalfeVal_FKUser == item.usr_ID &&
                                                           p.tbSavedFunctions.svdfunc_ID == id &&
                                                           p.tbSavedFunctions.svdfunc_BastehID == Baste)
                                               .Select(p => p.MoalfeVal_Value)
                                               .FirstOrDefault();
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = moslfetitle,
                    FontFamily = "B Nazanin",
                    //Bold = false,
                    //Enable = true,
                    //Wrap = false,
                    //FontSize = 12,
                    //Italic = false,
                    //Underline = false,
                    Index = index
                }
            });
                    index++;
                }

                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            return Moalefeexcelfile;
        }


        [AuthorizeAAA]
        public async Task<Workbook> SetDataExcel_Moalefe(int year, int month, List<int?> usrID, List<int> MoalefeID, DateTime FromDate, DateTime ToDate, int Baste, int id)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Moalefe.xlsx"));
            Row Row;


            //fill year
            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);

            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);


            //باید از کامنت خارج شود
            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            }

            var count = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            //name and lastname and groupjob
            int counter = 3;
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }

            foreach (var item in UserList)
            {
                //نام ، نام خانوادگی و کد پرسنلی
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.usr_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Family,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                    //مقادیر مولفه ها
                    int index = 3;
                    foreach (var item2 in MoalefeKarkardi)
                    {
                        var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID && p.tbSavedFunctions.svdfunc_ID == id
                        && /*p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate*/ /*&&*/ p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = moalefevalue,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        index++;
                    }
                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }

            return Moalefeexcelfile;
        }



        public Workbook SetDataExcel_MoalefeForAccept(int year, int month, List<int?> usrID, List<int> MoalefeID, int Baste, int PymanID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Moalefe accept.xlsx"));
            Row Row;

            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);

            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);


            //باید از کامنت خارج شود
            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            }
            var count7 = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 1 };
                Row.AddCells(new List<Cell>()
       {
           new Cell()
           {
               Value = db.tbCaranSettings.Where(p=>p.tbContractMoalefeDastmozdi.md_Title==item).Select(s=>s.CaranStandard).FirstOrDefault(),
               FontFamily = "B Nazanin",
               Bold = false,
               Enable = true,
               Wrap = false,
               FontSize = 12,
               Italic = false,
               Underline = false,
               Index = count7
           },
       }); ;
                count7++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            var count722 = 3;
            foreach (var item in MoalefeID)
            {
                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
       {
           new Cell()
           {
               Value =item,
               FontFamily = "B Nazanin",
               Bold = false,
               Enable = true,
               Wrap = false,
               FontSize = 12,
               Italic = false,
               Underline = false,
               Index = count722
           },
       }); ;
                count722++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            var count = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            //name and lastname and groupjob
            int counter = 3;
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }

            foreach (var item in UserList)
            {
                //نام ، نام خانوادگی و کد پرسنلی
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.usr_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Family,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                    //مقادیر مولفه ها
                    //int index = 3;
                    //foreach (var item2 in MoalefeKarkardi)
                    //{
                    //    var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                    //    && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate && p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                    //    Row.AddCells(new List<Cell>()
                    //    {
                    //         new Cell()
                    //    {
                    //        Value = moalefevalue,
                    //        FontFamily = "B Nazanin",
                    //        Bold = false,
                    //        Enable = true,
                    //        Wrap = false,
                    //        FontSize = 12,
                    //        Italic = false,
                    //        Underline = false,
                    //        Index = index
                    //    }


                    //     });
                    //    index++;
                    //}
                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }

























     //       //fill year
     //       Row = new Row() { Height = 20, Index = 0 };
     //       Row.AddCells(new List<Cell>()
     //       {
     //           new Cell()
     //           {
     //               Value = year,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 1
     //           },
     //       });

     //       Moalefeexcelfile.Sheets[1].AddRow(Row);

     //       //fill month
     //       Row = new Row() { Height = 20, Index = 1 };
     //       Row.AddCells(new List<Cell>()
     //       {
     //           new Cell()
     //           {
     //               Value = month,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 1
     //           },
     //       });

     //       Moalefeexcelfile.Sheets[1].AddRow(Row);


     //       //باید از کامنت خارج شود
     //       //List<string> MoalefeKarkardi = new List<string>();
     //       //foreach (var item in MoalefeID)
     //       //{
     //       //    MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
     //       //}

     //       var countu = 5;
     //       foreach (var item in MoalefeKarkardi)
     //       {
     //           Row = new Row() { Height = 20, Index = 2 };
     //           Row.AddCells(new List<Cell>()
     //           {
     //               new Cell()
     //               {
     //                   Value = item,
     //                   FontFamily = "B Nazanin",
     //                   Bold = false,
     //                   Enable = true,
     //                   Wrap = false,
     //                   FontSize = 12,
     //                   Italic = false,
     //                   Underline = false,
     //                   Index = countu
     //               },
     //           });
     //           countu++;
     //           Moalefeexcelfile.Sheets[1].AddRow(Row);
     //       }

     //       //name and lastname and groupjob
     //       int counter4 = 3;
     //       List<tbUsers> UserList4 = new List<tbUsers>();
     //       foreach (var item in usrID)
     //       {
     //           var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
     //           UserList4.Add(user);
     //       }
     //       var User_create = db.tbSavedFunctions
     //.Where(p => p.svdfunc_BastehID == Baste &&
     //            p.svdfunc_pymnID == PymanID &&
     //            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year) &&
     //            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Month == month))
     //.ToList();

     //       List<tbUsers> UserList_Create = new List<tbUsers>();
     //       foreach (var item in User_create)
     //       {
     //           var user = db.tbUsers.FirstOrDefault(p => p.usr_ID == item.svdfunc_UserSaveID);
     //           if (user != null)
     //           {
     //               UserList_Create.Add(user);
     //           }
     //       }
     //       int startRowIndex = 4;
     //       int endRowIndex = startRowIndex + UserList_Create.Count() - 1;


     //       foreach (var item in UserList4)
     //       {
     //           foreach (var item3 in UserList_Create)
     //           {
     //               var userr2 = db.tbSavedFunctions.Where(p => p.svdfunc_UserSaveID == item3.usr_ID && p.svdfunc_BastehID == Baste && p.For_Accept == true).Select(p => p.svdfunc_ID).FirstOrDefault();
     //               string rol = "";


     //               if (db.tbFucntionExcel.Where(p => p.FK_SavedFunction == userr2).FirstOrDefault() != null)
     //               {
     //                   rol = "تایید کننده";
     //               }
     //               else
     //               {
     //                   rol = "ثبت کننده";
     //               }
     //               // نام ، نام خانوادگی و کد پرسنلی
     //               Row = new Row() { Height = 20, Index = counter4 };
     //               {
     //                   Row.AddCells(new List<Cell>()
     //       {
     //           new Cell()
     //           {
     //               Value = item.usr_Name,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 0
     //           },
     //           new Cell()
     //           {
     //               Value = item.usr_Family,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 1
     //           },
     //           new Cell()
     //           {
     //               Value = item.usr_Personal_ID,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 2
     //           },
     //           new Cell()
     //           {
     //               Value = item3.FullName,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 3
     //           },

     //            new Cell()
     //           {


     //               Value = rol,
     //               FontFamily = "B Nazanin",
     //               Bold = false,
     //               Enable = true,
     //               Wrap = false,
     //               FontSize = 12,
     //               Italic = false,
     //               Underline = false,
     //               Index = 4
     //           }
     //       });

     //                   // مقادیر مولفه ها
     //                   int index = 5;

     //                   foreach (var item2 in MoalefeKarkardi)
     //                   {
     //                       var userr = db.tbSavedFunctions.Where(p => p.svdfunc_UserSaveID == item3.usr_ID && p.svdfunc_BastehID == Baste).Select(p => p.svdfunc_ID).FirstOrDefault();
     //                       var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel
     //                           .Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 &&
     //                                       p.MoalfeVal_FKUser == item.usr_ID &&
     //                                       p.FK_SavedFunctionsID == userr &&
     //                                       p.tbSavedFunctions.svdfunc_BastehID == Baste)
     //                           .Select(p => p.MoalfeVal_Value)
     //                           .FirstOrDefault();

     //                       Row.AddCells(new List<Cell>()
     //           {
     //               new Cell()
     //               {
     //                   Value = moalefevalue,
     //                   FontFamily = "B Nazanin",
     //                   Bold = false,
     //                   Enable = true,
     //                   Wrap = false,
     //                   FontSize = 12,
     //                   Italic = false,
     //                   Underline = false,
     //                   Index = index
     //               }
     //           });
     //                       index++;
     //                   }
     //               }

     //               counter4++;
     //               Moalefeexcelfile.Sheets[1].AddRow(Row);
     //           }


     //           var x = UserList_Create.Count();
     //           if (x > 1)
     //           {
     //               Moalefeexcelfile.Sheets[1].AddMergedCells($"A{startRowIndex}:A{endRowIndex}");
     //               Moalefeexcelfile.Sheets[1].AddMergedCells($"C{startRowIndex}:C{endRowIndex}");
     //               Moalefeexcelfile.Sheets[1].AddMergedCells($"B{startRowIndex}:B{endRowIndex}");
     //               startRowIndex += x;
     //               endRowIndex += x;
     //           }
     //       }






            return Moalefeexcelfile;
        }



        public ActionResult shoWABZAR()
        {
            return View(db.tbEquipmentBunch.Where(p => p.FK_EqpgrpID == 4).ToList());
        }





        public async Task< ActionResult> DELETESORAT(int ID=0)
        {
            var find =await db.tbSoratSavefromExcelMoalfe.Where(p => p.FK_SavedFunctionsID == ID).ToListAsync();
            db.tbSoratSavefromExcelMoalfe.RemoveRange(find);
           await db.SaveChangesAsync();
            var find2 =await db.tbsaveSoratBasteh.Where(p => p.ID == ID).FirstOrDefaultAsync();
            db.tbsaveSoratBasteh.Remove(find2);
           await db.SaveChangesAsync();
            return Content("True");
        }






        public Workbook SetDataExcel_MoalefeForsorat(int year, int month, List<int?> usrID, List<int> MoalefeID, int Baste, int PymanID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelforsorat.xlsx"));
            Row Row;

            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);


            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);



            //باید از کامنت خارج شود
            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            }
            var counttt = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == PymanID).ToList();
           
            var count = 1;
            foreach(var it in counttt)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = it.UnitTitle,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            int counter2 = 3;
            foreach (var item in usrID)
            {



                var x = db.tbCities.Where(p => p.ID == item).FirstOrDefault();
                //نام شهرستان ها 
                Row = new Row() { Height = 20, Index = counter2 };


                Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = x.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                             });

                counter2++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            //name and lastname and groupjob


            int counter = 3;
            List<tbpeymancities> UserList = new List<tbpeymancities>();
            foreach (var item in usrID)
            {
                var user = db.tbpeymancities.Where(p => p.FK_PYMN == PymanID).FirstOrDefault();

                UserList.Add(user);
            }

            //برای شیت اول


            Row = new Row() { Height = 20, Index = 0 };

            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[1].AddRow(Row);


            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[1].AddRow(Row);



            //باید از کامنت خارج شود
            //List<string> MoalefeKarkardi = new List<string>();
            //foreach (var item in MoalefeID)
            //{
            //    MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            //}

            var count4 = 2;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count4
                    },
                });
                count4++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }



            var User_create = db.tbsaveSoratBasteh
.Where(p => p.fk_Basteh == Baste &&
     p.fk_pymn == PymanID &&
     p.tbSoratSavefromExcelMoalfe.Any(m => m.Year == year) &&
     p.tbSoratSavefromExcelMoalfe.Any(m => m.Month == month))
.ToList();
            List<tbUsers> UserList_Create = new List<tbUsers>();
            foreach (var item in User_create)
            {
                var user = db.tbUsers.FirstOrDefault(p => p.usr_ID == item.Fk_usrsave);
                if (user != null)
                {
                    UserList_Create.Add(user);
                }
            }
            int startRowIndex = 4;
            int endRowIndex = startRowIndex + UserList_Create.Count() - 1;
            foreach (var item in usrID)
            {
                foreach (var item3 in UserList_Create)
                {


                    var x = db.tbCities.Where(p => p.ID == item).FirstOrDefault();
                    //نام شهرستان ها 
                    Row = new Row() { Height = 20, Index = counter };

                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = x.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                {
                    Value = item3.FullName,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                }


                    });
                        int index = 2;

                        foreach (var item2 in MoalefeKarkardi)
                        {
                            var userr = db.tbsaveSoratBasteh.Where(p => p.Fk_usrsave == item3.usr_ID && p.fk_Basteh == Baste && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == month && s.Year == year)).Select(p => p.ID).FirstOrDefault();
                            var moalefevalue = db.tbSoratSavefromExcelMoalfe
                                .Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 &&
                                            p.FK_city == item &&
                                            p.FK_SavedFunctionsID == userr &&
                                            p.tbsaveSoratBasteh.fk_Basteh == Baste)
                                .Select(p => p.Value)
                                .FirstOrDefault();

                            Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = moalefevalue,
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = index
    }
});
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[1].AddRow(Row);
                    }
                }






                var x2 = UserList_Create.Count();
                if (x2 > 1)
                {
                    Moalefeexcelfile.Sheets[1].AddMergedCells($"A{startRowIndex}:A{endRowIndex}");
                    //Moalefeexcelfile.Sheets[1].AddMergedCells($"C{startRowIndex}:C{endRowIndex}");
                    //Moalefeexcelfile.Sheets[1].AddMergedCells($"B{startRowIndex}:B{endRowIndex}");
                    startRowIndex += x2;
                    endRowIndex += x2;
                }






            }


















            return Moalefeexcelfile;

        }


        public Workbook SetDataExcel_Moalefe2(int year, int month, List<int?> usrID, List<int> MoalefeID, DateTime FromDate, DateTime ToDate, int Baste, int PymanID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Moalefe (41).xlsx"));
            Row Row;

            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);

            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);


            //باید از کامنت خارج شود
            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            }
            var count7 = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = db.tbCaranSettings.Where(p=>p.tbContractMoalefeDastmozdi.md_Title==item).Select(s=>s.CaranStandard).FirstOrDefault(),
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count7
                    },
                }); ;
                count7++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            var count722 = 3;
            foreach (var item in MoalefeID)
            {
                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count722
                    },
                }); ;
                count722++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            var count = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }


            //name and lastname and groupjob
            int counter = 3;
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }

            foreach (var item in UserList)
            {
                //نام ، نام خانوادگی و کد پرسنلی
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.usr_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Family,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                    //مقادیر مولفه ها
                    //int index = 3;
                    //foreach (var item2 in MoalefeKarkardi)
                    //{
                    //    var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                    //    && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate && p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                    //    Row.AddCells(new List<Cell>()
                    //    {
                    //         new Cell()
                    //    {
                    //        Value = moalefevalue,
                    //        FontFamily = "B Nazanin",
                    //        Bold = false,
                    //        Enable = true,
                    //        Wrap = false,
                    //        FontSize = 12,
                    //        Italic = false,
                    //        Underline = false,
                    //        Index = index
                    //    }


                    //     });
                    //    index++;
                    //}
                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }

























            //       //fill year
            //       Row = new Row() { Height = 20, Index = 0 };
            //       Row.AddCells(new List<Cell>()
            //       {
            //           new Cell()
            //           {
            //               Value = year,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 1
            //           },
            //       });

            //       Moalefeexcelfile.Sheets[1].AddRow(Row);

            //       //fill month
            //       Row = new Row() { Height = 20, Index = 1 };
            //       Row.AddCells(new List<Cell>()
            //       {
            //           new Cell()
            //           {
            //               Value = month,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 1
            //           },
            //       });

            //       Moalefeexcelfile.Sheets[1].AddRow(Row);


            //       //باید از کامنت خارج شود
            //       //List<string> MoalefeKarkardi = new List<string>();
            //       //foreach (var item in MoalefeID)
            //       //{
            //       //    MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            //       //}

            //       var countu = 4;
            //       foreach (var item in MoalefeKarkardi)
            //       {
            //           Row = new Row() { Height = 20, Index = 2 };
            //           Row.AddCells(new List<Cell>()
            //           {
            //               new Cell()
            //               {
            //                   Value = item,
            //                   FontFamily = "B Nazanin",
            //                   Bold = false,
            //                   Enable = true,
            //                   Wrap = false,
            //                   FontSize = 12,
            //                   Italic = false,
            //                   Underline = false,
            //                   Index = countu
            //               },
            //           });
            //           countu++;
            //           Moalefeexcelfile.Sheets[1].AddRow(Row);
            //       }

            //       //name and lastname and groupjob
            //       int counter4 = 3;
            //       List<tbUsers> UserList4 = new List<tbUsers>();
            //       foreach (var item in usrID)
            //       {
            //           var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //           UserList4.Add(user);
            //       }
            //       var User_create = db.tbSavedFunctions
            //.Where(p => p.svdfunc_BastehID == Baste &&
            //            p.svdfunc_pymnID == PymanID &&
            //            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == year) &&
            //            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Month == month))
            //.ToList();

            //       List<tbUsers> UserList_Create = new List<tbUsers>();
            //       foreach (var item in User_create)
            //       {
            //           var user = db.tbUsers.FirstOrDefault(p => p.usr_ID == item.svdfunc_UserSaveID);
            //           if (user != null)
            //           {
            //               UserList_Create.Add(user);
            //           }
            //       }
            //       int startRowIndex = 4;
            //       int endRowIndex = startRowIndex + UserList_Create.Count() - 1;

            //       foreach (var item in UserList4)
            //       {
            //           foreach (var item3 in UserList_Create)
            //           {
            //               // نام ، نام خانوادگی و کد پرسنلی
            //               Row = new Row() { Height = 20, Index = counter4 };
            //               {
            //                   Row.AddCells(new List<Cell>()
            //       {
            //           new Cell()
            //           {
            //               Value = item.usr_Name,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 0
            //           },
            //           new Cell()
            //           {
            //               Value = item.usr_Family,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 1
            //           },
            //           new Cell()
            //           {
            //               Value = item.usr_Personal_ID,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 2
            //           },
            //           new Cell()
            //           {
            //               Value = item3.FullName,
            //               FontFamily = "B Nazanin",
            //               Bold = false,
            //               Enable = true,
            //               Wrap = false,
            //               FontSize = 12,
            //               Italic = false,
            //               Underline = false,
            //               Index = 3
            //           }
            //       });

            //                   // مقادیر مولفه ها
            //                   int index = 4;

            //                   foreach (var item2 in MoalefeKarkardi)
            //                   {
            //                       var userr = db.tbSavedFunctions.Where(p => p.svdfunc_UserSaveID == item3.usr_ID && p.svdfunc_BastehID == Baste).Select(p => p.svdfunc_ID).FirstOrDefault();
            //                       var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel
            //                           .Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 &&
            //                                       p.MoalfeVal_FKUser == item.usr_ID &&
            //                                       p.FK_SavedFunctionsID == userr &&
            //                                       p.tbSavedFunctions.svdfunc_BastehID == Baste)
            //                           .Select(p => p.MoalfeVal_Value)
            //                           .FirstOrDefault();

            //                       Row.AddCells(new List<Cell>()
            //           {
            //               new Cell()
            //               {
            //                   Value = moalefevalue,
            //                   FontFamily = "B Nazanin",
            //                   Bold = false,
            //                   Enable = true,
            //                   Wrap = false,
            //                   FontSize = 12,
            //                   Italic = false,
            //                   Underline = false,
            //                   Index = index
            //               }
            //           });
            //                       index++;
            //                   }
            //               }

            //               counter4++;
            //               Moalefeexcelfile.Sheets[1].AddRow(Row);
            //           }
            //           var x = UserList_Create.Count();
            //           if (x > 1)
            //           {
            //               Moalefeexcelfile.Sheets[1].AddMergedCells($"A{startRowIndex}:A{endRowIndex}");
            //               Moalefeexcelfile.Sheets[1].AddMergedCells($"C{startRowIndex}:C{endRowIndex}");
            //               Moalefeexcelfile.Sheets[1].AddMergedCells($"B{startRowIndex}:B{endRowIndex}");
            //               startRowIndex += x;
            //               endRowIndex += x;
            //           }
            //       }






            return Moalefeexcelfile;
        }

        //public Workbook SetDataExcel_MoalefeForsorat(int year, int month, List<int?> usrID, List<int> MoalefeID, int Baste, int PymanID)
        //{
        //    var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelforsorat.xlsx"));
        //    Row Row;

        //    Row = new Row() { Height = 20, Index = 0 };
        //    Row.AddCells(new List<Cell>()
        //    {
        //        new Cell()
        //        {
        //            Value = year,
        //            FontFamily = "B Nazanin",
        //            Bold = false,
        //            Enable = true,
        //            Wrap = false,
        //            FontSize = 12,
        //            Italic = false,
        //            Underline = false,
        //            Index = 1
        //        },
        //    });

        //    Moalefeexcelfile.Sheets[0].AddRow(Row);


        //    fill month
        //    Row = new Row() { Height = 20, Index = 1 };
        //    Row.AddCells(new List<Cell>()
        //    {
        //        new Cell()
        //        {
        //            Value = month,
        //            FontFamily = "B Nazanin",
        //            Bold = false,
        //            Enable = true,
        //            Wrap = false,
        //            FontSize = 12,
        //            Italic = false,
        //            Underline = false,
        //            Index = 1
        //        },
        //    });

        //    Moalefeexcelfile.Sheets[0].AddRow(Row);



        //    باید از کامنت خارج شود
        //    List<string> MoalefeKarkardi = new List<string>();
        //    foreach (var item in MoalefeID)
        //    {
        //        MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
        //    }
        //    var counttt = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == PymanID).ToList();

        //    var count = 1;
        //    foreach (var it in counttt)
        //    {
        //        Row = new Row() { Height = 20, Index = 2 };
        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = it.UnitTitle,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = count
        //            },
        //        });
        //        count++;
        //        Moalefeexcelfile.Sheets[0].AddRow(Row);
        //    }
        //    foreach (var item in MoalefeKarkardi)
        //    {
        //        Row = new Row() { Height = 20, Index = 2 };
        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = count
        //            },
        //        });
        //        count++;
        //        Moalefeexcelfile.Sheets[0].AddRow(Row);
        //    }
        //    int counter2 = 3;
        //    foreach (var item in usrID)
        //    {



        //        var x = db.tbCities.Where(p => p.ID == item).FirstOrDefault();
        //        نام شهرستان ها
        //        Row = new Row() { Height = 20, Index = counter2 };


        //        Row.AddCells(new List<Cell>()
        //            {
        //                new Cell()
        //                {
        //                    Value = x.Name,
        //                    FontFamily = "B Nazanin",
        //                    Bold = false,
        //                    Enable = true,
        //                    Wrap = false,
        //                    FontSize = 12,
        //                    Italic = false,
        //                    Underline = false,
        //                    Index = 0
        //                },

        //                     });

        //        counter2++;
        //        Moalefeexcelfile.Sheets[0].AddRow(Row);
        //    }
        //    name and lastname and groupjob




















        //    return Moalefeexcelfile;

        //}


        public Workbook SetDataExcel_Moalefe2edit(int year, int month,int id)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/133443.xlsx"));
            Row Row;
            List<int> usrID2 = new List<int>();

            List<int> Bastehint = new List<int>();
            List<string> Bastehstr = new List<string>();
            var tbmarahlsabtbastedit3 = db.tbmarahlsabtbastedit3.ToList();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int usr = 0;
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                        //return PartialView(User);
                    }
                }
            }

            var tbAdamAdam = db.tbAdamAdam.Where(p=>p.Name==usr&&p.Value==1).ToList();

            

            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id).FirstOrDefault();
            if (tbmarahesabt4 != null)
            {
                var mod11 = tbmarahlsabtbastedit3.Where(s => s.FK_name == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST&&s.tbmarahlsabtbastedit1.typemoalfeh==4 && s.FK_moalfe != null  && s.del != true).ToList();

                var mod = tbmarahlsabtbastedit3.Where(s => s.FK_name == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && s.FK_moalfe != null&&s.del!=true).ToList();
                var mod2 = tbmarahlsabtbastedit3.Where(s => s.FK_name == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && s.FK_taghiz != null && s.del != true).ToList();
              ;
                if (mod11.Count != 0)
                {
                    foreach (var it in mod11)
                    {
                        Bastehstr.Add(it.tbContractMoalefeDastmozdi.md_Title);
                        Bastehint.Add((int)it.FK_moalfe);
                    }

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });
                    Moalefeexcelfile.Sheets[0].AddRow(Row);

                    Row = new Row() { Height = 20, Index = 1 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });
                    var counttt = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn).ToList();
                    var count7 = counttt.Count()+ 3;

                    var count = 3;
                    
                    foreach (var it in counttt)
                    {
                        Row = new Row() { Height = 20, Index = 2 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = it.UnitTitle,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                        count++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }

                    count = 3;
                    foreach (var it in counttt)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = it.ID,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                        count++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                    foreach (var item in Bastehint)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = db.tbCaranSettings.Where(p=>p.tbContractMoalefeDastmozdi.md_ID==item).Select(s=>s.CaranStandard).FirstOrDefault(),
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count7
                    },
                }); ;
                        count7++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    var count722 = counttt.Count() + 3;
                    foreach (var item in Bastehint)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count722
                    },
                }); ;
                        count722++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    var count233 = counttt.Count() + 3;
                    foreach (var item in Bastehstr)
                    {
                        Row = new Row() { Height = 20, Index = 2 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count233
                    },
                });
                        count233++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    int counter = 3;
                    List<tbUsers> UserList = new List<tbUsers>();
                    foreach (var item in tbAdamAdam)
                    {
                        var user = db.tbUsers.Where(p => p.usr_ID == item.nam&& p.usr_Personal_ID != null).FirstOrDefault();
                        if (user != null)
                        {
                            UserList.Add(user);

                        }
                    }
                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = tbmarahesabt4.tbmarahlsabtbastedit2.tbCities.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = tbmarahesabt4.tbmarahlsabtbastedit2.tbCities.ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = tbmarahesabt4.tbmarahlsabtbastedit2.tbCities.ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                        //مقادیر مولفه ها
                        //int index = 3;
                        //foreach (var item2 in MoalefeKarkardi)
                        //{
                        //    var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                        //    && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate && p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                        //    Row.AddCells(new List<Cell>()
                        //    {
                        //         new Cell()
                        //    {
                        //        Value = moalefevalue,
                        //        FontFamily = "B Nazanin",
                        //        Bold = false,
                        //        Enable = true,
                        //        Wrap = false,
                        //        FontSize = 12,
                        //        Italic = false,
                        //        Underline = false,
                        //        Index = index
                        //    }


                        //     });
                        //    index++;
                        //}
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }


                }
                else if (mod.Count != 0)
                {
                    foreach(var it in mod)
                    {
                        Bastehstr.Add(it.tbContractMoalefeDastmozdi.md_Title);
                        Bastehint.Add((int)it.FK_moalfe);
                    }

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

                    Moalefeexcelfile.Sheets[0].AddRow(Row);

                    //fill month
                    Row = new Row() { Height = 20, Index = 1 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

                    Moalefeexcelfile.Sheets[0].AddRow(Row);


                    //باید از کامنت خارج شود
                    //List<string> MoalefeKarkardi = new List<string>();
                    //foreach (var item in MoalefeID)
                    //{
                    //    MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
                    //}
                    var count7 = 3;
                    foreach (var item in Bastehint)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = db.tbCaranSettings.Where(p=>p.tbContractMoalefeDastmozdi.md_ID==item).Select(s=>s.CaranStandard).FirstOrDefault(),
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count7
                    },
                }); ;
                        count7++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    var count722 = 3;
                    foreach (var item in Bastehint)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count722
                    },
                }); ;
                        count722++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                    var count = 3;
                    foreach (var item in Bastehstr)
                    {
                        Row = new Row() { Height = 20, Index = 2 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                        count++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }


                    //name and lastname and groupjob
                    int counter = 3;
                    List<tbUsers> UserList = new List<tbUsers>();
                    foreach (var item in tbAdamAdam)
                    {
                        var user = db.tbUsers.Where(p => p.usr_ID == item.nam && p.usr_Personal_ID != null).FirstOrDefault();
                        if (user != null)
                        {
                            UserList.Add(user);

                        }
                     
                    }

                    foreach (var item in UserList)
                    {
                        //نام ، نام خانوادگی و کد پرسنلی
                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.usr_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Family,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                            //مقادیر مولفه ها
                            //int index = 3;
                            //foreach (var item2 in MoalefeKarkardi)
                            //{
                            //    var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                            //    && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate && p.tbSavedFunctions.svdfunc_BastehID == Baste).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                            //    Row.AddCells(new List<Cell>()
                            //    {
                            //         new Cell()
                            //    {
                            //        Value = moalefevalue,
                            //        FontFamily = "B Nazanin",
                            //        Bold = false,
                            //        Enable = true,
                            //        Wrap = false,
                            //        FontSize = 12,
                            //        Italic = false,
                            //        Underline = false,
                            //        Index = index
                            //    }


                            //     });
                            //    index++;
                            //}
                            counter++;
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }




                    }















                }
                else if (mod2.Count != 0)
                {
                    Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
                    var count3 = 2;
                    var count = 2;

                    foreach (var it in mod2)
                    {
                        var equipments = db.tbEquipments.Where(p => p.FK_Peyman == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.tbPeymanContracts.Inactive != true&&p.tbEquipmentBunch.Eqpbnch_ID==it.tbEquipmentBunch.Eqpbnch_ID).ToList();

                        Bastehstr.Add(it.tbEquipmentBunch.Eqpbnch_Name);
                        Bastehint.Add((int)it.FK_taghiz);
                        if (equipments.Count() > 0)
                        {
                            int counter = 1;
                            var counttt = 2;

                            var count1 = 4;
                            var count2 = 1;

                          
                            foreach (var item in equipments)
                            {
                                string allinfo = "";
                                foreach (var item2 in item.tbEquipmentSpecificationData)
                                {
                                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                                }
                                var c = "";

                                if (item.Users != null)
                                {
                                    var model2 = item.Users.Split(',');
                                    foreach (var item3 in model2)
                                    {
                                        var id345 = System.Convert.ToInt32(item3);
                                        var temp = db.tbUsers.Where(p => p.usr_ID == id345).FirstOrDefault();
                                        c += temp.FullName + " , ";

                                    }
                                }
                                var t = "";
                                if (item.personal != null)
                                {
                                    if (item.personal == true)
                                    {
                                        t = "پرسنلی";
                                    }
                                    else
                                    {
                                        t = "کارگزاری";

                                    }
                                }
                                var t1 = "";
                                if (item.type != null)
                                {
                                    if (item.type == 1)
                                    {
                                        t1 = "تیرسان ";
                                    }
                                    else if (item.type == 3)
                                    {
                                        t1 = "بدون یونیفرم";

                                    }
                                    else if (item.type == 2)
                                    {
                                        t1 = "یونیفرم ساده";

                                    }
                                }




                                Row = new Row() { Height = 20, Index = 0 };
                                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                                count++;
                                Moalefeexcelfile.Sheets[0].AddRow(Row);









                            }

                            foreach (var item in equipments)
                            {
                                string allinfo = "";
                                foreach (var item2 in item.tbEquipmentSpecificationData)
                                {
                                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                                }
                                var c = "";

                                if (item.Users != null)
                                {
                                    var model2 = item.Users.Split(',');
                                    foreach (var item3 in model2)
                                    {
                                        var id5 = System.Convert.ToInt32(item3);
                                        var temp = db.tbUsers.Where(p => p.usr_ID == id5).FirstOrDefault();
                                        c += temp.FullName + " , ";

                                    }
                                }
                                var t = "";
                                if (item.personal != null)
                                {
                                    if (item.personal == true)
                                    {
                                        t = "پرسنلی";
                                    }
                                    else
                                    {
                                        t = "کارگزاری";

                                    }
                                }
                                var t1 = "";
                                if (item.type != null)
                                {
                                    if (item.type == 1)
                                    {
                                        t1 = "تیرسان ";
                                    }
                                    else if (item.type == 3)
                                    {
                                        t1 = "بدون یونیفرم";

                                    }
                                    else if (item.type == 2)
                                    {
                                        t1 = "یونیفرم ساده";

                                    }
                                }

                              
                                    Row = new Row() { Height = 20, Index = 1 };
                                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name+item.tbEquipmentBunch.Eqpbnch_Name +"-----"+"  نوع ماشین:  "+  t1,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                                
                         



                                count3++;
                                Moalefeexcelfile.Sheets[0].AddRow(Row);









                            }
                            var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true).ToList();
                            var pymn2 = db.tbPeymanContracts.Where(p => p.pec_ID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.Inactive != true).FirstOrDefault();
                            var company = db.tbCompanies.Where(p => p.ID == pymn2.FK_KarfarmaID).FirstOrDefault();

                            Row = new Row() { Height = 20, Index = 2 };
                            Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = company.CompanyName,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 0
                    },
                         new Cell()
                    {
                        Value = company.Company_National_ID,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 1
                    },

            });
                            Moalefeexcelfile.Sheets[0].AddRow(Row);

                            Row = new Row() { Height = 20, Index = 3 };
                            Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = pymn2.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 0
                    },
                         new Cell()
                    {
                        Value = pymn2.pec_ProjectCode,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 1
                    },

            });
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                            List<tbUsers> UserList = new List<tbUsers>();
                            foreach (var item in tbAdamAdam)
                            {
                                var user222 = db.tbUsers.Where(p => p.usr_ID == item.nam && p.usr_Personal_ID != null).FirstOrDefault();
                                if (user222 != null)
                                {
                                    UserList.Add(user222);

                                }
                           
                            }
                            foreach (var item in UserList)
                            {
                                // ایجاد ردیف جدید
                                Row = new Row() { Height = 20, Index = count1 };

                                // اضافه کردن سلول‌های مربوط به کاربر
                                Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = item.FullName,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },
        new Cell()
        {
            Value = item.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
                                });



                                count1++; // افزایش ایندکس ردیف‌ها

                                // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                                Moalefeexcelfile.Sheets[0].AddRow(Row);
                            }


                           

                            var user = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
                            var pymn = db.tbPeymanContracts.Where(p => p.Inactive == null).ToList();
                            foreach (var item in UserList)
                            {
                                Row = new Row() { Height = 20, Index = count1 };
                                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.FullName,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 0
                    },
                         new Cell()
                    {
                        Value = item.usr_Personal_ID,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                             Index = 1
                    },
                });
                                count1++;
                                Moalefeexcelfile.Sheets[1].AddRow(Row);
                            }

                            foreach (var item in pymn)
                            {
                                Row = new Row() { Height = 20, Index = count2 };
                                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 2
                    },
                         new Cell()
                    {
                        Value = item.pec_ProjectCode,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                             Index = 3
                    },
                });
                                count2++;
                                Moalefeexcelfile.Sheets[1].AddRow(Row);
                            }
                        }
                        }
                }
            }
         








            return Moalefeexcelfile;
        }



        public ActionResult viewsabttaeedbyfilter(int year, int month, int id)
        {
            var model = new FunctionModel
            {
                tbAdamAdam2 = new List<tbAdamAdam2>(),
                tbAdamAdam3 = new List<tbAdamAdam3>(),
                sabttaeedviewsoton = new List<sabttaeedviewsoton>(),
                sabttaeedviewAdama = new List<sabttaeedviewAdama>(),

                tbAdamAdam = new List<tbAdamAdam>(),
                arzyzbi5_time = new List<arzyzbi5_time>(),
                arzyzbi5_time2 = new List<arzyzbi5_time2>()

            };
            var tbmarahlsabtbastedit3 = db.tbmarahlsabtbastedit3.ToList();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int usr = 0;
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                        //return PartialView(User);
                    }
                }
            }

            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id).FirstOrDefault();
            if (tbmarahesabt4 != null)
            {
                var mod2 = tbmarahlsabtbastedit3.Where(s => s.FK_name == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && s.FK_taghiz != null && s.del != true).ToList();
                if (mod2.Count != 0)
                {
                    foreach(var it in mod2)
                    {
                        var equipments = db.tbEquipments.Where(p => p.FK_Peyman == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.Eqpbnch_ID == it.tbEquipmentBunch.Eqpbnch_ID).ToList();
                        foreach(var it1 in equipments)
                        {
                            model.sabttaeedviewsoton.Add(new sabttaeedviewsoton
                            {
                                fullname = it1.tbEquipmentBunch.Eqpbnch_Name,
                                ID = it1.ID,
                               
                                // مقدار z به عنوان ID
                            });
                        }
                    }
                }
            }
            foreach(var it in tbAdamAdam)
            {
                var find = db.tbUsers.Where(s => s.usr_ID == it.nam &&s.usr_City_Dutysystem== tbmarahesabt4.tbmarahlsabtbastedit2.FK_City).FirstOrDefault();
                if (find != null && find.usr_Personal_ID != null)
                {
                    model.sabttaeedviewAdama.Add(new sabttaeedviewAdama
                    {
                        fullname = find.FullName,
                        ID = (int)it.nam,
                        personalID = (int)find.usr_Personal_ID,
                        month = month,
                        year = year,
                        ID4 = id

                        // مقدار z به عنوان ID
                    });
                }
            
            }
            model.savedValues = db.tbEquipmentMoalefeValue
    .Where(v => v.tbEquipmentMoalefeValueReffrenceSave.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn&&v.tbEquipmentMoalefeValueReffrenceSave.del!=true&& v.tbEquipmentMoalefeValueReffrenceSave.FK_tbmarahesabt4==id&& v.tbEquipmentMoalefeValueReffrenceSave.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City&& v.tbEquipmentMoalefeValueReffrenceSave.FK_Basteh== tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.ID&&v.tbEquipmentMoalefeValueReffrenceSave.FK_User==usr&&v.CountDays==1&& v.Month == month && v.Year == year&&v.CountDays==1)
    .Select(v => new EquipmentSavedValueDto
    {
        FK_Equipment = (int)v.FK_Equipment,
        Fk_user = (int)v.Fk_user,
        CountDays = (double)v.CountDays
    })
    .ToList();
            return PartialView("~/Areas/Salaries/Views/Moalefe/viewsabttaeedbyfilter.cshtml", model);
        }
        public ActionResult viewsabttaeedbyfilterdetail(int year, int month, int id,int person)
        {
            var model = new FunctionModel
            {
                tbAdamAdam2 = new List<tbAdamAdam2>(),
                tbAdamAdam3 = new List<tbAdamAdam3>(),
                sabttaeedviewsoton = new List<sabttaeedviewsoton>(),
                sabttaeedviewAdama = new List<sabttaeedviewAdama>(),

                tbAdamAdam = new List<tbAdamAdam>(),
                arzyzbi5_time = new List<arzyzbi5_time>(),
                arzyzbi5_time2 = new List<arzyzbi5_time2>()

            };
            var tbmarahlsabtbastedit3 = db.tbmarahlsabtbastedit3.ToList();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int usr = 0;
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                        //return PartialView(User);
                    }
                }
            }
            usr = person;
            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id).FirstOrDefault();
            if (tbmarahesabt4 != null)
            {
                var mod2 = tbmarahlsabtbastedit3.Where(s => s.FK_name == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && s.FK_taghiz != null && s.del != true).ToList();
                if (mod2.Count != 0)
                {
                    foreach (var it in mod2)
                    {
                        var equipments = db.tbEquipments.Where(p => p.FK_Peyman == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.Eqpbnch_ID == it.tbEquipmentBunch.Eqpbnch_ID).ToList();
                        foreach (var it1 in equipments)
                        {
                            model.sabttaeedviewsoton.Add(new sabttaeedviewsoton
                            {
                                fullname = it1.tbEquipmentBunch.Eqpbnch_Name,
                                ID = it1.ID,

                                // مقدار z به عنوان ID
                            });
                        }
                    }
                }
            }
            foreach (var it in tbAdamAdam)
            {
                var find = db.tbUsers.Where(s => s.usr_ID == it.nam && s.usr_City_Dutysystem == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City).FirstOrDefault();
                if (find != null && find.usr_Personal_ID != null)
                {
                    model.sabttaeedviewAdama.Add(new sabttaeedviewAdama
                    {
                        fullname = find.FullName,
                        ID = (int)it.nam,
                        personalID = (int)find.usr_Personal_ID,
                        month = month,
                        year = year,
                        ID4 = id

                        // مقدار z به عنوان ID
                    });
                }

            }
            model.savedValues = db.tbEquipmentMoalefeValue
    .Where(v => v.tbEquipmentMoalefeValueReffrenceSave.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && v.tbEquipmentMoalefeValueReffrenceSave.del != true && v.tbEquipmentMoalefeValueReffrenceSave.FK_tbmarahesabt4 == id && v.tbEquipmentMoalefeValueReffrenceSave.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && v.tbEquipmentMoalefeValueReffrenceSave.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.ID && v.tbEquipmentMoalefeValueReffrenceSave.FK_User == usr && v.CountDays == 1 && v.Month == month && v.Year == year && v.CountDays == 1)
    .Select(v => new EquipmentSavedValueDto
    {
        FK_Equipment = (int)v.FK_Equipment,
        Fk_user = (int)v.Fk_user,
        CountDays = (double)v.CountDays
    })
    .ToList();
            return PartialView("~/Areas/Salaries/Views/Moalefe/viewsabttaeedbyfilterdetail.cshtml", model);
        }

        [HttpPost]
        public async Task< ActionResult> Linvaluerefrence(List<EquipmentValueDto> dataList, int id, int month, int year)
        {
            try
            {
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                int usr = 0;
                if (cookie_user != null)
                {
                    var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);

                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                        //return PartialView(User);
                    }

                }
                var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id).FirstOrDefault();
                tbEquipmentMoalefeValueReffrenceSave tbEquipmentMoalefeValueReffrenceSave = new tbEquipmentMoalefeValueReffrenceSave();

                if (tbmarahesabt4 != null)
                {
                    var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
                    var sabt = tbmarhesabt3.marahelsabt;
                    var taeed = tbmarhesabt3.taeed;

                    int vaziat = 0;
                    int vaziat2 = 0;

                    if (tbmarahesabt4.sabt == true)
                    {
                        vaziat = 1;
                        vaziat2 = (int)tbmarahesabt4.number;

                    }
                    else if (tbmarahesabt4.taeed == true)
                    {
                        vaziat = 2;
                        vaziat2 = (int)tbmarahesabt4.number;
                    }
                    var find = db.tbEquipmentMoalefeValueReffrenceSave.Where(s => s.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn&&s.FK_User== usr && s.del != true && s.FK_tbmarahesabt4==id&& s.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && s.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.ID&&s.tbEquipmentMoalefeValue.Any(t=>t.Month==month&&t.Year==year)).FirstOrDefault();
                    if (find != null)
                    {
                        find.Datatmie= DateTime.Now;
                        tbEquipmentMoalefeValueReffrenceSave = find;
                        await db.SaveChangesAsync();
                      
                            string sql = "DELETE FROM [Contract].[tbEquipmentMoalefeValue] WHERE FK_tbEquipmentMoalefeValueReffrenceSave = @id";
                            var param = new SqlParameter("@id", find.ID);

                            await db.Database.ExecuteSqlCommandAsync(sql, param);
                        
                    }
                    else
                    {
                        if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                        {
                            tbEquipmentMoalefeValueReffrenceSave.Final_Accept = true;
                            tbEquipmentMoalefeValueReffrenceSave.For_Acceot = true;

                        }
                        else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                        {
                            tbEquipmentMoalefeValueReffrenceSave.Final_Accept = false;
                            tbEquipmentMoalefeValueReffrenceSave.For_Acceot = true;

                        }
                        else
                        {
                            tbEquipmentMoalefeValueReffrenceSave.Final_Accept = false;
                            tbEquipmentMoalefeValueReffrenceSave.For_Acceot = false;
                        }
                        tbEquipmentMoalefeValueReffrenceSave.FK_tbmarahesabt4 = id;
                        tbEquipmentMoalefeValueReffrenceSave.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;
                        tbEquipmentMoalefeValueReffrenceSave.FK_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                        tbEquipmentMoalefeValueReffrenceSave.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.ID;
                        tbEquipmentMoalefeValueReffrenceSave.FK_User = usr;
                        tbEquipmentMoalefeValueReffrenceSave.DateTime = DateTime.Now;
                        tbEquipmentMoalefeValueReffrenceSave.del = false;
                        tbEquipmentMoalefeValueReffrenceSave.FK_Baste = null;


                        db.tbEquipmentMoalefeValueReffrenceSave.Add(tbEquipmentMoalefeValueReffrenceSave);
                        await db.SaveChangesAsync();
                    }
                }
                // مرحله 1: ذخیره در tbEquipmentMoalefeValueReffrenceSave

                // تا اینجا ID ساخته بشه

                
                

                int newRefId = tbEquipmentMoalefeValueReffrenceSave.ID;
                List<tbEquipmentMoalefeValue> tbEquipmentMoalefeValue = new List<tbEquipmentMoalefeValue>();
                // مرحله 2: ذخیره در tbEquipmentMoalefeValue
                foreach (var item in dataList)
                {
                    var newValue = new tbEquipmentMoalefeValue
                    {
                        FK_tbEquipmentMoalefeValueReffrenceSave = newRefId,
                        FK_Equipment = item.FK_arzyzbi1,
                        Fk_user = item.FK_usr,
                        Month = month,
                        Year = year,
                        CountDays = item.vaziat ? 1 : 0,
                        Status = null, // یا مقدار پیش‌فرض
                        Description = null, // اگر نیاز داری
                        accepted = false,
                        Final_accept = false
                    };
                    tbEquipmentMoalefeValue.Add(newValue);
                }
                db.tbEquipmentMoalefeValue.AddRange(tbEquipmentMoalefeValue);

                await db.SaveChangesAsync();


                return Json("true", JsonRequestBehavior.AllowGet);

            }
            
              catch( Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);

            }
        }
        public class EquipmentValueDto
        {
            public int FK_arzyzbi1 { get; set; } // marahl.ID
            public int FK_usr { get; set; }      // contract.ID
            public bool vaziat { get; set; }     // checked
        }

        public ActionResult show_Code()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == CodeMeli).FirstOrDefault();
                    if (User != null)
                    {
                        return PartialView(User);
                    }
                }
            }

            return PartialView();
        }


        public string ImportExcel_Moalefe222(HttpPostedFileBase file, int Month, int Year, int ID)
        {

            if (file == null)
            {
                return "فایل بطور صحیح بارگذاری نشده است";
            }



            return GetDataFromExcel_Moalefe_Foraccept2222(file, Month, Year, ID);

        }
        public string GetDataFromExcel_Moalefe_Foraccept2222(HttpPostedFileBase files, int Month2, int Year2, int ID)
        {
            string name, family, personalid, year, month = "";
            string value;


            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            //var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID)).Select(p => p.FK_User_ID).ToList();
            var usrID = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == ID && p.Status == true).ToList();
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item.FK_Peyman_ID).FirstOrDefault();
                UserList.Add(user);
            }

            int numberOfUsers = UserList.Count + 3;


            List<string> lstMoalefe = new List<string>();
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {




                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }

                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString());
                                }

                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);


                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault().usr_ID;
                                    int shomarande = 3;



                                    foreach (var item in lstMoalefe)
                                    {
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                //var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID;


                                                value = Value.Value.ToString();
                                                personalabzar dastmozdexcel = new personalabzar
                                                {
                                                    FK_PYM = ID,
                                                    FK_usr = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    value = System.Convert.ToDouble(value),
                                                    FK_Equgroup = db.tbEquipmentBunch.Where(p => p.Eqpbnch_Name == item).Select(s => s.Eqpbnch_ID).FirstOrDefault(),

                                                };
                                                db.personalabzar.Add(dastmozdexcel);
                                                db.SaveChanges();
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}


                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                personalabzar dastmozdexcel = new personalabzar
                                                {
                                                    FK_PYM = ID,
                                                    FK_usr = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    value = 0,
                                                    FK_Equgroup = db.tbEquipmentBunch.Where(p => p.Eqpbnch_Name == item).Select(s => s.Eqpbnch_ID).FirstOrDefault(),

                                                };
                                                db.personalabzar.Add(dastmozdexcel);
                                                db.SaveChanges();
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            personalabzar dastmozdexcel = new personalabzar
                                            {
                                                FK_PYM = ID,
                                                FK_usr = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                value = 0,
                                                FK_Equgroup = db.tbEquipmentBunch.Where(p => p.Eqpbnch_Name == item).Select(s => s.Eqpbnch_ID).FirstOrDefault(),

                                            };
                                            db.personalabzar.Add(dastmozdexcel);
                                            db.SaveChanges();

                                        }

                                        shomarande++;
                                    }



                                }


                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }






        [AuthorizeAAA]
        [HttpPost]




        public async Task<ActionResult> ImportExcel_Moalefe(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate, int Basteh, int PeymanID)
        {

            if (MyExcelStream == null)
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");
            }



            return  await GetDataFromExcel_Moalefe(MyExcelStream, FromDate, ToDate, Basteh, PeymanID);

        }







        public async Task<ActionResult> ImportExcel_Moalefesistan(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate, int Basteh, int PeymanID)
        {

            if (MyExcelStream == null)
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");
            }



            return await GetDataFromExcel_Moalefesistan(MyExcelStream, FromDate, ToDate, Basteh, PeymanID);

        }

        public ActionResult ImportExcel_Moalefe_abzar(HttpPostedFileBase MyExcelStream, int Basteh, int PeymanID)
        {

            if (MyExcelStream == null)
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");
            }

            return GETEXCEL(MyExcelStream, Basteh, PeymanID);
            //return GetDataFromExcel_Moalefe(MyExcelStream, FromDate, ToDate, Basteh, PeymanID);

        }



        public ActionResult ImportExcel_Moalefe_abzar3(HttpPostedFileBase MyExcelStream, int Basteh, int PeymanID)
        {

            if (MyExcelStream == null)
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");
            }

            return GETEXCEL2(MyExcelStream, Basteh, PeymanID);
            //return GetDataFromExcel_Moalefe(MyExcelStream, FromDate, ToDate, Basteh, PeymanID);

        }


        public ActionResult GETEXCEL2(HttpPostedFileBase MyExcelStream, int Basteh, int PeymanID)
        {
            List<string> lstMoalefe = new List<string>();
            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;
            List<tbEquipments> obj2 = new List<tbEquipments>();
            int ff = 0;
            float uni = 0;

            ;

            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();


                            if (count == 1)
                            {
                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                            }

                            for (int i = 13; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }




                            for (int i = 1; i < count; i++)
                            {
                                var row = workbook.Worksheets[0].Rows[i];
                                var Name144 = row.Cells[0];
                                if (Name144.Value!=null)
                                {
                                    name = Name144.Value.ToString();
                                }
                                else
                                {
                                    name = null;

                                    //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                }
                                var Name = row.Cells[1];
                                if (float.TryParse(Name.Value.ToString(), out float vahed2))
                                {
                                    vahed = vahed2;
                                }
                                else
                                {
                                    vahed = null;

                                    //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                }
                                var Name2 = row.Cells[2];
                                if (float.TryParse(Name2.Value.ToString(), out float vahed3))
                                {
                                    countvahed = vahed3;
                                }
                                else
                                {
                                    //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                    countvahed = null;
                                }
                                var Name3 = row.Cells[3];
                                if (float.TryParse(Name3.Value.ToString(), out float vahed4))
                                {
                                    numbervahed = vahed4;
                                }
                                else
                                {
                                    numbervahed = null;
                                }
                                var Name4 = row.Cells[4];
                                if (float.TryParse(Name4.Value.ToString(), out float vahed5))
                                {
                                    colmablagh = vahed5;
                                }
                                else
                                {
                                    colmablagh = null;
                                }
                                var Name5 = row.Cells[5];
                                if (float.TryParse(Name5.Value.ToString(), out float vahed6))
                                {
                                    numbermondareg = vahed6;
                                }
                                else
                                {
                                    numbermondareg = null;
                                }
                                var Name6 = row.Cells[6];
                                if (float.TryParse(Name6.Value.ToString(), out float vahed7))
                                {
                                    typetazmin = vahed7;
                                }
                                else
                                {
                                    typetazmin = null;
                                }
                                var Name7 = row.Cells[7];
                                if (float.TryParse(Name7.Value.ToString(), out float vahed8))
                                {
                                    typemalk = vahed8;
                                }
                                else
                                {
                                    typemalk = null;
                                }
                                var Name8 = row.Cells[8];
                                if (float.TryParse(Name8.Value.ToString(), out float vahed9))
                                {
                                    numbertazmin = vahed9;
                                }
                                else
                                {
                                    numbertazmin = null;
                                }
                                var Name9 = row.Cells[10];
                                if (float.TryParse(Name9.Value.ToString(), out float vahed10))
                                {
                                    typecar = vahed10;
                                }
                                else
                                {
                                    typecar = null;
                                }
                                var Name10 = row.Cells[11];
                                if (float.TryParse(Name10.Value.ToString(), out float vahed11))
                                {
                                    codepersenly = vahed11;
                                }
                                else
                                {
                                    codepersenly = null;
                                }
                                var Name11 = row.Cells[12];
                                if (float.TryParse(Name11.Value.ToString(), out float vahed12))
                                {
                                    egareh = vahed12;
                                }
                                else
                                {
                                    egareh = null;
                                }
                                var Name12 = row.Cells[13];
                                if (float.TryParse(Name12.Value.ToString(), out float vahed13))
                                {
                                    soght = vahed13;
                                }
                                else
                                {
                                    soght = null;
                                }
                                var Name13 = row.Cells[14];
                                if (float.TryParse(Name13.Value.ToString(), out float vahed14))
                                {
                                    uni = vahed14;
                                }
                                else
                                {
                                    uni = 0;
                                }

                                var Name14 = row.Cells[15];
                                if (int.TryParse(Name14.Value.ToString(), out int vahed15))
                                {
                                    ff = vahed15;
                                }
                                else
                                {
                                    ff = 0;
                                }

                                var f = db.tbEquipmentBunch.Where(p=>p.Eqpbnch_Name==name).FirstOrDefault();

                                tbEquipments obj = new tbEquipments();
                                obj.FK_Bunch = f.Eqpbnch_ID;
                                obj.FK_Peyman = PeymanID;
                                obj.Unit = (int?)vahed;
                                obj.TotalPrice = (int?)colmablagh;
                                obj.EachValue = (int?)numbervahed;
                                obj.TypeTazmin = (int?)typetazmin;
                                obj.UnitCount = (int?)countvahed;
                                obj.Count_InQardad = (int?)numbermondareg;
                                obj.TypeTazmin = (int?)typetazmin;
                                obj.uniform = uni;
                                obj.model = ff;
                                if (typemalk == 1)
                                {
                                    obj.personal = true;

                                }
                                else
                                {
                                    obj.personal = false;

                                }

                                obj.TazminPrice = (int?)numbertazmin;
                                obj.type = (int?)typecar;

                                obj.Fk_usr = db.tbUsers.Where(p => p.usr_Personal_ID == codepersenly).Select(p => p.usr_ID).FirstOrDefault();
                                obj.Soght = (float?)soght;
                                obj.egareh = (float?)egareh;
                                obj2.Add(obj);
                             
                                //int savedItemId = db.tbEquipments
                                //                    .OrderByDescending(p => p.ID)
                                //                    .Select(s => s.ID)
                                //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                           

                            }
                            db.tbEquipments.AddRange(obj2);
                            db.SaveChanges();

                        }
                    }
                }
                catch (Exception ex)
                {
                    return Content(ex.Message);
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return Content("True");

            //return Content("False");
            //var Model33 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

        }



        public ActionResult GETEXCEL(HttpPostedFileBase MyExcelStream, int Basteh, int PeymanID)
        {
            List<string> lstMoalefe = new List<string>();
            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;


            ;
           
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {

                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            //var count = workbook.Sheets[0].Rows.Count();

                            if (Rows.Length >= 1)
                            {
                                var title = workbook.Worksheets[0].Rows[0].Cells;
                                var count = workbook.Worksheets[0].Rows.Count();


                                if (count == 1)
                                {
                                    return Content("فایل اکسل فاقد اطلاعات می باشد");
                                }

                                for (int i = 13; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString());
                                }




                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (float.TryParse(Name.Value.ToString(), out float vahed2))
                                    {
                                        vahed = vahed2;
                                    }
                                    else
                                    {
                                    vahed = null;

                                    //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                }
                                var Name2 = row.Cells[1];
                                    if (float.TryParse(Name2.Value.ToString(), out float vahed3))
                                    {
                                        countvahed = vahed3;
                                    }
                                    else
                                    {
                                    //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                    countvahed = null;
                                    }
                                    var Name3= row.Cells[2];
                                    if (float.TryParse(Name3.Value.ToString(), out float vahed4))
                                    {
                                        numbervahed = vahed4;
                                    }
                                    else
                                    {
                                    numbervahed = null; 
                                        }
                                    var Name4 = row.Cells[3];
                                    if (float.TryParse(Name4.Value.ToString(), out float vahed5))
                                    {
                                        colmablagh = vahed5;
                                    }
                                    else
                                    {
                                    colmablagh = null;
                                }
                                    var Name5 = row.Cells[4];
                                    if (float.TryParse(Name5.Value.ToString(), out float vahed6))
                                    {
                                        numbermondareg = vahed6;
                                    }
                                    else
                                    {
                                    numbermondareg = null;
                                    }
                                    var Name6 = row.Cells[5];
                                    if (float.TryParse(Name6.Value.ToString(), out float vahed7))
                                    {
                                        typetazmin = vahed7;
                                    }
                                    else
                                    {
                                    typetazmin = null;
                                }
                                    var Name7 = row.Cells[6];
                                    if (float.TryParse(Name7.Value.ToString(), out float vahed8))
                                    {
                                        typemalk = vahed8;
                                    }
                                    else
                                    {
                                    typemalk = null;                                    }
                                    var Name8 = row.Cells[7];
                                    if (float.TryParse(Name8.Value.ToString(), out float vahed9))
                                    {
                                        numbertazmin = vahed9;
                                    }
                                    else
                                    {
                                    numbertazmin = null;
                                }
                                    var Name9 = row.Cells[9];
                                    if (float.TryParse(Name9.Value.ToString(), out float vahed10))
                                    {
                                        typecar = vahed10;
                                    }
                                    else
                                    {
                                    typecar = null;
                                }
                                    var Name10 = row.Cells[10];
                                    if (float.TryParse(Name10.Value.ToString(), out float vahed11))
                                    {
                                        codepersenly = vahed11;
                                    }
                                    else
                                    {
                                    codepersenly = null;
                                }
                                    var Name11 = row.Cells[11];
                                    if (float.TryParse(Name11.Value.ToString(), out float vahed12))
                                    {
                                        egareh = vahed12;
                                    }
                                    else
                                    {
                                    egareh = null;
                                }
                                    var Name12= row.Cells[12];
                                    if (float.TryParse(Name12.Value.ToString(), out float vahed13))
                                    {
                                        soght = vahed13;
                                    }
                                    else
                                    {
                                    soght = null;
                                }
                                    tbEquipments obj = new tbEquipments();
                                    obj.FK_Bunch = Basteh;
                                    obj.FK_Peyman = PeymanID;
                                    obj.Unit = (int?)vahed;
                                    obj.TotalPrice = (int?)colmablagh;
                                    obj.EachValue = (int?)numbervahed;
                                    obj.TypeTazmin = (int?)typetazmin;
                                    obj.UnitCount = (int?)countvahed;
                                    obj.Count_InQardad = (int?)numbermondareg;
                                    obj.TypeTazmin = (int?)typetazmin;
                                    if (typemalk == 1)
                                    {
                                        obj.personal = true;

                                    }
                                    else
                                    {
                                        obj.personal = false;

                                    }

                                obj.TazminPrice = (int?)numbertazmin;
                                obj.type = (int?)typecar;

                                    obj.Fk_usr = db.tbUsers.Where(p => p.usr_Personal_ID == codepersenly).Select(p => p.usr_ID).FirstOrDefault();
                                    obj.Soght = (float?)soght;
                                    obj.egareh = (float?)egareh;
                                    db.tbEquipments.Add(obj);
                                    db.SaveChanges();
                                    //int savedItemId = db.tbEquipments
                                    //                    .OrderByDescending(p => p.ID)
                                    //                    .Select(s => s.ID)
                                    //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                    // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                                    int shomarande = 13;
                                    //var Model = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();
                                    foreach (var it in lstMoalefe)
                                    {
                                        var tt = db.tbEquipmentSpecifications.Where(p => p.sp_SpecTitle == it && p.FK_EqpBunch_ID == Basteh).FirstOrDefault();
                                        var Value = row.Cells[shomarande];
                                       var value = Value.Value.ToString()?? null;

                                        tbEquipmentSpecificationData obj2 = new tbEquipmentSpecificationData();
                                        obj2.SpcData_ID = tt.sp_ID;
                                        obj2.FK_Equipment = obj.ID;
                                        obj2.SpcData = value;
                                        db.tbEquipmentSpecificationData.Add(obj2);
                                        db.SaveChanges();
                                        shomarande++;





                                    }

                                }

                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        return Content(ex.Message);
                    }
                }
                //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

                return Content("True");
            
            //return Content("False");
            //var Model33 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

        }


        [AuthorizeAAA]

        // در اینجا تابع GetDataFromExcel_Moalefe را قرار دهید
        public async Task<ActionResult> GetDataFromExcel_Moalefe(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate, int Basteh, int PeymanID)
        {
            string name, family, personalid, year, month = "";
            string value;
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToList();
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);

            }

            int numberOfUsers = UserList.Count + 3;

            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();

            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh && s.Deleted != true).FirstOrDefault();

            var us = db.tbReffrenceSave
                .Where(p => p.FK_PeymanID == number)
                .SelectMany(s => s.tbReffrenceSaveLevel)
                .Where(s => s.ID == Basteh)
                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
                .Where(p => p.FK_RRSave == us.FK_RRSave)

                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var rus3 = db.tbReffrenceSaveLevelUser
                .Where(p => p.FK_LevelID == us2.ID)
                .OrderByDescending(p => p.ID)
                .Select(p => p.FK_UserID)
                .FirstOrDefault();

            List<string> lstMoalefe = new List<string>();
            List<int> lstMoalefeint = new List<int>();

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
                    }
                }
            }

            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.UsersMoalefe = new List<MoalefeUserInfo>();
            List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
            lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {

                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;

                            if (Rows.Length >= 1)
                            {
                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var count = workbook.Worksheets[0].Rows.Count();
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;


                                if (count == 1)
                                {
                                    return Content("فایل اکسل فاقد اطلاعات می باشد");
                                }

                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    //IRange cell = title[i];
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }
                                }
                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return Content(" لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ");
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return Content(" لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ");
                                }

                                int intYear = System.Convert.ToInt32(year);
                                int intMonth = System.Convert.ToInt32(month);

                                var find = db.tbSavedFunctions
                                    .Where(p => p.svdfunc_BastehID == Basteh &&
                                                p.svdfunc_pymnID == number &&
                                                p.Final_Accept == true &&
                                                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == intYear && m.MoalfeVal_Month == intMonth))
                                    .FirstOrDefault();

                                if (find != null)
                                {
                                    return Content("شما در این بسته ثبت و تایید شده است ");
                                }
                                var sabtt = db.tbSavedFunctions
                                    .Where(p => p.svdfunc_BastehID == Basteh &&
                                                p.svdfunc_pymnID == number &&
                                                p.svdfunc_UserSaveID == userid &&

                                                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == intYear && m.MoalfeVal_Month == intMonth))
                                    .FirstOrDefault();
                                if (sabtt != null)
                                {
                                    return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
                                }

                                lstMoalefeexcel2.SabtHeader = new Header
                                {
                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                    PeymanName = PymN.pec_Title,
                                    BastehName = BastehName2.Title,
                                    FromDate = FromDate,
                                    ToDate = ToDate,


                                };
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault().usr_ID;
                                    var UserID2 = await db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefaultAsync();
                                    int shomarande = 3;
                                    //FunctionModel functionModel = new FunctionModel
                                    //{
                                    //    SabtHeader = new Header
                                    //    {
                                    //        MoalfeVal_Month = Convert.ToInt32(month),
                                    //        MoalfeVal_Year = Convert.ToInt32(year),
                                    //        PeymanName = PymN.pec_Title,
                                    //        BastehName = BastehName2.Title
                                    //    },
                                    //    UsersMoalefe = usersMoalefeList
                                    //};

                                    var obj = new MoalefeUserInfo();
                                    obj.FullName = UserID2.FullName;
                                    obj.PersonalCode = System.Convert.ToInt32(UserID2.usr_Personal_ID);
                                    obj.listMoalefe = new List<MoalefeInfo>();
                                    foreach (var item in lstMoalefeint)
                                    {
                                        var MoalefeID =await db.tbContractMoalefeDastmozdi.Where(p => p.md_ID==item).FirstOrDefaultAsync();

                                        var Value = row.Cells[shomarande];
                                        //if (Value.Value != null && Value.Value.ToString() != "")
                                        //{
                                        if (!string.IsNullOrWhiteSpace(MoalefeID.md_Title))
                                        {
                                            //var MoalefeID =await db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(item)).FirstOrDefaultAsync(s=>s.md_ID);
                                            //var MoalefeID = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(s => s.md_ID).FirstOrDefault();


                                            value = Value.Value.ToString() ?? "0";

                                            // MoalefeInfo dastmozdexce4 =
                                            obj.listMoalefe.Add(new MoalefeInfo
                                            {
                                                MoalefeTitle = MoalefeID.md_Title,
                                                MoalefeValue = value
                                            });

                                            //List<MoalefeUserInfo> usersMoalefeList = new List<MoalefeUserInfo>();
                                            //MoalefeUserInfo userInfo = new MoalefeUserInfo
                                            //{
                                            //    FullName = UserID2.FullName,
                                            //    PersonalCode = (int)UserID2.usr_Personal_ID,
                                            //    listMoalefe = new List<MoalefeInfo> { dastmozdexce4 }
                                            //};

                                            //usersMoalefeList.Add(userInfo);



                                            //lstMoalefeexcel2.UsersMoalefe = usersMoalefeList;
                                        }

                                        //}
                                        //else
                                        //{
                                        //    if (!string.IsNullOrWhiteSpace(item))
                                        //    {
                                        //        MoalefeInfo dastmozdexce4 = new MoalefeInfo
                                        //        {
                                        //            MoalefeTitle = item,
                                        //            MoalefeValue = "0"
                                        //        };

                                        //        List<MoalefeUserInfo> usersMoalefeList =new List<MoalefeUserInfo>();
                                        //        MoalefeUserInfo userInfo = new MoalefeUserInfo
                                        //        {
                                        //            FullName = UserID2.FullName,
                                        //            PersonalCode = (int)UserID2.usr_Personal_ID,
                                        //            listMoalefe = new List<MoalefeInfo> { dastmozdexce4 }
                                        //        };

                                        //        usersMoalefeList.Add(userInfo);

                                        //        FunctionModel functionModel = new FunctionModel
                                        //        {
                                        //            SabtHeader = new Header
                                        //            {
                                        //                MoalfeVal_Month = Convert.ToInt32(month),
                                        //                MoalfeVal_Year = Convert.ToInt32(year),
                                        //                PeymanName = PymN.pec_Title,
                                        //                BastehName = BastehName2.Title
                                        //            },
                                        //            UsersMoalefe = usersMoalefeList
                                        //        };

                                        //        lstMoalefeexcel2.Add(functionModel);
                                        //    }
                                        //}
                                        shomarande++;

                                    }
                                    lstMoalefeexcel2.UsersMoalefe.Add(obj); //اطلاعات یک نفر اضافه شود 

                                }
                            }
                            else
                            {
                                return Content("این شیت فاقد سطر می باشد");
                            }
                        }
                        else
                        {
                            return Content("هیچ شیتی در این اکسل وجود ندارد");
                        }

                        transaction.Commit();
                        //if (userid == rus3)
                        //{
                        //    var matchedRows = db.tbMoalefeValuePishkhan
                        //        .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                        //        .ToList();

                        //    foreach (var row in matchedRows)
                        //    {
                        //        row.mlfval_Value = "1";
                        //    }

                        //    db.SaveChanges();
                        //}

                        //TempData["DataFromExcel"] = lstMoalefeexcel2;

                        //return "با موفقیت ثبت شد";
                        return PartialView("~/Areas/Salaries/Views/Moalefe/GetDataFromExcel_Moalefe.cshtml", lstMoalefeexcel2);

                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return Content("اکسل دارای خطا است، دقت کنید عنوان ستون نال دیگری جد نداشته باشد");
                    }
                }
            }
        }
        public async Task<ActionResult> GetDataFromExcel_Moalefesistan(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate, int Basteh, int PeymanID)
        {
            string name, family, personalid, year, month = "";
            string value;
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToList();
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);

            }

            int numberOfUsers = UserList.Count + 3;

            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();

            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh && s.Deleted != true).FirstOrDefault();

            var us = db.tbReffrenceSave
                .Where(p => p.FK_PeymanID == number)
                .SelectMany(s => s.tbReffrenceSaveLevel)
                .Where(s => s.ID == Basteh)
                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
                .Where(p => p.FK_RRSave == us.FK_RRSave)

                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var rus3 = db.tbReffrenceSaveLevelUser
                .Where(p => p.FK_LevelID == us2.ID)
                .OrderByDescending(p => p.ID)
                .Select(p => p.FK_UserID)
                .FirstOrDefault();

            List<string> lstMoalefe = new List<string>();
            List<int> lstMoalefeint = new List<int>();

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
                    }
                }
            }

            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.UsersMoalefe = new List<MoalefeUserInfo>();
            List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
            lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {

                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;

                            if (Rows.Length >= 1)
                            {
                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var count = workbook.Worksheets[0].Rows.Count();
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;


                                if (count == 1)
                                {
                                    return Content("فایل اکسل فاقد اطلاعات می باشد");
                                }
                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    //IRange cell = title[i];
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }
                                }
                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return Content(" لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ");
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return Content(" لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ");
                                }

                                int intYear = System.Convert.ToInt32(year);
                                int intMonth = System.Convert.ToInt32(month);

                                var find = db.tbSavedFunctions
                                    .Where(p => p.svdfunc_BastehID == Basteh &&
                                                p.svdfunc_pymnID == number &&
                                                p.Final_Accept == true &&
                                                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == intYear && m.MoalfeVal_Month == intMonth))
                                    .FirstOrDefault();

                                if (find != null)
                                {
                                    return Content("شما در این بسته ثبت و تایید شده است ");
                                }
                                var sabtt = db.tbSavedFunctions
                                    .Where(p => p.svdfunc_BastehID == Basteh &&
                                                p.svdfunc_pymnID == number &&
                                                p.svdfunc_UserSaveID == userid &&

                                                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == intYear && m.MoalfeVal_Month == intMonth))
                                    .FirstOrDefault();
                                if (sabtt != null)
                                {
                                    return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
                                }

                                lstMoalefeexcel2.SabtHeader = new Header
                                {
                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                    PeymanName = PymN.pec_Title,
                                    BastehName = BastehName2.Title,
                                    FromDate = FromDate,
                                    ToDate = ToDate,


                                };
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return Content(" لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ");
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault().usr_ID;
                                    var UserID2 = await db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefaultAsync();
                                    int shomarande = 3;
                                    //FunctionModel functionModel = new FunctionModel
                                    //{
                                    //    SabtHeader = new Header
                                    //    {
                                    //        MoalfeVal_Month = Convert.ToInt32(month),
                                    //        MoalfeVal_Year = Convert.ToInt32(year),
                                    //        PeymanName = PymN.pec_Title,
                                    //        BastehName = BastehName2.Title
                                    //    },
                                    //    UsersMoalefe = usersMoalefeList
                                    //};

                                    var obj = new MoalefeUserInfo();
                                    obj.FullName = UserID2.FullName;
                                    obj.PersonalCode = System.Convert.ToInt32(UserID2.usr_Personal_ID);
                                    obj.listMoalefe = new List<MoalefeInfo>();
                                    foreach (var item in lstMoalefeint)
                                    {
                                        var MoalefeID = await db.tbContractMoalefeDastmozdi.Where(p => p.md_ID == item).FirstOrDefaultAsync();

                                        var Value = row.Cells[shomarande];
                                        //if (Value.Value != null && Value.Value.ToString() != "")
                                        //{
                                        if (!string.IsNullOrWhiteSpace(MoalefeID.md_Title))
                                        {
                                            //var MoalefeID =await db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(item)).FirstOrDefaultAsync(s=>s.md_ID);
                                            //var MoalefeID = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(s => s.md_ID).FirstOrDefault();


                                            value = Value.Value.ToString() ?? "0";

                                            // MoalefeInfo dastmozdexce4 =
                                            obj.listMoalefe.Add(new MoalefeInfo
                                            {
                                                MoalefeTitle = MoalefeID.md_Title,
                                                MoalefeValue = value
                                            });

                                            //List<MoalefeUserInfo> usersMoalefeList = new List<MoalefeUserInfo>();
                                            //MoalefeUserInfo userInfo = new MoalefeUserInfo
                                            //{
                                            //    FullName = UserID2.FullName,
                                            //    PersonalCode = (int)UserID2.usr_Personal_ID,
                                            //    listMoalefe = new List<MoalefeInfo> { dastmozdexce4 }
                                            //};

                                            //usersMoalefeList.Add(userInfo);



                                            //lstMoalefeexcel2.UsersMoalefe = usersMoalefeList;
                                        }

                                        //}
                                        //else
                                        //{
                                        //    if (!string.IsNullOrWhiteSpace(item))
                                        //    {
                                        //        MoalefeInfo dastmozdexce4 = new MoalefeInfo
                                        //        {
                                        //            MoalefeTitle = item,
                                        //            MoalefeValue = "0"
                                        //        };

                                        //        List<MoalefeUserInfo> usersMoalefeList =new List<MoalefeUserInfo>();
                                        //        MoalefeUserInfo userInfo = new MoalefeUserInfo
                                        //        {
                                        //            FullName = UserID2.FullName,
                                        //            PersonalCode = (int)UserID2.usr_Personal_ID,
                                        //            listMoalefe = new List<MoalefeInfo> { dastmozdexce4 }
                                        //        };

                                        //        usersMoalefeList.Add(userInfo);

                                        //        FunctionModel functionModel = new FunctionModel
                                        //        {
                                        //            SabtHeader = new Header
                                        //            {
                                        //                MoalfeVal_Month = Convert.ToInt32(month),
                                        //                MoalfeVal_Year = Convert.ToInt32(year),
                                        //                PeymanName = PymN.pec_Title,
                                        //                BastehName = BastehName2.Title
                                        //            },
                                        //            UsersMoalefe = usersMoalefeList
                                        //        };

                                        //        lstMoalefeexcel2.Add(functionModel);
                                        //    }
                                        //}
                                        shomarande++;

                                    }
                                    lstMoalefeexcel2.UsersMoalefe.Add(obj); //اطلاعات یک نفر اضافه شود 

                                }
                            }
                            else
                            {
                                return Content("این شیت فاقد سطر می باشد");
                            }
                        }
                        else
                        {
                            return Content("هیچ شیتی در این اکسل وجود ندارد");
                        }

                        transaction.Commit();
                        //if (userid == rus3)
                        //{
                        //    var matchedRows = db.tbMoalefeValuePishkhan
                        //        .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                        //        .ToList();

                        //    foreach (var row in matchedRows)
                        //    {
                        //        row.mlfval_Value = "1";
                        //    }

                        //    db.SaveChanges();
                        //}

                        //TempData["DataFromExcel"] = lstMoalefeexcel2;

                        //return "با موفقیت ثبت شد";
                        return PartialView("~/Areas/Salaries/Views/Moalefe/GetDataFromExcel_Moalefesistan.cshtml", lstMoalefeexcel2);

                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return Content("اکسل دارای خطا است، دقت کنید عنوان ستون نال دیگری جد نداشته باشد");
                    }
                }
            }
        }


        public ActionResult ImportExcel_Moalefe_Forsorat(HttpPostedFileBase files, int svdfunc_BastehID, int svdfunc_pymnID, DateTime FromDate, DateTime ToDate)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = GetMoalfefromExcel_forsabt_forsorat(files, svdfunc_BastehID, svdfunc_pymnID, ToDate, FromDate);
            return Content(Message);


        }


        public ActionResult ImportExcele_TaedSorat(HttpPostedFileBase files, int svdfunc_BastehID, int svdfunc_pymnID)
        {
            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = GetMoalfefromExcel_fotaead_forsorat(files, svdfunc_BastehID, svdfunc_pymnID);
            return Content(Message);
        }




        public string GetMoalfefromExcel_forsabt_forsorat(HttpPostedFileBase files, int Basteh, int PeymanID, DateTime ToDate, DateTime FromDate)
        {

            string name, family, personalid, year, month = "";
            string value;
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var city = db.tbpeymancities.Where(p => listpymn.Contains(p.FK_PYMN)).Select(p => p.FK_City).ToList();
            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<string> lstMoalefe23 = new List<string>();

            List<tbSoratSavefromExcelMoalfe> lstMoalefeexcel = new List<tbSoratSavefromExcelMoalfe>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;
                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var cont = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == PeymanID).ToList();

                                for (int i = 1; i <= cont.Count; i++)
                                {
                                    lstMoalefe23.Add(title[i].Value.ToString().TrimEnd('\n'));

                                }
                                for (int i = cont.Count + 1; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year);
                                int numberOfUsers = city.Count + 3;
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString().TrimEnd('\n');
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 1;

                                    foreach (var item in lstMoalefe23)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID23 = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = System.Convert.ToDouble(value),
                                                    Value = 0,
                                                    FK_price = MoalefeID23,


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = 0,
                                                    Value = 0,
                                                    FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = null,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                valuenergh = 0,
                                                Value = 0,
                                                FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }


                                    foreach (var item in lstMoalefe)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = MoalefeID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = System.Convert.ToDouble(value),


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = 0,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                Value = 0,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }
                                }

                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }


                                tbsaveSoratBasteh savedfunctions = new tbsaveSoratBasteh();

                                savedfunctions.fk_Basteh = Basteh;
                                savedfunctions.Fromdata = FromDate;
                                savedfunctions.Todata = ToDate;
                                savedfunctions.Data_time = DateTime.Now;
                                savedfunctions.fk_pymn = PeymanID;
                                savedfunctions.Fk_usrsave = userid2;
                                savedfunctions.ISsubmit = false;
                                savedfunctions.For_Accept = false;
                                savedfunctions.Final_Accept = false;

                                savedfunctions.tbSoratSavefromExcelMoalfe = lstMoalefeexcel;
                                if (files != null)
                                {


                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/SoratSalari/" + filename));
                                    savedfunctions.FileSystemName = filename;
                                    savedfunctions.FileName = files.FileName;



                                }
                                var Month4 = System.Convert.ToInt32(month);
                                var yer = System.Convert.ToInt32(year);
                                if (moaleSorat.Create(savedfunctions) == "True")
                                {

                                    var r = db.tbFinalCheckSoorat.Where(p => p.Month == Month4 && p.Year == yer && p.FK_Peyman == PeymanID).FirstOrDefault();
                                    if (r != null)
                                    {
                                        r.FinalSave = true;
                                        db.SaveChanges();
                                    }
                                    else
                                    {
                                        tbFinalCheckSoorat detail = new tbFinalCheckSoorat
                                        {
                                            Month = Month4,
                                            Year = yer,
                                            FinalSave = true
                                        };

                                        db.tbFinalCheckSoorat.Add(detail); // Add the new instance to the context
                                        db.SaveChanges(); // Save changes to the database
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                }

                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }

                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }
                        transaction.Commit();
                        return "True";




                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }
            }
        }



        public async Task<string> GetMoalfefromExcel_forsabt_forsoratedit(HttpPostedFileBase files, int Basteh, int PeymanID, DateTime ToDate, DateTime FromDate)
        {

            string name, family, personalid, year, month = "";
            string value;
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var city = db.tbpeymancities.Where(p => listpymn.Contains(p.FK_PYMN)).Select(p => p.FK_City).ToList();
            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<string> lstMoalefe23 = new List<string>();

            List<tbSoratSavefromExcelMoalfe> lstMoalefeexcel = new List<tbSoratSavefromExcelMoalfe>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;
                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var cont = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == PeymanID).ToList();

                                for (int i = 1; i <= cont.Count; i++)
                                {
                                    lstMoalefe23.Add(title[i].Value.ToString().TrimEnd('\n'));

                                }
                                for (int i = cont.Count+1; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year);
                                int numberOfUsers = city.Count + 3;
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString().TrimEnd('\n');
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 1;

                                    foreach (var item in lstMoalefe23)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID23 = db.tbPeymanContractPrice.Where(p => p.UnitTitle==item&&p.FKPeymanID== PeymanID).FirstOrDefault().ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = System.Convert.ToDouble(value),
                                                    Value = 0,
                                                    FK_price = MoalefeID23,


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi =null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = 0,
                                                    Value = 0,
                                                    FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = null,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                valuenergh = 0,
                                                Value = 0,
                                                FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }


                                    foreach (var item in lstMoalefe)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = MoalefeID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = System.Convert.ToDouble(value),


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = 0,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                Value = 0,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }
                                }

                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }


                                tbsaveSoratBasteh savedfunctions = new tbsaveSoratBasteh();

                                savedfunctions.fk_Basteh = Basteh;
                                savedfunctions.Fromdata = FromDate;
                                savedfunctions.Todata = ToDate;
                                savedfunctions.Data_time = DateTime.Now;
                                savedfunctions.fk_pymn = PeymanID;
                                savedfunctions.Fk_usrsave = userid2;
                                savedfunctions.ISsubmit = false;
                                savedfunctions.For_Accept = false;
                                savedfunctions.Final_Accept = false;

                                savedfunctions.tbSoratSavefromExcelMoalfe = lstMoalefeexcel;
                                if (files != null)
                                {


                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/SoratSalari/" + filename));
                                    savedfunctions.FileSystemName = filename;
                                    savedfunctions.FileName = files.FileName;



                                }
                                var Month4 = System.Convert.ToInt32(month);
                                var yer = System.Convert.ToInt32(year);
                                if (moaleSorat.Create(savedfunctions) == "True")
                                {

                                    var r = db.tbFinalCheckSoorat.Where(p => p.Month == Month4 && p.Year == yer && p.FK_Peyman == PeymanID).FirstOrDefault();
                                    if (r != null)
                                    {
                                        r.FinalSave = true;
                                        db.SaveChanges();
                                    }
                                    else
                                    {
                                        tbFinalCheckSoorat detail = new tbFinalCheckSoorat
                                        {
                                            Month = Month4,
                                            Year = yer,
                                            FinalSave = true
                                        };

                                        db.tbFinalCheckSoorat.Add(detail); // Add the new instance to the context
                                        db.SaveChanges(); // Save changes to the database
                                    }
                                }
                                else
                                {
                                    transaction.Rollback();
                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                }

                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }

                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }
                        transaction.Commit();
                        return "True";




                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }
            }
        }


        public string GetMoalfefromExcel_fotaead_forsorat(HttpPostedFileBase files, int Basteh, int PeymanID)
        {

            string name, family, personalid, year, month = "";
            string value;
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var city = db.tbpeymancities.Where(p => listpymn.Contains(p.FK_PYMN)).Select(p => p.FK_City).ToList();
            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<string> lstMoalefe23 = new List<string>();

            List<tbSoratSavefromExcelMoalfe> lstMoalefeexcel = new List<tbSoratSavefromExcelMoalfe>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;
                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var countt = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == PeymanID).ToList();
                                for (int i = 1; i <= countt.Count; i++)
                                {
                                    lstMoalefe23.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = countt.Count+1; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year);
                                int numberOfUsers = city.Count + 3;
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 1;

                                    foreach (var item in lstMoalefe23)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID23 = db.tbPeymanContractPrice.Where(p => p.UnitTitle==item&&p.FKPeymanID== PeymanID).FirstOrDefault().ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = System.Convert.ToDouble(value),
                                                    Value=0,
                                                    FK_price= MoalefeID23,


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = null,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    valuenergh = 0,
                                                    Value = 0,
                                                    FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = null,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                valuenergh = 0,
                                                Value = 0,
                                                FK_price = db.tbPeymanContractPrice.Where(p => p.UnitTitle == item && p.FKPeymanID == PeymanID).FirstOrDefault().ID,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }

                                    foreach (var item in lstMoalefe)
                                    {
                                        var Value = row.Cells[shomarande];
                                        var UserID = db.tbCities.Where(p => p.Name == name).FirstOrDefault().ID;

                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID;


                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = MoalefeID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = System.Convert.ToDouble(value),


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID,
                                                    FK_city = UserID,
                                                    Month = System.Convert.ToInt32(month),
                                                    Year = System.Convert.ToInt32(year),
                                                    Value = 0,

                                                    //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                                };
                                                if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title==item).FirstOrDefault().md_ID,
                                                FK_city = UserID,
                                                Month = System.Convert.ToInt32(month),
                                                Year = System.Convert.ToInt32(year),
                                                Value = 0,
                                                //FK_SavedFunctionsID = sabtt.svdfunc_ID
                                            };
                                            if (moalefeexcelRepoSorat.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }
                                }

                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }

                                var c = db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == Basteh).FirstOrDefault();

                                var v = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == c.ID).OrderByDescending(D => D.ID).FirstOrDefault();
                                var b = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == v.ID).OrderByDescending(x => x.ID).Select(s => s.FK_User).FirstOrDefault();


                                tbsaveSoratBasteh savedfunctions = new tbsaveSoratBasteh();

                                savedfunctions.fk_Basteh = Basteh;
                                savedfunctions.Fromdata = null;
                                savedfunctions.Todata = null;
                                savedfunctions.Data_time = DateTime.Now;
                                savedfunctions.fk_pymn = PeymanID;
                                savedfunctions.Fk_usrsave = userid2;
                                savedfunctions.For_Accept = true;
                                if (b == userid2)
                                {
                                    savedfunctions.Final_Accept = true;
                                    savedfunctions.ISsubmit = true;



                                }
                                else
                                {
                                    savedfunctions.Final_Accept = false;
                                    savedfunctions.ISsubmit = false;
                                }

                                savedfunctions.tbSoratSavefromExcelMoalfe = lstMoalefeexcel;
                                if (files != null)
                                {


                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/SoratSalari/" + filename));
                                    savedfunctions.FileSystemName = filename;
                                    savedfunctions.FileName = files.FileName;




                                }

                                if (moaleSorat.Create(savedfunctions) == "True")
                                {
                                    if (b == userid2)
                                    {

                                        var cc = db.tbsaveSoratBasteh.Where(p => p.fk_Basteh == Basteh && p.fk_pymn == PeymanID && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == montht && s.Year == yyer)).ToList();
                                        foreach (var it in cc)
                                        {
                                            tbsaveSoratBasteh tb = new tbsaveSoratBasteh();
                                            it.ISsubmit = true;
                                            db.SaveChanges();
                                        }
                                        var Month4 = System.Convert.ToInt32(month);
                                        var yer = System.Convert.ToInt32(year);
                                        var r = db.tbFinalCheckSoorat.Where(p => p.Month == Month4 && p.Year == yer && p.FK_Peyman == PeymanID).FirstOrDefault();
                                        if (r != null)

                                        {
                                            r.FinalAccept = true;
                                            db.SaveChanges();
                                        }
                                    }


                                }
                                else
                                {
                                    transaction.Rollback();
                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                }

                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }

                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }
                        transaction.Commit();
                        return "True";




                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }
            }

        }

        public async Task<ActionResult> ImportExcel_Moalefe_forsabt(HttpPostedFileBase files, int svdfunc_BastehID, int svdfunc_pymnID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_Fosabt(files, svdfunc_BastehID, svdfunc_pymnID);

            return Content(Message);
        }
        public async Task<ActionResult> ImportExcel_Moalefe_forsabtsaabtjadid(HttpPostedFileBase files, int id, int month,int year)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }
            //    int usr = 0;
            //var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            //if (cookie_user != null)
            //{
            //    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //    using (SaabEntities db = new SaabEntities())
            //    {
            //        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
            //        if (User != null)
            //        {
            //            usr = User.usr_ID;
            //        }
            //    }
            //}

            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            //var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id).FirstOrDefault();
            if (tbmarahesabt4 != null)
            {
                if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 1)){

                    Message = await GetDataFromExcel_Moalefe_Fosabtsabttjadid(files, id, month, year);


                }
                else if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_taghiz != null))
                {


                    Message = await GetDataFromExcelMachinsOrTools4(files, id, month, year);

                }

                else if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 4))
                {


                    Message = await GetDataFromExcel_Moalefe_Fosabtsabttjadidtestsorat(files, id, month, year);

                }

                else if (tbmarahesabt4.tbmarahlsabtbastedit2.tbmarahlsabtbastedit1.tbmarahlsabtbastedit3.Any(s => s.FK_moalfe != null&&s.tbmarahlsabtbastedit1.typemoalfeh==2))
                {


                    Message = await GetDataFromExcel_Moalefe_Fosabtsabttjadidrialy(files, id, month, year);

                }
            }

            return Content(Message);
        }
        public async Task<ActionResult> ImportExcel_Moalefe_forsabtedit(HttpPostedFileBase files, int svdfunc_BastehID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_Fosabtedit(files, svdfunc_BastehID);

            return Content(Message);
        }









        public async Task<ActionResult> ImportExcel_Moalefe_forstaededit(HttpPostedFileBase files, int svdfunc_BastehID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_Fosabtedit(files, svdfunc_BastehID);

            return Content(Message);
        }

        public async Task<ActionResult> ImportExcel_Moalefe_forstaesbat(HttpPostedFileBase files, int svdfunc_BastehID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_Fosabtedit(files, svdfunc_BastehID);

            return Content(Message);
        }



        public async Task<string> GetDataFromExcel_Moalefe_Fosabt(HttpPostedFileBase files, int Basteh, int PeymanID)
        {
            string name, family, personalid, year, month = "";
            string value;


            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }

            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).ToList();
            List<tbUsers> UserList = new List<tbUsers>();
            UserList= usrID.Select(p => p.tbUsers).ToList();
            //foreach (var item in usrID)
            //{
            //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    UserList.Add(user);
            //}

            int numberOfUsers = UserList.Count + 3;
            List<int> lstMoalefeint = new List<int>();

            int number = PeymanID;
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
            var us4 = db.tbReffrenceSave
        .Where(p => p.FK_PeymanID == PeymanID)
        .SelectMany(s => s.tbReffrenceSaveLevel)
        .Where(s => s.ID == Basteh)
        .OrderByDescending(p => p.ID)
        .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
                .Where(p => p.FK_RRSave == us4.FK_RRSave)
                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var rus3 = db.tbReffrenceSaveLevelUser
                .Where(p => p.FK_LevelID == us2.ID)
                .OrderByDescending(p => p.ID)
                .Select(p => p.FK_UserID)
                .FirstOrDefault();
            var User2 = new tbUsers();
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                //if (User2 != null)
                //{

                //    userid = User2.usr_ID;
                //    if (userid == rus3)
                //    {

                //        var matchedRows = db.tbMoalefeValuePishkhan
                //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                //            .ToList();

                //        foreach (var row1 in matchedRows)
                //        {
                //            row1.mlfval_Value = "1";

                //        }

                //        db.SaveChanges();

                //    }
                //}

            }


            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {

                        var tbmarahlsabtbastedit6 = db.tbmarahlsabtbastedit6.ToList();


                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                          

                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year); var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == Basteh &&
p.svdfunc_pymnID == PeymanID &&
p.tbMoalefeDastmozdiValueFromExcel
.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();
                                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                                List<tbSavedFunctions> savedfunctions2 = new List<tbSavedFunctions>();


                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }
                                if (cookie_user2 != null)
                                {
                                    //int userid = 0;
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                                    }
                                    if (User2 != null)
                                    {

                                        int userid4 = User2.usr_ID;
                                        if (userid4 == rus3)
                                        {
                                            savedfunctions.Final_Sabt = true;



                                        }
                                    }
                                }
                                var month2= System.Convert.ToInt32(month);
                                var year2= System.Convert.ToInt32(year);
                                //}
                                //else//add
                                //{
                                var find = db.tbSavedFunctions
    .Where(p => p.svdfunc_BastehID == Basteh &&
                p.svdfunc_pymnID == number &&
                p.Final_Accept == true &&
                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();

                                if (find != null)
                                {
                                    return " این بسته ثبت و تایید شده است ";
                                }
                                var sabttttt = db.tbSavedFunctions
                                    .Where(p => p.svdfunc_BastehID == Basteh &&
                                                p.svdfunc_pymnID == number &&
                                                p.svdfunc_UserSaveID == userid2 &&

                                                p.tbMoalefeDastmozdiValueFromExcel
                                                                                                      .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();

                                if (sabttttt != null)
                                {
                                    return "شما قبلا این بسته را در این سال و ماه ثبت کرده اید ";
                                }
                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }

                                }
                                    savedfunctions.svdfunc_BastehID = Basteh;
                                ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                savedfunctions.svdfunc_pymnID = PeymanID;
                                savedfunctions.svdfunc_UserSaveID = userid2;
                                savedfunctions.svdfunc_IsSubmmit = false;
                                //savedfunctions.For_Accept = true;
                                savedfunctions.Final_Accept = false;


                                
                                 
                                      var segment = files.FileName.Split('.');
                                            string file_type = segment[segment.Length - 1];
                                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                            savedfunctions.svdfunc_FileNameExcel = filename;
                                            savedfunctions.svdfunc_FileSystemNameExcel = files.FileName;


                                        
                                       
                                    
                              


                                //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                savedfunctions2.Add(savedfunctions);
                                if (savedfunctionRepo.Create(savedfunctions) == "True")
                                {
                                }
                                else
                                {
                                    savedfunctionRepo.Create5(savedfunctions2);
                                }


                                    List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                                    lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = 
                                        db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).Select(p => p.usr_ID).FirstOrDefault();
                                    int shomarande = 3;


                                    foreach (var item in lstMoalefeint)
                                    {
                                        var findmoalf = tbmarahlsabtbastedit6.Where(p => p.FK_moalfeh2 == item).FirstOrDefault();
                                        var fi = item;
                                        if (findmoalf != null)
                                        {
                                            fi = (int)findmoalf.FK_moalfe;
                                        }
                                        
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID =
                                                    //lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault();
                                                //lstmoalafe2.Where(p => p.md_ID==item).FirstOrDefault();



                                                value = Value.Value.ToString();
                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = fi,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = System.Convert.ToDouble(value),
                                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,



                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}


                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                    MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID== fi).Select(p => p.md_ID).FirstOrDefault(),

                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0,

                                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                            {
                                                //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID== fi).Select(p => p.md_ID).FirstOrDefault(),

                                                MoalfeVal_FKUser = UserID,
                                                MoalfeVal_Month = System.Convert.ToInt32(month),
                                                MoalfeVal_Year = System.Convert.ToInt32(year),
                                                MoalfeVal_Value = 0,
                                                FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                            };
                                            lstMoalefeexcel.Add(dastmozdexcel);

                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                            //{
                                            //}
                                            //else
                                            //{
                                            //    transaction.Rollback();
                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                            //}
                                        }

                                        shomarande++;
                                    }



                                }
                                if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.Create2Async(lstMoalefeexcel) > 0)
                                    {
                                        // Saving successful, number of saved items is returned by Create2Async
                                        // You can use the returned value here, for example:
                                        //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                                        //Console.WriteLine($"Successfully saved {savedCount} items.");
                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }
                                //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();








                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }




        public async Task<string> GetDataFromExcel_Moalefe_Fosabtsabttjadidtestsorat(HttpPostedFileBase files, int id, int month2, int year2)
        {
            string name, family, personalid, year, month = "";
            string value;
            int usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                    }
                }
            }

            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id && p.del != true).FirstOrDefault();
            var tbmarahesabt41111 = db.tbmarahesabt4.Where(p => p.del != true).ToList();

            //var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).ToList();
            //List<tbUsers> UserList = new List<tbUsers>();
            //UserList = tbAdamAdam.Select(p => p.tbUsers).ToList();
            ////foreach (var item in usrID)
            ////{
            ////    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            ////    UserList.Add(user);
            ////}

            //int numberOfUsers = UserList.Count + 3;
            List<int> lstMoalefeint = new List<int>();
            List<int> lstMoalefeintaval = new List<int>();

            //int number = PeymanID;
            //var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
            List<tbSoratSavefromExcelMoalfe> tbSoratSavefromExcelMoalfe = new List<tbSoratSavefromExcelMoalfe>();



            //    var us4 = db.tbReffrenceSave
            //.Where(p => p.FK_PeymanID == PeymanID)
            //.SelectMany(s => s.tbReffrenceSaveLevel)
            //.Where(s => s.ID == Basteh)
            //.OrderByDescending(p => p.ID)
            //.FirstOrDefault();
            //var us2 = db.tbReffrenceSaveLevel
            //    .Where(p => p.FK_RRSave == us4.FK_RRSave)
            //    .OrderByDescending(p => p.ID)
            //    .FirstOrDefault();
            //var rus3 = db.tbReffrenceSaveLevelUser
            //    .Where(p => p.FK_LevelID == us2.ID)
            //    .OrderByDescending(p => p.ID)
            //    .Select(p => p.FK_UserID)
            //    .FirstOrDefault();
            var User2 = new tbUsers();
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                //if (User2 != null)
                //{

                //    userid = User2.usr_ID;
                //    if (userid == rus3)
                //    {

                //        var matchedRows = db.tbMoalefeValuePishkhan
                //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                //            .ToList();

                //        foreach (var row1 in matchedRows)
                //        {
                //            row1.mlfval_Value = "1";

                //        }

                //        db.SaveChanges();

                //    }
                //}

            }


            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        var tbmarahlsabtbastedit6 = db.tbmarahlsabtbastedit6.ToList();



                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;
                                //var count = workbook.Worksheets[0].Rows.Count();

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                ///// year
                                //var Yearrow = workbook.Worksheets[0].Rows[0];
                                //var Year = Yearrow.Cells[1];
                                //if (Year.Value != null || Year.Value.ToString() != "")
                                //{
                                //    year = Year.Value.ToString();
                                //}
                                //else
                                //{
                                //    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                //}
                                //year = year2;
                                //////// month
                                //var Monthrow = workbook.Worksheets[0].Rows[1];
                                //var Month = Monthrow.Cells[1];
                                //if (Month.Value != null || Month.Value.ToString() != "")
                                //{
                                //    month = Month.Value.ToString();
                                //}
                                //else
                                //{
                                //    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                //}

                                //month = month2;
                                var montht = month2;
                                var yyer = year2;
                                var sabtt = db.tbsaveSoratBasteh
.Where(p => p.FK_Bastehedit == id &&
p.fk_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
p.tbSoratSavefromExcelMoalfe
.Any(m => m.Year == yyer && m.Month == montht)).OrderByDescending(sf => sf.ID)
.FirstOrDefault();
                                tbsaveSoratBasteh savedfunctions = new tbsaveSoratBasteh();
                                List<tbsaveSoratBasteh> savedfunctions2 = new List<tbsaveSoratBasteh>();


                                int userid2 = 0;
                                var cookie_user222 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user222 != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user222.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }
                                if (cookie_user2 != null)
                                {
                                    //int userid = 0;
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                                    }
                                    if (User2 != null)
                                    {

                                        //int userid4 = User2.usr_ID;
                                        //if (userid4 == rus3)
                                        //{
                                        //    savedfunctions.Final_Sabt = true;



                                        //}
                                    }
                                }

                                var month22 = month2;
                                var year22 = year2;
                                //}
                                //else//add
                                //{

                                var maxNumber = tbmarahesabt41111
                                 .Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true)
                                 .Max(p => (int?)p.number);

                                var naheii = 0;

                                if (maxNumber.HasValue)
                                {
                                    naheii = maxNumber.Value; // استفاده از Value برای دسترسی به مقدار
                                }
                                var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
                                var sabt = tbmarhesabt3.marahelsabt;
                                var taeed = tbmarhesabt3.taeed;
                                int vaziat = 0;
                                int vaziat2 = 0;

                                if (tbmarahesabt4.sabt == true)
                                {
                                    vaziat = 1;
                                    vaziat2 = (int)tbmarahesabt4.number;

                                }
                                else if (tbmarahesabt4.taeed == true)
                                {
                                    vaziat = 2;
                                    vaziat2 = (int)tbmarahesabt4.number;
                                }

                                var find = db.tbsaveSoratBasteh
    .Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City &&
    p.FK_Bastehedit == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST &&
                p.fk_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                p.Final_Accept == true && p.del != true &&

                p.tbSoratSavefromExcelMoalfe
                                                    .Any(m => m.Year == year2 && m.Month == month2)).FirstOrDefault();
                                if (vaziat == 2)
                                {
                                    if (find != null && vaziat2 != naheii)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
                                else
                                {
                                    if (find != null)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
//                                var findcontrol = db.tbsaveSoratBasteh
//.Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
//&& p.del != true &&
//       p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
//       p.FK_control != null &&
//       p.tbMoalefeValueFish
//                                           .Any(m => m.mlfvlfsh_Year == year2 && m.mlfvlfsh_Month == month2)).FirstOrDefault();

//                                if (findcontrol != null)
//                                {
//                                    return " این بسته در مرحله کنترل   شده است ";
//                                }



                                var findtaeed = db.tbsaveSoratBasteh
  .Where(p => p.FK_tbmarahesabt4 == id &&
              p.fk_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
              p.For_Accept == true &&
              p.del != true &&
              p.tbSoratSavefromExcelMoalfe
                                                  .Any(m => m.Year == year2 && m.Month == month2)).FirstOrDefault();
                                if (vaziat != 2)
                                {
                                    if (findtaeed != null && vaziat2 != naheii)
                                    {
                                        return " این بسته ثبت و تایید شده است ";
                                    }
                                }
                                //else
                                //{
                                //    if (findtaeed != null)
                                //    {
                                //        return " این بسته ثبت و تایید شده است ";
                                //    }
                                //}
                                //if (findtaeed != null)
                                //{
                                //    return " این بسته ثبت و تایید شده است ";
                                //}
                                var sabttttt = db.tbsaveSoratBasteh
                                    .Where(p => p.FK_tbmarahesabt4 == id &&
                                                p.fk_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                                                p.Fk_usrsave == userid2 && p.del != true &&


                                                p.tbSoratSavefromExcelMoalfe
                                                                                                      .Any(m => m.Year == year2 && m.Month == month2)).FirstOrDefault();

                                //if (sabttttt != null)
                                //{
                                //    savedfunctions = sabttttt;
                                //}

                                var cont = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn).ToList();

                                //for (int i = 1; i <= cont.Count; i++)
                                //{
                                //    lstMoalefe23.Add(title[i].Value.ToString().TrimEnd('\n'));

                                //}
                                //for (int i = cont.Count + 1; i < title.Length; i++)
                                //{
                                //    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                //}

                                for (int i = cont.Count + 1; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < cont.Count+3; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeintaval.Add(cellValue);
                                    }

                                }
                                for (int i = cont.Count + 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }

                                }
                                if (sabttttt == null)
                                {
                                    savedfunctions.FK_tbmarahesabt4 = id;
                                    savedfunctions.FK_Bastehedit = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    savedfunctions.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    savedfunctions.Data_time_edit = DateTime.Now;
                                    savedfunctions.fk_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    savedfunctions.Fk_usrsave = userid2;
                                    savedfunctions.submitt = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        savedfunctions.Final_Accept = true;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        savedfunctions.Final_Accept = false;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else
                                    {
                                        savedfunctions.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    savedfunctions.FileName = filename;
                                    savedfunctions.FileSystemName = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions2.Add(savedfunctions);
                                    db.tbsaveSoratBasteh.AddRange(savedfunctions2);
                                    await db.SaveChangesAsync();
                                }
                                else if (sabttttt != null)
                                {
                                    sabttttt.FK_tbmarahesabt4 = id;
                                    sabttttt.FK_Bastehedit = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    sabttttt.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    sabttttt.Data_time_edit = DateTime.Now;
                                    sabttttt.fk_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    sabttttt.Fk_usrsave = userid2;
                                    sabttttt.submitt = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        sabttttt.Final_Accept = true;
                                        sabttttt.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        sabttttt.Final_Accept = false;
                                        sabttttt.For_Accept = true;

                                    }
                                    else
                                    {
                                        sabttttt.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    sabttttt.FileName = filename;
                                    sabttttt.FileSystemName = files.FileName;
                                    await db.SaveChangesAsync();
                                    savedfunctions = sabttttt;
                                }
                                else
                                {
                                    savedfunctions.FK_tbmarahesabt4 = id;
                                    savedfunctions.FK_Bastehedit = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    savedfunctions.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    savedfunctions.Data_time_edit = DateTime.Now;
                                    savedfunctions.fk_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    savedfunctions.Fk_usrsave = userid2;
                                    savedfunctions.ISsubmit = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        savedfunctions.Final_Accept = true;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        savedfunctions.Final_Accept = false;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else
                                    {
                                        savedfunctions.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    savedfunctions.FileName = filename;
                                    savedfunctions.FileSystemName = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions2.Add(savedfunctions);
                                    db.tbsaveSoratBasteh.AddRange(savedfunctions2);
                                    await db.SaveChangesAsync();
                                }
                                var count2 = workbook.Worksheets[0].Rows.Count();
                                var usr12334 = db.tbUsers.ToList();

                                List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                                lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                                for (int i = 3; i < count2; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    //var Name = row.Cells[0];
                                    //if (Name.Value != null || Name.Value.ToString() != "")
                                    //{
                                    //    name = Name.Value.ToString();
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    //}
                                    int UserID = 0;
                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        if (int.TryParse(Family.Value.ToString(), out int cellValue))
                                        {
                                            UserID= cellValue;
                                        }
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    //var PersonalID = row.Cells[2];
                                    //if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    //{
                                    //    personalid = PersonalID.Value.ToString();
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    //}
                                
                                   
                                        //var perID = System.Convert.ToInt32(personalid);
                                        //var us11 = usr12334.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault();
                                      
                                        
                                                int shomarande = 3;

                                                var taeed23 = false;
                                                foreach (var item in lstMoalefeintaval)
                                                {
                                                    var findmoalf = tbmarahlsabtbastedit6.Where(p => p.FK_moalfeh2 == item).FirstOrDefault();
                                                    var fi = item;
                                                    if (findmoalf != null)
                                                    {
                                                        fi = (int)findmoalf.FK_moalfe;
                                                    }
                                                    var Value = row.Cells[shomarande];
                                                    if (Value.Value != null)
                                                    {
                                                        if (Value.Value.ToString() != "")
                                                        {
                                                double vbaluep = 0;
                                                if (double.TryParse(Value.Value.ToString(), out double cellValue))
                                                {
                                                    vbaluep = cellValue;
                                                }
                                                var MoalefeID = 0;
                                                            //lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault();
                                                            //lstmoalafe2.Where(p => p.md_ID==item).FirstOrDefault();

                                                            if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                                                            {
                                                                taeed23 = true;
                                                                //sabttttt.Final_Accept = true;
                                                                //sabttttt.For_Accept = true;

                                                            }
                                                            else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                                                            {
                                                                //sabttttt.Final_Accept = false;
                                                                //sabttttt.For_Accept = true;

                                                            }
                                                            else
                                                            {
                                                                //sabttttt.Final_Accept = false;

                                                            }
                                                            if (savedfunctions.Final_Accept == true)
                                                            {
                                                                taeed23 = true;


                                                            }
                                                            value = Value.Value.ToString();
                                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                            {
                                                                FK_price = fi,
                                                                FK_city = UserID,
                                                                Month = month22,
                                                                Year = year22,
                                                                valuenergh = System.Convert.ToDouble(vbaluep),
                                                                FK_SavedFunctionsID = savedfunctions.ID,
                                                                Value = 0,


                                                            };
                                                tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);
                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}


                                                        }

                                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                        {
                                                            if (savedfunctions.Final_Accept == true)
                                                            {
                                                                taeed23 = true;


                                                            }
                                                            value = Value.Value.ToString();
                                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                            {
                                                                FK_price = fi,
                                                                FK_city = UserID,
                                                                Month = month22,
                                                                Year = year22,
                                                                valuenergh = 0,
                                                                FK_SavedFunctionsID = savedfunctions.ID,
                                                                Value = 0,


                                                            };
                                                tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);

                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}
                                                        }
                                                    }

                                                    else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                    {
                                                        if (savedfunctions.Final_Accept == true)
                                                        {
                                                            taeed23 = true;


                                                        }

                                                        value = Value.Value.ToString();
                                                        tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                        {
                                                            FK_price = fi,
                                                            FK_city = UserID,
                                                            Month = month22,
                                                            Year = year22,
                                                            valuenergh = 0,
                                                            FK_SavedFunctionsID = savedfunctions.ID,
                                                            Value = 0,


                                                        };
                                            tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);

                                                        //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                        //{
                                                        //}
                                                        //else
                                                        //{
                                                        //    transaction.Rollback();
                                                        //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                        //}
                                                    }

                                                    shomarande++;
                                                }
                                    foreach (var item in lstMoalefeint)
                                    {
                                        var findmoalf = tbmarahlsabtbastedit6.Where(p => p.FK_moalfeh2 == item).FirstOrDefault();
                                        var fi = item;
                                        if (findmoalf != null)
                                        {
                                            fi = (int)findmoalf.FK_moalfe;
                                        }
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                double vbaluep = 0;
                                                if (double.TryParse(Value.Value.ToString(), out double cellValue))
                                                {
                                                    vbaluep = cellValue;
                                                }
                                                var MoalefeID = 0;
                                                //lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault();
                                                //lstmoalafe2.Where(p => p.md_ID==item).FirstOrDefault();

                                                if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                                                {
                                                    taeed23 = true;
                                                    //sabttttt.Final_Accept = true;
                                                    //sabttttt.For_Accept = true;

                                                }
                                                else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                                                {
                                                    //sabttttt.Final_Accept = false;
                                                    //sabttttt.For_Accept = true;

                                                }
                                                else
                                                {
                                                    //sabttttt.Final_Accept = false;

                                                }
                                                if (savedfunctions.Final_Accept == true)
                                                {
                                                    taeed23 = true;


                                                }
                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = fi,
                                                    FK_city = UserID,
                                                    Month = month22,
                                                    Year = year22,
                                                    Value = System.Convert.ToDouble(vbaluep),
                                                    FK_SavedFunctionsID = savedfunctions.ID,
                                                    valuenergh = 0,


                                                };
                                                tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}


                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {
                                                if (savedfunctions.Final_Accept == true)
                                                {
                                                    taeed23 = true;


                                                }
                                                value = Value.Value.ToString();
                                                tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                                {
                                                    FK_MoalfeDastmozdi = fi,
                                                    FK_city = UserID,
                                                    Month = month22,
                                                    Year = year22,
                                                    Value = 0,
                                                    FK_SavedFunctionsID = savedfunctions.ID,
                                                    valuenergh = 0,


                                                };
                                                tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);

                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {
                                            if (savedfunctions.Final_Accept == true)
                                            {
                                                taeed23 = true;


                                            }

                                            value = Value.Value.ToString();
                                            tbSoratSavefromExcelMoalfe dastmozdexcel = new tbSoratSavefromExcelMoalfe
                                            {
                                                FK_MoalfeDastmozdi = fi,
                                                FK_city = UserID,
                                                Month = month22,
                                                Year = year22,
                                                valuenergh = 0,
                                                FK_SavedFunctionsID = savedfunctions.ID,
                                                Value = 0,


                                            };
                                            tbSoratSavefromExcelMoalfe.Add(dastmozdexcel);

                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                            //{
                                            //}
                                            //else
                                            //{
                                            //    transaction.Rollback();
                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                            //}
                                        }

                                        shomarande++;
                                    }
                                            


                                        
                                    





                                }
                                if (sabttttt != null)
                                {
                                    var listtbMoalefeDastmozdiValueFromExcel = await db.tbSoratSavefromExcelMoalfe.Where(p => p.FK_SavedFunctionsID == sabttttt.ID).ToListAsync();
                                    db.tbSoratSavefromExcelMoalfe.RemoveRange(listtbMoalefeDastmozdiValueFromExcel);
                                    await db.SaveChangesAsync();
                                }
                                db.tbSoratSavefromExcelMoalfe.AddRange(tbSoratSavefromExcelMoalfe);
                                await db.SaveChangesAsync();
                                //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();








                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }





        public async Task<string> GetDataFromExcel_Moalefe_Fosabtsabttjadid(HttpPostedFileBase files, int id, int month2,int year2)
        {
            string name, family, personalid, year, month = "";
            string value;
            int usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                    }
                }
            }

            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id&&p.del!=true).FirstOrDefault(); 
            var tbmarahesabt41111 = db.tbmarahesabt4.Where(p =>  p.del != true).ToList();

            //var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).ToList();
            //List<tbUsers> UserList = new List<tbUsers>();
            //UserList = tbAdamAdam.Select(p => p.tbUsers).ToList();
            ////foreach (var item in usrID)
            ////{
            ////    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            ////    UserList.Add(user);
            ////}

            //int numberOfUsers = UserList.Count + 3;
            List<int> lstMoalefeint = new List<int>();

            //int number = PeymanID;
            //var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
        //    var us4 = db.tbReffrenceSave
        //.Where(p => p.FK_PeymanID == PeymanID)
        //.SelectMany(s => s.tbReffrenceSaveLevel)
        //.Where(s => s.ID == Basteh)
        //.OrderByDescending(p => p.ID)
        //.FirstOrDefault();
            //var us2 = db.tbReffrenceSaveLevel
            //    .Where(p => p.FK_RRSave == us4.FK_RRSave)
            //    .OrderByDescending(p => p.ID)
            //    .FirstOrDefault();
            //var rus3 = db.tbReffrenceSaveLevelUser
            //    .Where(p => p.FK_LevelID == us2.ID)
            //    .OrderByDescending(p => p.ID)
            //    .Select(p => p.FK_UserID)
            //    .FirstOrDefault();
            var User2 = new tbUsers();
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                //if (User2 != null)
                //{

                //    userid = User2.usr_ID;
                //    if (userid == rus3)
                //    {

                //        var matchedRows = db.tbMoalefeValuePishkhan
                //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                //            .ToList();

                //        foreach (var row1 in matchedRows)
                //        {
                //            row1.mlfval_Value = "1";

                //        }

                //        db.SaveChanges();

                //    }
                //}

            }


            //using (DbContextTransaction transaction = db.Database.BeginTransaction())
            //{
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {
                        var tbmarahlsabtbastedit6 = db.tbmarahlsabtbastedit6.ToList();



                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;
                                //var count = workbook.Worksheets[0].Rows.Count();

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year); 
                                var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == id &&
p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
p.tbMoalefeDastmozdiValueFromExcel
.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();
                                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                                List<tbSavedFunctions> savedfunctions2 = new List<tbSavedFunctions>();


                                int userid2 = 0;
                                var cookie_user222 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user222 != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user222.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }
                                if (cookie_user2 != null)
                                {
                                    //int userid = 0;
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                                    }
                                    if (User2 != null)
                                    {
                                   
                                        //int userid4 = User2.usr_ID;
                                        //if (userid4 == rus3)
                                        //{
                                        //    savedfunctions.Final_Sabt = true;



                                        //}
                                    }
                                }

                                var month22 = System.Convert.ToInt32(month);
                                var year22 = System.Convert.ToInt32(year);
                                //}
                                //else//add
                                //{

                                var maxNumber = tbmarahesabt41111
                                 .Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true)
                                 .Max(p => (int?)p.number);

                                var naheii = 0;

                                if (maxNumber.HasValue)
                                {
                                    naheii = maxNumber.Value; // استفاده از Value برای دسترسی به مقدار
                                }
                                var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
                                var sabt = tbmarhesabt3.marahelsabt;
                                var taeed = tbmarhesabt3.taeed;
                                int vaziat = 0;
                                int vaziat2 = 0;

                                if (tbmarahesabt4.sabt == true)
                                {
                                    vaziat = 1;
                                    vaziat2 = (int)tbmarahesabt4.number;

                                }
                                else if (tbmarahesabt4.taeed == true)
                                {
                                    vaziat = 2;
                                    vaziat2 =(int) tbmarahesabt4.number;
                                }
                               
                                var find = db.tbSavedFunctions
    .Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City &&
    p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST &&
                p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                p.Final_Accept == true && p.del != true &&

                p.tbMoalefeDastmozdiValueFromExcel
                                                    .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();
                                if(vaziat == 2)
                                {
                                    if (find != null && vaziat2 != naheii)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
                                else
                                {
                                    if (find != null)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
                                var findcontrol = db.tbmoalfefishexcel
.Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
&& p.del != true &&
       p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
       p.FK_control != null &&
       p.tbMoalefeValueFish
                                           .Any(m => m.mlfvlfsh_Year == year2 && m.mlfvlfsh_Month == month2)).FirstOrDefault();

                                //if (findcontrol != null)
                                //{
                                //    return " این بسته در مرحله کنترل   شده است ";
                                //}



                                var findtaeed = db.tbSavedFunctions
  .Where(p => p.FK_tbmarahesabt4 == id &&
              p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
              p.For_Accept == true &&
              p.del!=true&&
              p.tbMoalefeDastmozdiValueFromExcel
                                                  .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();
                                if (vaziat != 2)
                                {
                                    if (findtaeed != null && vaziat2 != naheii)
                                    {
                                        return " این بسته ثبت و تایید شده است ";
                                    }
                                }
                                //else
                                //{
                                //    if (findtaeed != null)
                                //    {
                                //        return " این بسته ثبت و تایید شده است ";
                                //    }
                                //}
                                //if (findtaeed != null)
                                //{
                                //    return " این بسته ثبت و تایید شده است ";
                                //}
                                var sabttttt = db.tbSavedFunctions
                                    .Where(p => p.FK_tbmarahesabt4 == id &&
                                                p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                                                p.svdfunc_UserSaveID == userid2 && p.del != true &&


                                                p.tbMoalefeDastmozdiValueFromExcel
                                                                                                      .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();

                                //if (sabttttt != null)
                                //{
                                //    savedfunctions = sabttttt;
                                //}
                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }

                                }
                                if (sabttttt == null)
                                {
                                    savedfunctions.FK_tbmarahesabt4 = id;
                                    savedfunctions.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    savedfunctions.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                    savedfunctions.svdfunc_pymnID = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    savedfunctions.svdfunc_UserSaveID = userid2;
                                    savedfunctions.svdfunc_IsSubmmit = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        savedfunctions.Final_Accept = true;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        savedfunctions.Final_Accept = false;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else
                                    {
                                        savedfunctions.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    savedfunctions.svdfunc_FileNameExcel = filename;
                                    savedfunctions.svdfunc_FileSystemNameExcel = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions2.Add(savedfunctions);
                                    if (savedfunctionRepo.Create(savedfunctions) == "True")
                                    {
                                    }
                                    else
                                    {
                                        savedfunctionRepo.Create5(savedfunctions2);
                                    }
                                }
                                else if (sabttttt != null)
                                {
                                    sabttttt.FK_tbmarahesabt4 = id;
                                    sabttttt.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    sabttttt.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    sabttttt.svdfunc_SavedDateTime = DateTime.Now;
                                    sabttttt.svdfunc_pymnID = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    sabttttt.svdfunc_UserSaveID = userid2;
                                    sabttttt.svdfunc_IsSubmmit = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        sabttttt.Final_Accept = true;
                                        sabttttt.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        sabttttt.Final_Accept = false;
                                        sabttttt.For_Accept = true;

                                    }
                                    else
                                    {
                                        sabttttt.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    sabttttt.svdfunc_FileNameExcel = filename;
                                    sabttttt.svdfunc_FileSystemNameExcel = files.FileName;
                                    await db.SaveChangesAsync();
                                    savedfunctions = sabttttt;
                                }
                                else
                                {
                                    savedfunctions.FK_tbmarahesabt4 = id;
                                    savedfunctions.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    savedfunctions.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                    savedfunctions.svdfunc_pymnID = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    savedfunctions.svdfunc_UserSaveID = userid2;
                                    savedfunctions.svdfunc_IsSubmmit = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        savedfunctions.Final_Accept = true;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else if (vaziat == 2 && vaziat2 != naheii)
                                    {
                                        savedfunctions.Final_Accept = false;
                                        savedfunctions.For_Accept = true;

                                    }
                                    else
                                    {
                                        savedfunctions.Final_Accept = false;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    savedfunctions.svdfunc_FileNameExcel = filename;
                                    savedfunctions.svdfunc_FileSystemNameExcel = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions2.Add(savedfunctions);
                                    if (savedfunctionRepo.Create(savedfunctions) == "True")
                                    {
                                    }
                                    else
                                    {
                                        savedfunctionRepo.Create5(savedfunctions2);
                                    }
                                }
                                var count2 = workbook.Worksheets[0].Rows.Count();
                                var usr12334 = db.tbUsers.ToList();

                                List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                                lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                                for (int i = 3; i < count2; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }
                                    int UserID = 0;
                                    if(personalid!= "")
                                    {
                                        var perID = System.Convert.ToInt32(personalid);
                                        var us11 = usr12334.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault();
                                        if (us11 != null)
                                        {
                                            var usrlistt = tbAdamAdam.Where(p => p.nam == us11.usr_ID).FirstOrDefault();
                                            if (usrlistt != null)
                                            {
                                                UserID = us11.usr_ID;
                                                int shomarande = 3;

                                                var taeed23 = false;

                                                foreach (var item in lstMoalefeint)
                                                {
                                                    var findmoalf = tbmarahlsabtbastedit6.Where(p => p.FK_moalfeh2 == item).FirstOrDefault();
                                                    var fi = item;
                                                    if (findmoalf != null)
                                                    {
                                                        fi = (int)findmoalf.FK_moalfe;
                                                    }
                                                    var Value = row.Cells[shomarande];
                                                    if (Value.Value != null)
                                                    {
                                                        if (Value.Value.ToString() != "")
                                                        {
                                                            var MoalefeID = 0;
                                                            //lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault();
                                                            //lstmoalafe2.Where(p => p.md_ID==item).FirstOrDefault();

                                                            if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                                                            {
                                                                taeed23 = true;
                                                                //sabttttt.Final_Accept = true;
                                                                //sabttttt.For_Accept = true;

                                                            }
                                                            else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                                                            {
                                                                //sabttttt.Final_Accept = false;
                                                                //sabttttt.For_Accept = true;

                                                            }
                                                            else
                                                            {
                                                                //sabttttt.Final_Accept = false;

                                                            }
                                                            if (savedfunctions.Final_Accept == true)
                                                            {
                                                                taeed23 = true;


                                                            }
                                                            value = Value.Value.ToString();
                                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                            {
                                                                MoalfeVal_FKMoalafeDastmozdi = fi,
                                                                MoalfeVal_FKUser = UserID,
                                                                MoalfeVal_Month = System.Convert.ToInt32(month),
                                                                MoalfeVal_Year = System.Convert.ToInt32(year),
                                                                MoalfeVal_Value = System.Convert.ToDouble(value),
                                                                FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                                Final_accept = taeed23,


                                                            };
                                                            lstMoalefeexcel.Add(dastmozdexcel);
                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}


                                                        }

                                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                        {
                                                            if (savedfunctions.Final_Accept == true)
                                                            {
                                                                taeed23 = true;


                                                            }
                                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                            {
                                                                //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                                MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID == fi).Select(p => p.md_ID).FirstOrDefault(),

                                                                MoalfeVal_FKUser = UserID,
                                                                MoalfeVal_Month = System.Convert.ToInt32(month),
                                                                MoalfeVal_Year = System.Convert.ToInt32(year),
                                                                MoalfeVal_Value = 0,
                                                                Final_accept = taeed23,

                                                                FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                            };
                                                            lstMoalefeexcel.Add(dastmozdexcel);

                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}
                                                        }
                                                    }

                                                    else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                    {
                                                        if (savedfunctions.Final_Accept == true)
                                                        {
                                                            taeed23 = true;


                                                        }

                                                        tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                        {
                                                            //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                            MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID == fi).Select(p => p.md_ID).FirstOrDefault(),
                                                            Final_accept = taeed23,

                                                            MoalfeVal_FKUser = UserID,
                                                            MoalfeVal_Month = System.Convert.ToInt32(month),
                                                            MoalfeVal_Year = System.Convert.ToInt32(year),
                                                            MoalfeVal_Value = 0,
                                                            FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                        };
                                                        lstMoalefeexcel.Add(dastmozdexcel);

                                                        //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                        //{
                                                        //}
                                                        //else
                                                        //{
                                                        //    transaction.Rollback();
                                                        //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                        //}
                                                    }

                                                    shomarande++;
                                                }
                                            }

                                       
                                        }
                                    }
                               
                                 



                                }
                            if (sabttttt != null)
                            {
                                string sql = "DELETE FROM [Salary].[tbMoalefeDastmozdiValueFromExcel] WHERE FK_SavedFunctionsID = @id";
                                var param = new SqlParameter("@id", sabttttt.svdfunc_ID);

                                await db.Database.ExecuteSqlCommandAsync(sql, param);
                            }
                                   
                                        //transaction2.Commit();

                                    
                                
                                
                                if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.BulkInsertDataAsync(lstMoalefeexcel) > 0)
                                    {

                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }
                                //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();








                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        //transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        //transaction.Rollback();
                        return ex.Message;
                    }
                }

            //}
        }
        public async Task<string> GetDataFromExcelMachinsOrTools4(HttpPostedFileBase MyExcelStream,         int id, int Month, int Year)
        {
            int usr = 0;
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User2 = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User2 != null)
                    {
                        usr = User2.usr_ID;
                    }
                }
            }

            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id&&p.del!=true).FirstOrDefault();

            var pyy = tbmarahesabt4.tbmarahlsabtbastedit2.tbPeymanContracts ;
            //var tool = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == Type).FirstOrDefault();
            var pymanname = pyy.pec_Title;
            //var tools = tool.Eqpgrp_Name;
            //var ty = db.tbmarahelsabt.Where(p => p.FK_Group == Type).Select(s => s.ID).FirstOrDefault();
     //       var us = db.tbReffrenceSave
     //           .Where(p => p.FK_PeymanID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn&& p.tbPeymanContracts.Inactive != true && p.IsFor == ty)


     //            .OrderByDescending(p => p.ID)
     //            .FirstOrDefault();
     //       var us2 = db.tbReffrenceSaveLevel
     //.Where(p => p.FK_RRSave == us.ID)
     //.OrderByDescending(p => p.ID)
     //.FirstOrDefault();

     //       var rus3 = db.tbReffrenceSaveLevelUser
     //     .Where(p => p.FK_LevelID == us2.ID)
     //     .OrderByDescending(p => p.ID)
     //     .Select(p => p.FK_UserID)
     //     .FirstOrDefault();
            var User = new tbUsers();
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
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
            var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
            var sabt = tbmarhesabt3.marahelsabt;
            var taeed = tbmarhesabt3.taeed;

            int vaziat = 0;
            int vaziat2 = 0;

            if (tbmarahesabt4.sabt == true)
            {
                vaziat = 1;
                vaziat2 = (int)tbmarahesabt4.number;

            }
            else if (tbmarahesabt4.taeed == true)
            {
                vaziat = 2;
                vaziat2 = (int)tbmarahesabt4.number;
            }
            var sabtt = db.tbEquipmentMoalefeValueReffrenceSave
.Where(p => p.FK_tbmarahesabt4 == id && p.del != true&&

          p.FK_User == userid &&

          p.tbEquipmentMoalefeValue
              .Any(m => m.Year == Year && m.Month == Month))
.FirstOrDefault();
            //var findmarhaleh=db.tb

            var find = await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_tbmarahesabt4 == id && p.Final_Accept == true &&p.del!=true&& p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefaultAsync();
            if (find != null && vaziat2 != tbmarhesabt3.taeed && vaziat != 2)
            {
                return "این بسته در این ماه و سال  تایید نهایی شده است ";
            }

            else if (find != null && vaziat2 != tbmarhesabt3.taeed && vaziat == 2)
            {
                return "این بسته در این ماه و سال  تایید نهایی شده است ";
            }
            List<string> lstMoalefe = new List<string>();
            List<int> intt = new List<int>();

            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;
            tbEquipmentSpecificationData obj2 = new tbEquipmentSpecificationData();
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();
            List<tbEquipmentMoalefeValue> list = new List<tbEquipmentMoalefeValue>();

            tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
            if (sabtt == null)
            {
                objsave.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;

                objsave.FK_User = userid;
                objsave.DateTime = DateTime.Now;
                objsave.FK_tbmarahesabt4 = id;
                objsave.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;
                objsave.FK_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                {
                    objsave.Final_Accept = true;
                    objsave.For_Acceot = true;

                }
                else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                {
                    objsave.Final_Accept = false;
                    objsave.For_Acceot = true;

                }
                else
                {
                    objsave.Final_Accept = false;
                    objsave.For_Acceot = false;
                }
                //var firstFilter = Filters.First();
                //var fkBASTE = firstFilter.FK_BASTE;

                //            var rus33 = await db.tbReffrenceSaveLevelUser
                //.Where(p => p.FK_LevelID == FK_BASTE)
                //.OrderByDescending(p => p.ID)
                //.Select(p => p.FK_UserID)
                //.FirstOrDefaultAsync();
                //            if (userid == rus3)
                //            {
                //                objsave.Final_Sabt = true;
                //            }

                if (MyExcelStream != null)
                {

                    if (MyExcelStream.ContentLength > 0)
                    {
                        var segment = MyExcelStream.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                        objsave.File_SystemNameForAccept = filename;
                        objsave.FileNameForAccept = MyExcelStream.FileName;
                    }
                }



                db.tbEquipmentMoalefeValueReffrenceSave.Add(objsave);
                db.SaveChanges();

            }

         else   if (sabtt != null)
            {
                sabtt.FK_Basteh = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;

                sabtt.FK_User = userid;
                sabtt.DateTime = DateTime.Now;
                sabtt.FK_tbmarahesabt4 = id;
                sabtt.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;
                sabtt.FK_pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                if (vaziat == 2 && vaziat2 == tbmarhesabt3.taeed)
                {
                    sabtt.Final_Accept = true;
                    sabtt.For_Acceot = true;

                }
                else if (vaziat == 2 && vaziat2 != tbmarhesabt3.taeed)
                {
                    sabtt.Final_Accept = false;
                    sabtt.For_Acceot = true;

                }
                else
                {
                    sabtt.Final_Accept = false;
                    sabtt.For_Acceot = false;
                }
                //var firstFilter = Filters.First();
                //var fkBASTE = firstFilter.FK_BASTE;

                //            var rus33 = await db.tbReffrenceSaveLevelUser
                //.Where(p => p.FK_LevelID == FK_BASTE)
                //.OrderByDescending(p => p.ID)
                //.Select(p => p.FK_UserID)
                //.FirstOrDefaultAsync();
                //            if (userid == rus3)
                //            {
                //                objsave.Final_Sabt = true;
                //            }

                if (MyExcelStream != null)
                {

                    if (MyExcelStream.ContentLength > 0)
                    {
                        var segment = MyExcelStream.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                        sabtt.File_SystemNameForAccept = filename;
                        sabtt.FileNameForAccept = MyExcelStream.FileName;
                    }
                }



                db.SaveChanges();
                objsave = sabtt;
            }

            //var tbAdamAdam = db.tbAdamAdam.Where(s => s.Name == userid &&s.Value==1).ToList();

            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();


                            if (count == 1)
                            {
                                return "فایل اکسل فاقد اطلاعات می باشد";
                            }

                            for (int i = 2; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }



                            var user = await db.tbUsers.Where(p => p.usr_Personal_ID != null).ToListAsync();
                            var pymn = await db.tbPeymanContracts.Where(p => p.Inactive != true).ToListAsync();
                            List<tbEquipmentMoalefeValue> objectsToAdd = new List<tbEquipmentMoalefeValue>();

                            for (int i = 3; i < count; i++)
                            {
                                tbEquipmentMoalefeValue list2 = new tbEquipmentMoalefeValue();

                                var row = workbook.Worksheets[0].Rows[i];
                                var Name = row.Cells[1];


                                //int savedItemId = db.tbEquipments
                                //                    .OrderByDescending(p => p.ID)
                                //                    .Select(s => s.ID)
                                //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                                int shomarande = 2;

                                foreach (var it in lstMoalefe)
                                {
                                    tbEquipmentMoalefeValue obj22 = new tbEquipmentMoalefeValue();
                                    if (int.TryParse(Name.Value.ToString(), out int vahed2))
                                    {
                                        var find2 = pymn.Where(p => p.pec_ProjectCode == vahed2).FirstOrDefault();
                                        var findusr = user.Where(p => p.usr_Personal_ID == vahed2).FirstOrDefault();

                                        if (find2 != null)
                                        {
                                            obj22.Fk_pymn = find2.pec_ID;
                                        }
                                        else
                                        {
                                            obj22.Fk_user = findusr.usr_ID;
                                        }
                                    }
                                    else
                                    {
                                        if (sabtt == null)
                                        {
                                            db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                                            await db.SaveChangesAsync();
                                        }
                                    
                                        return " فایل را دوباره بررسی بفرمایید ";
                                    }

                                    var Value = row.Cells[shomarande];
                                    var value = Value.Value ?? null;
                                    float? countDays;

                                    if (float.TryParse(value, out float parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }
                                    else
                                    {
                                        countDays = null; // Or handle the case where the value cannot be parsed
                                    }

                                    int? FK_Equipment = null;

                                    if (int.TryParse(it, out int FK_Equipment2))
                                    {
                                        FK_Equipment = FK_Equipment2;
                                    }
                                    else
                                    {
                                        FK_Equipment = null; // Or handle the case where the value cannot be parsed
                                    }

                                    obj22.CountDays = countDays;
                                    obj22.FK_Equipment = FK_Equipment;
                                    obj22.FK_tbEquipmentMoalefeValueReffrenceSave = objsave.ID;
                                    obj22.Month = Month;
                                    obj22.Year = Year;
                                    if (obj22.Fk_user != null)
                                    {
                                        
                                        var usrlistt = tbAdamAdam.Where(s => s.nam == obj22.Fk_user&&s.Value==1).FirstOrDefault();
                                        if (usrlistt != null)
                                        {
                                            objectsToAdd.Add(obj22);

                                        }

                                    }
                                    else
                                    {
                                        objectsToAdd.Add(obj22);

                                    }

                                    shomarande++;
                                }

                                // Add all objects in the list to the database using AddRange



                            }
                            if (sabtt != null)
                            {
                                var listtbEquipmentMoalefeValue = await db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == sabtt.ID).ToListAsync();
                                db.tbEquipmentMoalefeValue.RemoveRange(listtbEquipmentMoalefeValue);
                                await db.SaveChangesAsync();
                            }
                            db.tbEquipmentMoalefeValue.AddRange(objectsToAdd);
                            await db.SaveChangesAsync();
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (sabtt == null)
                    {
                        db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                        await db.SaveChangesAsync();
                    }
                    return ex.Message;
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return "True";









        }
        public async Task<string> GetDataFromExcel_Moalefe_Fosabtsabttjadidrialy(HttpPostedFileBase files, int id, int month2, int year2)
        {
            string name, family, personalid, year, month = "";
            string value;
            int usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                    }
                }
            }

            //var idpy = System.Convert.ToString(PeymanID);
            //var pymn = idpy.Split(',');
            //List<int?> listpymn = new List<int?>();
            //for (int i = 0; i < pymn.Count(); i++)
            //{
            //    listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //}
            var tbAdamAdam = db.tbAdamAdam.Where(p => p.Name == usr && p.Value == 1).ToList();
            var tbmarahesabt4 = db.tbmarahesabt4.Where(p => p.ID == id&&p.del!=true).FirstOrDefault();

            //var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).ToList();
            List<int> UserListint = tbAdamAdam.Select(p => p.nam ?? 0).ToList();


            //foreach (var item in usrID)
            //{
            //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    UserList.Add(user);
            //}

            List<int> lstMoalefeint = new List<int>();

            //int number = PeymanID;
            //var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<tbMoalefeValueFish> lstMoalefeexcel = new List<tbMoalefeValueFish>();
            //    var us4 = db.tbReffrenceSave
            //.Where(p => p.FK_PeymanID == PeymanID)
            //.SelectMany(s => s.tbReffrenceSaveLevel)
            //.Where(s => s.ID == Basteh)
            //.OrderByDescending(p => p.ID)
            //.FirstOrDefault();
            //var us2 = db.tbReffrenceSaveLevel
            //    .Where(p => p.FK_RRSave == us4.FK_RRSave)
            //    .OrderByDescending(p => p.ID)
            //    .FirstOrDefault();
            //var rus3 = db.tbReffrenceSaveLevelUser
            //    .Where(p => p.FK_LevelID == us2.ID)
            //    .OrderByDescending(p => p.ID)
            //    .Select(p => p.FK_UserID)
            //    .FirstOrDefault();
            var User2 = new tbUsers();
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                //if (User2 != null)
                //{

                //    userid = User2.usr_ID;
                //    if (userid == rus3)
                //    {

                //        var matchedRows = db.tbMoalefeValuePishkhan
                //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                //            .ToList();

                //        foreach (var row1 in matchedRows)
                //        {
                //            row1.mlfval_Value = "1";

                //        }

                //        db.SaveChanges();

                //    }
                //}

            }


        
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {


                        var tbmarahesabt41111 = db.tbmarahesabt4.Where(p => p.del != true).ToList();


                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year);
                                var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == id &&
p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
p.tbMoalefeDastmozdiValueFromExcel
.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();
                                tbmoalfefishexcel savedfunctions = new tbmoalfefishexcel();
                                List<tbmoalfefishexcel> savedfunctions2 = new List<tbmoalfefishexcel>();


                                int userid2 = 0;
                                var cookie_user222 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user222 != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user222.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }
                                if (cookie_user2 != null)
                                {
                                    //int userid = 0;
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                                    }
                                    if (User2 != null)
                                    {
                                        if (tbmarahesabt4.sabt == true)
                                        {
                                            var findusrend = db.tbmarahesabt4.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true && p.del != true).OrderByDescending(s => s.number).FirstOrDefault();
                                            if (findusrend != null)
                                            {
                                                if (findusrend.ID == tbmarahesabt4.ID)
                                                {
                                                    //savedfunctions.Final_Sabt = true;

                                                }

                                            }

                                        }
                                        //int userid4 = User2.usr_ID;
                                        //if (userid4 == rus3)
                                        //{
                                        //    savedfunctions.Final_Sabt = true;



                                        //}
                                    }
                                }
                                var month22 = System.Convert.ToInt32(month);
                                var year22 = System.Convert.ToInt32(year);
                                //}
                                //else//add
                                //{
                                var maxNumber = tbmarahesabt41111
       .Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true)
       .Max(p => (int?)p.number);

                                var naheii = 0;

                                if (maxNumber.HasValue)
                                {
                                    naheii = maxNumber.Value; // استفاده از Value برای دسترسی به مقدار
                                }
                                var tbmarhesabt3 = db.tbmarhesabt3.FirstOrDefault();
                                var sabt = tbmarhesabt3.marahelsabt;
                                var taeed = tbmarhesabt3.taeed;

                                int vaziat = 0;
                                int vaziat2 = 0;

                                if (tbmarahesabt4.sabt == true)
                                {
                                    vaziat = 1;
                                    vaziat2 = (int)tbmarahesabt4.number;

                                }
                                else if (tbmarahesabt4.taeed == true)
                                {
                                    vaziat = 2;
                                    vaziat2 = (int)tbmarahesabt4.number;
                                }
                                var find = db.tbmoalfefishexcel
    .Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City &&
    p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST &&
                p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                p.Final_accept == 1 && p.del != true &&
                p.tbMoalefeValueFish
                                                    .Any(m => m.mlfvlfsh_Year == year2 && m.mlfvlfsh_Month == month2)).FirstOrDefault();
                                var findcontrol = db.tbmoalfefishexcel
.Where(p => p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City 
 &&p.del!=true&&
        p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
        p.FK_control != null &&
        p.tbMoalefeValueFish
                                            .Any(m => m.mlfvlfsh_Year == year2 && m.mlfvlfsh_Month == month2)).FirstOrDefault();
                                if (vaziat == 2)
                                {
                                    if (find != null && vaziat2 != naheii)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
                                else
                                {
                                    if (find != null)
                                    {
                                        return " این بسته ثبت و تایید نهایی  شده است ";
                                    }
                                }
                                //if (findcontrol != null)
                                //{
                                //    return " این بسته در مرحله کنترل   شده است ";
                                //}
                                var sabttttt = db.tbmoalfefishexcel
                                    .Where(p => p.FK_tbmarahesabt4 == id &&
                                                p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn &&
                                                p.usr_sabt == userid2 &&

                                                p.tbMoalefeValueFish
                                                                                                      .Any(m => m.mlfvlfsh_Year == year2 && m.mlfvlfsh_Month == month2)).FirstOrDefault();

                             
                               
                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }

                                }

                                if (sabttttt != null)
                                {
                                    sabttttt.FK_tbmarahesabt4 = id;
                                    sabttttt.FK_Basteh1 = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    sabttttt.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    sabttttt.Datatmie = DateTime.Now;
                                    sabttttt.Fk_Pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    sabttttt.usr_sabt = userid2;
                                    sabttttt.submitt = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        sabttttt.Final_accept = 1;
                                        

                                    }
                                    else
                                    {
                                        sabttttt.Final_accept = 0;

                                    }





                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    sabttttt.FileName = filename;
                                    sabttttt.Filsesystem = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions = sabttttt;
                                }
                                else if (sabttttt == null)
                                {
                                    savedfunctions.FK_tbmarahesabt4 = id;
                                    savedfunctions.FK_Basteh1 = tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST;
                                    savedfunctions.city = tbmarahesabt4.tbmarahlsabtbastedit2.FK_City;

                                    ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                    ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                    savedfunctions.Datatmie = DateTime.Now;
                                    savedfunctions.Fk_Pymn = tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn;
                                    savedfunctions.usr_sabt = userid2;
                                    savedfunctions.submitt = false;
                                    //savedfunctions.For_Accept = true;
                                    if (vaziat == 2 && vaziat2 == naheii)
                                    {
                                        savedfunctions.Final_accept = 1;


                                    }
                                    else
                                    {
                                        savedfunctions.Final_accept = 0;

                                    }




                                    var segment = files.FileName.Split('.');
                                    string file_type = segment[segment.Length - 1];
                                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                    savedfunctions.FileName = filename;
                                    savedfunctions.Filsesystem = files.FileName;








                                    //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                    savedfunctions2.Add(savedfunctions);
                                    if (savedfunctionRepo.Createtoal(savedfunctions) == "True")
                                    {
                                    }
                                    else
                                    {
                                        savedfunctionRepo.Create5tial(savedfunctions2);
                                    }
                                }
                                var tbmarahlsabtbastedit6 = db.tbmarahlsabtbastedit6.ToList();

                                List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                                lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                                var count2 = workbook.Worksheets[0].Rows.Count();

                                for (int i = 3; i < count2 ; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }
                                    var usr12334 = db.tbUsers.ToList();

                                    var perID = System.Convert.ToInt32(personalid);
                                    int UserID = 0;
                                    if (personalid != "")
                                    {
                                        var usrr = usr12334.Where(p =>  p.usr_Personal_ID == perID).FirstOrDefault();
                                        if (usrr != null)
                                        {
                                            var usrlistt = tbAdamAdam.Where(p => p.nam == usrr.usr_ID).FirstOrDefault();
                                            if (usrlistt != null)
                                            {
                                                UserID = usrr.usr_ID;
                                                int shomarande = 3;


                                                foreach (var item in lstMoalefeint)
                                                {
                                                    var findmoalf = tbmarahlsabtbastedit6.Where(p => p.FK_moalfeh2 == item).FirstOrDefault();
                                                    var fi = item;
                                                    if (findmoalf != null)
                                                    {
                                                        fi = (int)findmoalf.FK_moalfe;
                                                    }
                                                    var Value = row.Cells[shomarande];
                                                    if (Value.Value != null)
                                                    {
                                                        if (Value.Value.ToString() != "")
                                                        {
                                                            var vaz = false;
                                                            if (vaziat == 2 && vaziat2 == naheii)
                                                            {
                                                                vaz = true;

                                                            }
                                                            var MoalefeID =
                                                            //lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault();
                                                            //lstmoalafe2.Where(p => p.md_ID==item).FirstOrDefault();



                                                            value = Value.Value.ToString();
                                                            tbMoalefeValueFish dastmozdexcel = new tbMoalefeValueFish
                                                            {
                                                                Finalaccept = vaz,
                                                                FK_Moalefe = fi,
                                                                FK_User = UserID,
                                                                mlfvlfsh_Month = System.Convert.ToInt32(month),
                                                                mlfvlfsh_Year = System.Convert.ToInt32(year),
                                                                mlfvlfsh_Value = System.Convert.ToDouble(value),
                                                                FK_EXCel = savedfunctions.ID,
                                                                Datatmie = DateTime.Now,


                                                            };
                                                            lstMoalefeexcel.Add(dastmozdexcel);
                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}


                                                        }

                                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                        {
                                                            var vaz = false;
                                                            if (vaziat == 2 && vaziat2 == naheii)
                                                            {
                                                                vaz = true;

                                                            }

                                                            tbMoalefeValueFish dastmozdexcel = new tbMoalefeValueFish
                                                            {
                                                                Finalaccept = vaz,
                                                                Datatmie = DateTime.Now,

                                                                FK_Moalefe = fi,
                                                                FK_User = UserID,
                                                                mlfvlfsh_Month = System.Convert.ToInt32(month),
                                                                mlfvlfsh_Year = System.Convert.ToInt32(year),
                                                                mlfvlfsh_Value = 0,
                                                                FK_EXCel = savedfunctions.ID,
                                                                //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                                //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID == item).Select(p => p.md_ID).FirstOrDefault(),

                                                                //MoalfeVal_FKUser = UserID,
                                                                //MoalfeVal_Month = System.Convert.ToInt32(month),
                                                                //MoalfeVal_Year = System.Convert.ToInt32(year),
                                                                //MoalfeVal_Value = 0,

                                                                //FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                            };
                                                            lstMoalefeexcel.Add(dastmozdexcel);

                                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                            //{
                                                            //}
                                                            //else
                                                            //{
                                                            //    transaction.Rollback();
                                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                            //}
                                                        }
                                                    }

                                                    else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                    {
                                                        var vaz = false;
                                                        if (vaziat == 2 && vaziat2 == naheii)
                                                        {
                                                            vaz = true;

                                                        }
                                                        tbMoalefeValueFish dastmozdexcel = new tbMoalefeValueFish
                                                        {
                                                            Finalaccept = vaz,
                                                            Datatmie = DateTime.Now,

                                                            FK_Moalefe = fi,
                                                            FK_User = UserID,
                                                            mlfvlfsh_Month = System.Convert.ToInt32(month),
                                                            mlfvlfsh_Year = System.Convert.ToInt32(year),
                                                            mlfvlfsh_Value = 0,
                                                            FK_EXCel = savedfunctions.ID,
                                                            //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                            //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID == item).Select(p => p.md_ID).FirstOrDefault(),

                                                            //MoalfeVal_FKUser = UserID,
                                                            //MoalfeVal_Month = System.Convert.ToInt32(month),
                                                            //MoalfeVal_Year = System.Convert.ToInt32(year),
                                                            //MoalfeVal_Value = 0,

                                                            //FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                        };
                                                        //tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                        //{
                                                        //    //MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_Title.Contains(item)).Select(p => p.md_ID).FirstOrDefault(),
                                                        //    MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID == item).Select(p => p.md_ID).FirstOrDefault(),

                                                        //    MoalfeVal_FKUser = UserID,
                                                        //    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                        //    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                        //    MoalfeVal_Value = 0,
                                                        //    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                        //};
                                                        lstMoalefeexcel.Add(dastmozdexcel);

                                                        //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                        //{
                                                        //}
                                                        //else
                                                        //{
                                                        //    transaction.Rollback();
                                                        //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                        //}
                                                    }

                                                    shomarande++;
                                                }
                                            }
                                          
                                        }
                           

                                    }





                                }
                            
                                    if (sabttttt != null)
                                    {
                                    if (sabttttt != null)
                                    {
                                        string sql = "DELETE FROM [Salary].[tbMoalefeValueFish] WHERE FK_EXCel = @id";
                                        var param = new SqlParameter("@id", sabttttt.ID);

                                        await db.Database.ExecuteSqlCommandAsync(sql, param);
                                    }
                                    //var recordsToDelete = db.tbMoalefeValueFish
                                    //        .Where(p => p.FK_EXCel == sabttttt.ID);

                                    //    db.tbMoalefeValueFish.RemoveRange(recordsToDelete);
                                    //    await db.SaveChangesAsync();
                                    }
                                
                                if (moalefeexcelRepo.Createrial(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.BulkInsertDataAsyncfish(lstMoalefeexcel) > 0)
                                    {
                                        //await moalefeexcelRepo.Create2Asyncial(lstMoalefeexcel) > 0
                                        // Saving successful, number of saved items is returned by Create2Async
                                        // You can use the returned value here, for example:
                                        //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                                        //Console.WriteLine($"Successfully saved {savedCount} items.");
                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }
                                //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();








                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        //transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        //transaction.Rollback();
                        return ex.Message;
                    }
                }

            
        }

        public async Task<string> GetDataFromExcel_Moalefe_Fosabtedit(HttpPostedFileBase files, int Basteh)
        {
            string name, family, personalid, year, month = "";
            string value;


            //    var idpy = System.Convert.ToString(PeymanID);
            //    var pymn = idpy.Split(',');
            //    List<int?> listpymn = new List<int?>();
            //    for (int i = 0; i < pymn.Count(); i++)
            //    {
            //        listpymn.Add(System.Convert.ToInt32(pymn[i]));
            //    }
            //    var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).ToList();
            //    List<tbUsers> UserList = new List<tbUsers>();
            //    UserList = usrID.Select(p => p.tbUsers).ToList();
            //    //foreach (var item in usrID)
            //    //{
            //    //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    //    UserList.Add(user);
            //    //}

            //int numberOfUsers = UserList.Count + 3;

            //    int number = PeymanID;
            //    var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefault();
            List<string> lstMoalefe = new List<string>();
            List<int> lstMoalefeint = new List<int>();

            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
            //    var us4 = db.tbReffrenceSave
            //.Where(p => p.FK_PeymanID == PeymanID)
            //.SelectMany(s => s.tbReffrenceSaveLevel)
            //.Where(s => s.ID == Basteh)
            //.OrderByDescending(p => p.ID)
            //.FirstOrDefault();
            //    var us2 = db.tbReffrenceSaveLevel
            //        .Where(p => p.FK_RRSave == us4.FK_RRSave)
            //        .OrderByDescending(p => p.ID)
            //        .FirstOrDefault();
            //    var rus3 = db.tbReffrenceSaveLevelUser
            //        .Where(p => p.FK_LevelID == us2.ID)
            //        .OrderByDescending(p => p.ID)
            //        .Select(p => p.FK_UserID)
            //        .FirstOrDefault();
            //    var User2 = new tbUsers();
            //    var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            //    if (cookie_user2 != null)
            //    {
            //        //int userid = 0;
            //        var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
            //        using (SaabEntities db = new SaabEntities())
            //        {
            //            User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
            //        }
            //        //if (User2 != null)
            //        //{

            //        //    userid = User2.usr_ID;
            //        //    if (userid == rus3)
            //        //    {

            //        //        var matchedRows = db.tbMoalefeValuePishkhan
            //        //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
            //        //            .ToList();

            //        //        foreach (var row1 in matchedRows)
            //        //        {
            //        //            row1.mlfval_Value = "1";

            //        //        }

            //        //        db.SaveChanges();

            //        //    }
            //        //}

            //    }


            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {




                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year); 
                                //.Where(p => p.svdfunc_BastehID == Basteh &&
                                //p.svdfunc_pymnID == PeymanID &&
                                //p.tbMoalefeDastmozdiValueFromExcel
                                //.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
                                //.FirstOrDefault();
                                //                            tbSavedFunctions savedfunctions = new tbSavedFunctions();
                                //                            List<tbSavedFunctions> savedfunctions2 = new List<tbSavedFunctions>();


                                //                            int userid2 = 0;
                                //                            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                //                            if (cookie_user != null)
                                //                            {
                                //                                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                //                                using (SaabEntities db = new SaabEntities())
                                //                                {
                                //                                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                //                                    if (User != null)
                                //                                    {
                                //                                        userid2 = User.usr_ID;
                                //                                    }
                                //                                }
                                //                            }
                                //                            if (cookie_user2 != null)
                                //                            {
                                //                                //int userid = 0;
                                //                                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                                //                                using (SaabEntities db = new SaabEntities())
                                //                                {
                                //                                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                                //                                }
                                //                                if (User2 != null)
                                //                                {

                                //                                    int userid4 = User2.usr_ID;
                                //                                    if (userid4 == rus3)
                                //                                    {
                                //                                        savedfunctions.Final_Sabt = true;



                                //                                    }
                                //                                }
                                //                            }
                                //                            var month2 = System.Convert.ToInt32(month);
                                //                            var year2 = System.Convert.ToInt32(year);
                                //                            //}
                                //                            //else//add
                                //                            //{
                                //                            var find = db.tbSavedFunctions
                                //.Where(p => p.svdfunc_BastehID == Basteh &&
                                //            p.svdfunc_pymnID == number &&
                                //            p.Final_Accept == true &&
                                //            p.tbMoalefeDastmozdiValueFromExcel
                                //                                                .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();

                                //                            if (find != null)
                                //                            {
                                //                                return " این بسته ثبت و تایید شده است ";
                                //                            }
                                //                            var sabttttt = db.tbSavedFunctions
                                //                                .Where(p => p.svdfunc_BastehID == Basteh &&
                                //                                            p.svdfunc_pymnID == number &&
                                //                                            p.svdfunc_UserSaveID == userid2 &&

                                //                                            p.tbMoalefeDastmozdiValueFromExcel
                                //                                                                                                  .Any(m => m.MoalfeVal_Year == year2 && m.MoalfeVal_Month == month2)).FirstOrDefault();

                                //                            if (sabttttt != null)
                                //                            {
                                //                                return "شما قبلا این بسته را در این سال و ماه ثبت کرده اید ";
                                //                            }
                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lstMoalefeint.Add(cellValue);
                                    }

                                }
                                var tb = await db.tbSavedFunctions.Where(p => p.svdfunc_ID == Basteh).FirstOrDefaultAsync();
                                var pymn = await db.tbUsers.Where(p => p.Link_User_And_Peyman.Any(s => s.FK_Peyman_ID == tb.svdfunc_pymnID && s.Status == true)).ToListAsync();

                                // استفاده از EFCore.BulkExtensions برای حذف دسته‌ای
                                var delet = await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == tb.svdfunc_ID).ToListAsync();
                                //await db.BulkDelete(delet);

                                //await db.SaveChangesAsync();

                                //var tb = db.tbSavedFunctions.Where(p => p.svdfunc_ID == Basteh).FirstOrDefault();
                                //var pymn = db.tbUsers.Where(p => p.Link_User_And_Peyman.Any(s => s.FK_Peyman_ID == tb.svdfunc_pymnID && s.Status == true)).ToList();
                                //var delet = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == tb.svdfunc_ID).ToList();
                                List<tbMoalefeDastmozdiValueFromExcel> del = new List<tbMoalefeDastmozdiValueFromExcel>();
                                //foreach (var it in delet)
                                //{
                                //    del.Add(it);
                                //}
                                db.tbMoalefeDastmozdiValueFromExcel.RemoveRange(delet);
                               await db.SaveChangesAsync();
                                //                            savedfunctions.svdfunc_BastehID = Basteh;
                                //                            ////savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                //                            ////savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                //                            savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                //                            savedfunctions.svdfunc_pymnID = PeymanID;
                                //                            savedfunctions.svdfunc_UserSaveID = userid2;
                                //                            savedfunctions.svdfunc_IsSubmmit = false;
                                //                            //savedfunctions.For_Accept = true;
                                //                            savedfunctions.Final_Accept = false;


                                int numberOfUsers = pymn.Count + 3;

                                var segment = files.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                tb.svdfunc_FileNameExcel = filename;
                                tb.svdfunc_FileSystemNameExcel = files.FileName;





                                await db.SaveChangesAsync();


                                //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;



                                List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                                lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID =
                                        db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).Select(p => p.usr_ID).FirstOrDefault();
                                    int shomarande = 3;


                                    foreach (var item in lstMoalefeint)
                                    {
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID =
                                                    lstmoalafe2.Where(p => p.md_ID==item).Select(p => p.md_ID).FirstOrDefault();


                                                value = Value.Value.ToString();
                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = System.Convert.ToDouble(value),
                                                    FK_SavedFunctionsID = tb.svdfunc_ID,



                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}


                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID==item).Select(p => p.md_ID).FirstOrDefault(),

                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0,

                                                    FK_SavedFunctionsID = tb.svdfunc_ID,
                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);

                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                            {
                                                MoalfeVal_FKMoalafeDastmozdi = lstmoalafe2.Where(p => p.md_ID==item).Select(p => p.md_ID).FirstOrDefault(),
                                                MoalfeVal_FKUser = UserID,
                                                MoalfeVal_Month = System.Convert.ToInt32(month),
                                                MoalfeVal_Year = System.Convert.ToInt32(year),
                                                MoalfeVal_Value = 0,
                                                FK_SavedFunctionsID = tb.svdfunc_ID,
                                            };
                                            lstMoalefeexcel.Add(dastmozdexcel);

                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                            //{
                                            //}
                                            //else
                                            //{
                                            //    transaction.Rollback();
                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                            //}
                                        }

                                        shomarande++;
                                    }



                                }
                                if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.Create2Async(lstMoalefeexcel) > 0)
                                    {
                                        // Saving successful, number of saved items is returned by Create2Async
                                        // You can use the returned value here, for example:
                                        //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                                        //Console.WriteLine($"Successfully saved {savedCount} items.");
                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }
                                //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();








                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }
        public async Task<ActionResult> ImportExcel_Moalefe_forexccept(HttpPostedFileBase files, int svdfunc_BastehID, int svdfunc_pymnID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_Foraccept(files, svdfunc_BastehID, svdfunc_pymnID);

            return Content(Message);
        }


        public async Task<ActionResult> ImportExcel_Moalefe_taeeadedit(HttpPostedFileBase files,  int svdfunc_BastehID)
        {

            var Message = "";
            if (files == null)
            {
                Message = "فایل بطور صحیح بارگذاری نشده است";
            }

            Message = await GetDataFromExcel_Moalefe_edit(files, svdfunc_BastehID);

            return Content(Message);
        }


        public async Task<string> GetDataFromExcel_Moalefe_edit(HttpPostedFileBase files,int svdfunc_BastehID)
        {
            string name, family, personalid, year, month = "";
            string value;
            var saneedit = db.tbSavedFunctions.Where(p => p.svdfunc_ID == svdfunc_BastehID).FirstOrDefault();
            int Basteh = (int )saneedit.svdfunc_BastehID;
            int PeymanID = (int)saneedit.svdfunc_pymnID;
            List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
            lstmoalafe2 = await db.tbContractMoalefeDastmozdi.ToListAsync();
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = await db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToListAsync();
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = await db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefaultAsync();
                UserList.Add(user);
            }

            int numberOfUsers = UserList.Count + 3;

            int number = PeymanID;
            var PymN = await db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefaultAsync();
            List<string> lstMoalefe = new List<string>();
            List<int> lisstintmoalf = new List<int>();

            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {




                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }

                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }

                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lisstintmoalf.Add(cellValue);
                                    }

                                }
                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year); var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == Basteh &&
p.svdfunc_pymnID == PeymanID &&
p.tbMoalefeDastmozdiValueFromExcel
.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();






                                var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();



                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }


                                //}
                                //else//add
                                //{
                                tbSavedFunctions savedfunctions = new tbSavedFunctions();

                                //savedfunctions.svdfunc_BastehID = Basteh;
                                //savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                //savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                //savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                //savedfunctions.svdfunc_pymnID = PeymanID;
                                //savedfunctions.svdfunc_UserSaveID = userid2;
                                //savedfunctions.svdfunc_IsSubmmit = false;
                                //savedfunctions.For_Accept = true;
                                //savedfunctions.Final_Accept = false;

                                //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;

                                var taeed =await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == svdfunc_BastehID).ToListAsync();
                                db.tbMoalefeDastmozdiValueFromExcel.RemoveRange(taeed);
                              await  db.SaveChangesAsync();


                                var segment = files.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                saneedit.svdfunc_FileNameExcel = filename;
                                saneedit.svdfunc_FileSystemNameExcel = files.FileName;
                              await  db.SaveChangesAsync();
                                //if (savedfunctionRepo.Create(savedfunctions) == "True")
                                //{
                                //}
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = await db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).Select(o => o.usr_ID).FirstOrDefaultAsync();
                                    int shomarande = 3;



                                    foreach (var item in lisstintmoalf)
                                    {
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID = item;


                                                value = Value.Value.ToString();
                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = System.Convert.ToDouble(value),
                                                    FK_SavedFunctionsID = saneedit.svdfunc_ID,


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}


                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = item,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0,

                                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                };
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                            {
                                                MoalfeVal_FKMoalafeDastmozdi = item,
                                                MoalfeVal_FKUser = UserID,
                                                MoalfeVal_Month = System.Convert.ToInt32(month),
                                                MoalfeVal_Year = System.Convert.ToInt32(year),
                                                MoalfeVal_Value = 0,
                                                FK_SavedFunctionsID = saneedit.svdfunc_ID,
                                            };
                                            //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                            //{
                                            //}
                                            //else
                                            //{
                                            //    transaction.Rollback();
                                            //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                            //}
                                        }

                                        shomarande++;
                                    }



                                }
                                if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.Create2Async(lstMoalefeexcel) > 0)
                                    {
                                        // Saving successful, number of saved items is returned by Create2Async
                                        // You can use the returned value here, for example:
                                        //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                                        //Console.WriteLine($"Successfully saved {savedCount} items.");
                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }


                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }

        public async Task<string> GetDataFromExcel_Moalefe_Foraccept(HttpPostedFileBase files, int Basteh, int PeymanID)
        {
            string name, family, personalid, year, month = "";
            string value;

            List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
            lstmoalafe2 =await db.tbContractMoalefeDastmozdi.ToListAsync();
            var idpy = System.Convert.ToString(PeymanID);
            var pymn = idpy.Split(',');
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID =await db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true).Select(p => p.FK_User_ID).ToListAsync();
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user =await db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefaultAsync();
                UserList.Add(user);
            }

            int numberOfUsers = UserList.Count + 3;

            int number = PeymanID;
            var PymN =await db.tbPeymanContracts.Where(p => p.pec_ID == number && p.Inactive != true).FirstOrDefaultAsync();
            List<string> lstMoalefe = new List<string>();
            List<int> lisstintmoalf = new List<int>();

            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {




                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(files.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var title2 = workbook.Worksheets[0].Rows[0].Cells;

                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }

                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString().TrimEnd('\n'));
                                }
                               
                                for (int i = 3; i < title2.Length; i++)
                                {
                                    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                                    {
                                        lisstintmoalf.Add(cellValue);
                                    }

                                }
                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }
                                var montht = System.Convert.ToInt32(month);
                                var yyer = System.Convert.ToInt32(year); var sabtt = db.tbSavedFunctions
.Where(p => p.svdfunc_BastehID == Basteh &&
p.svdfunc_pymnID == PeymanID &&
p.tbMoalefeDastmozdiValueFromExcel
.Any(m => m.MoalfeVal_Year == yyer && m.MoalfeVal_Month == montht)).OrderByDescending(sf => sf.svdfunc_ID)
.FirstOrDefault();






                                var exist = db.tbSavedFunctions.Where(p => p.svdfunc_ID == sabtt.svdfunc_ID).FirstOrDefault();



                                int userid2 = 0;
                                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                                if (cookie_user != null)
                                {
                                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                                    using (SaabEntities db = new SaabEntities())
                                    {
                                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                                        if (User != null)
                                        {
                                            userid2 = User.usr_ID;
                                        }
                                    }
                                }


                                //}
                                //else//add
                                //{
                                tbSavedFunctions savedfunctions = new tbSavedFunctions();

                                savedfunctions.svdfunc_BastehID = Basteh;
                                savedfunctions.svdfunc_FromDate = exist.svdfunc_FromDate;
                                savedfunctions.svdfunc_ToDate = exist.svdfunc_ToDate;
                                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                savedfunctions.svdfunc_pymnID = PeymanID;
                                savedfunctions.svdfunc_UserSaveID = userid2;
                                savedfunctions.svdfunc_IsSubmmit = false;
                                savedfunctions.For_Accept = true;
                                savedfunctions.Final_Accept = false;

                                //savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;
                                var segment = files.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                                savedfunctions.svdfunc_FileNameExcel = filename;
                                savedfunctions.svdfunc_FileSystemNameExcel = files.FileName;
                                if (savedfunctionRepo.Create(savedfunctions) == "True")
                                {
                                }
                                for (int i = 3; i < numberOfUsers; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = System.Convert.ToInt32(personalid);
                                    var UserID = await db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).Select(o=>o.usr_ID).FirstOrDefaultAsync();
                                    int shomarande = 3;


                                    if (lisstintmoalf.Count != 0)
                                    {
                                        foreach (var item in lisstintmoalf)
                                        {
                                            var Value = row.Cells[shomarande];
                                            if (Value.Value != null)
                                            {
                                                if (Value.Value.ToString() != "")
                                                {
                                                    var MoalefeID = item;


                                                    value = Value.Value.ToString();
                                                    tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                    {
                                                        MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                                                        MoalfeVal_FKUser = UserID,
                                                        MoalfeVal_Month = System.Convert.ToInt32(month),
                                                        MoalfeVal_Year = System.Convert.ToInt32(year),
                                                        MoalfeVal_Value = System.Convert.ToDouble(value),
                                                        FK_SavedFunctionsID = savedfunctions.svdfunc_ID,


                                                    };
                                                    lstMoalefeexcel.Add(dastmozdexcel);
                                                    //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                    //{
                                                    //}
                                                    //else
                                                    //{
                                                    //    transaction.Rollback();
                                                    //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                    //}


                                                }

                                                else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                {

                                                    tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                    {
                                                        MoalfeVal_FKMoalafeDastmozdi = item,
                                                        MoalfeVal_FKUser = UserID,
                                                        MoalfeVal_Month = System.Convert.ToInt32(month),
                                                        MoalfeVal_Year = System.Convert.ToInt32(year),
                                                        MoalfeVal_Value = 0,

                                                        FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                    };
                                                    //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                    //{
                                                    //}
                                                    //else
                                                    //{
                                                    //    transaction.Rollback();
                                                    //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                    //}
                                                }
                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = item,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0,
                                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                };
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }

                                            shomarande++;
                                        }


                                    }
                                    else
                                    {
                                        foreach (var item in lstMoalefe)
                                        {
                                            var Value = row.Cells[shomarande];
                                            if (Value.Value != null)
                                            {
                                                if (Value.Value.ToString() != "")
                                                {
                                                    var MoalefeID = item;


                                                    value = Value.Value.ToString();
                                                    tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                    {
                                                        MoalfeVal_FKMoalafeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).Select(s => s.md_ID).FirstOrDefault(),
                                                        MoalfeVal_FKUser = UserID,
                                                        MoalfeVal_Month = System.Convert.ToInt32(month),
                                                        MoalfeVal_Year = System.Convert.ToInt32(year),
                                                        MoalfeVal_Value = System.Convert.ToDouble(value),
                                                        FK_SavedFunctionsID = savedfunctions.svdfunc_ID,


                                                    };
                                                    lstMoalefeexcel.Add(dastmozdexcel);
                                                    //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                    //{
                                                    //}
                                                    //else
                                                    //{
                                                    //    transaction.Rollback();
                                                    //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                    //}


                                                }

                                                else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                                {

                                                    tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                    {
                                                        MoalfeVal_FKMoalafeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).Select(s => s.md_ID).FirstOrDefault(),
                                                        MoalfeVal_FKUser = UserID,
                                                        MoalfeVal_Month = System.Convert.ToInt32(month),
                                                        MoalfeVal_Year = System.Convert.ToInt32(year),
                                                        MoalfeVal_Value = 0,

                                                        FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                    };
                                                    //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                    //{
                                                    //}
                                                    //else
                                                    //{
                                                    //    transaction.Rollback();
                                                    //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                    //}
                                                }
                                            }

                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).Select(s => s.md_ID).FirstOrDefault(),
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = System.Convert.ToInt32(month),
                                                    MoalfeVal_Year = System.Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0,
                                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,
                                                };
                                                //if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                //{
                                                //}
                                                //else
                                                //{
                                                //    transaction.Rollback();
                                                //    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                //}
                                            }

                                            shomarande++;
                                        }

                                    }



                                }
                                if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
                                {
                                }
                                else
                                {
                                    if (await moalefeexcelRepo.Create2Async(lstMoalefeexcel) > 0)
                                    {
                                        // Saving successful, number of saved items is returned by Create2Async
                                        // You can use the returned value here, for example:
                                        //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                                        //Console.WriteLine($"Successfully saved {savedCount} items.");
                                    }
                                    else
                                    {
                                        // Saving might have failed (0 items saved) or encountered an exception
                                        Console.WriteLine("Saving failed.");
                                    }

                                }


                                //    else
                                //    {
                                //        transaction.Rollback();
                                //        return "در ذخیره سازی مشکلی به وجود آمده است";
                                //    }
                                //}


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }








        //برای اکسل کران ها 

        public ActionResult ExportMoalefeExcel4()
        {
            var OutPutFile = SetDataExcel_Moalefe5();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Moalefestandard" + extension);
        }

        public ActionResult ExportMoalefeExcel8(int MoalefeID = 0, int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        }

        public ActionResult ExportMoalefeExcel81(int MoalefeID = 0, int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe172(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        }
        public ActionResult ExportMoalefeExcel8122(List<int> MoalefeID = null, int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe175(MoalefeID, Month, Year);
            string extension = ".xlsx";
            
            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        }

        public ActionResult ExportMoalefeExcel812232(List<string> MoalefeID = null, int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe1754(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        }

        //public ActionResult ExportMoalefeExcelcontrolmahdodat(List<string> MoalefeID = null, int Month = 0, int Year = 0)
        //{
        //    var OutPutFile = ExportMoalefeExcelcontrolmahdodatdown(MoalefeID, Month, Year);
        //    string extension = ".xlsx";

        //    var stream = new MemoryStream();
        //    OutPutFile.Save(stream, extension);
        //    var mimeType = MimeTypes.ByExtension[extension];

        //    return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        //}

        public ActionResult ExportMoalefeExcelcontrolmahdodat(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = ExportMoalefeExcelcontrolmahdodatdowncorrect(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }
        public ActionResult ExportMoalefeExcel81223(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfe" + extension);
        }
        public ActionResult hesabreyali(int Month = 11, int Year = 1404, int day = 30, int kasri = 3)
        {
            try
            {
                var FINDFISH = db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.Finalaccept == true).ToList();
                var find20 = db.tbkarkard_notsystem.Where(s => s.moadel_kol != null && s.month == Month && s.year == Year).ToList();

                var findusr = db.tbMoadelPadashJarimeAyab.Where(s => s.Moadel != null && s.Year == Year && s.Month == Month).ToList();
                foreach (var fin in findusr)
                {
                    double xx = 0;
                    var find2 = find20.Where(s => s.FK_usr == fin.UserID && s.month == Month && s.year == Year).OrderByDescending(s => s.ID).FirstOrDefault();
                    if (find2 != null)
                    {
                        xx = (double)find2.moadel_kol;

                    }
                    else
                    {
                        xx = (double)fin.Moadel;

                    }
                    if (xx > 1)
                    {
                        int kar = 0;
                        int kar2 = 0;
                        double kar5 = 0;
                        double kar6 = 0;

                        var find1964 = FINDFISH.Where(s => s.FK_User == fin.UserID && s.FK_Moalefe == 1964).FirstOrDefault();

                        var find = FINDFISH.Where(s => s.FK_User == fin.UserID && s.FK_Moalefe == 1965).FirstOrDefault();
                        var find1961 = FINDFISH.Where(s => s.FK_User == fin.UserID && s.FK_Moalefe == 1961).FirstOrDefault();
                        if (find != null && find.mlfvlfsh_Value != null)
                        {
                            kar = (int)(day - find.mlfvlfsh_Value);

                        }
                        else
                        {
                            kar = day;
                        }
                        if (find1964 != null && find1964.mlfvlfsh_Value != null)
                        {

                            kar2 = (int)(kar - find1964.mlfvlfsh_Value);



                        }
                        else
                        {
                            kar2 = kar;
                        }
                        kar5 = xx - 1;
                        kar6 = Math.Ceiling(kar5 * 7.3333 * kar2);
                        if (find1961 != null)
                        {
                            find1961.mlfvlfsh_Value = kar6;

                        }
                    }

                }
                db.SaveChanges();

                return Content("true");
            }
            catch (Exception ex)
            {
                return Content(ex.Message);


            }

        }

        public ActionResult ExportMoalefeExcel81223forcontrol(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542forcontrol(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }

        public ActionResult ExportMoalefeExcel81223forcontrolmoadel(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542forcontrol_moadel(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }

        public ActionResult ExportMoalefeExcel81223forcontrolversionasli(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542versionasli(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }
        public ActionResult ExportMoalefeExcel81223forcontrolversionasli_mod(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542versionasli_modd(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }

        public ActionResult ExportMoalefeExcel81223forcontrolversionasli_mahdodyary(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542versionasli_saghf(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol_sagf" + extension);
        }

        public ActionResult ExportMoalefeExcel81223forcontrolversionasliselect(string MoalefeID = "", int Month = 0, int Year = 0)
        {
            var OutPutFile = SetDataExcel_Moalefe17542versionaslitest(MoalefeID, Month, Year);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Hilowmoalfecontrol" + extension);
        }

        public ActionResult ExportMoalefeExcel7()
        {
            var OutPutFile = SetDataExcel_Moalefe6();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Moalefestandard" + extension);
        }
        public ActionResult ExportMoalefeExcel7forsabt()
        {
            var OutPutFile = SetDataExcel_Moalefe6forsabt();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "MOalfehfish" + extension);
        }
        public ActionResult ExportMoalefeExcel7forkarkardi()
        {
            var OutPutFile = SetDataExcel_Moalefe6karkardi();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "MOalfehkarakrdi" + extension);
        }
        public ActionResult sabttaeed(int id=0)
        {
            var OutPutFile = SetDataExcel_Moalefe6sabttaeed(id);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            if (id == 1)
            {
                return File(stream.ToArray(), mimeType, "sabt" + extension);

            }
            else if (id == 2)
            {
                return File(stream.ToArray(), mimeType, "taeed" + extension);

            }
            else  if(id==3){
                return File(stream.ToArray(), mimeType, "control" + extension);

            }
            else
            {
                return File(stream.ToArray(), mimeType, "sabttaeed" + extension);

            }
        }
        public ActionResult ExportMoalefeExcel7forkarsoart()
        {
            var OutPutFile = SetDataExcel_Moalefe6soratvaziat();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "MOalfehsoart" + extension);
        }
        public ActionResult ExportMoalefeExcel7taghiz()
        {
            var OutPutFile = SetDataExcel_Moalefe6taghiz();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "taghiz" + extension);
        }

        public ActionResult ExportMoalefeExcel10()
        {
            var OutPutFile = SetDataExcel_Moalefe16();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Amalkard" + extension);
        }
        public ActionResult ImportExcel_AddSabtmarahel(HttpPostedFileBase MyExcelStream, int Basteh)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_FishtAEAD(MyExcelStream, Basteh);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public string GetDataFromExcel_FishtAEAD(HttpPostedFileBase MyExcelStream, int Basteh)
        {
            int mlfvlfsh_Year, mlfvlfsh_Month, malefe;
            float mlfvlfsh_Value;
            int id = 0;
            string name = "";
            var moalf = db.tbContractMoalefeDastmozdi.ToList();
            List<tbReffrenceSaveLevelMoalefe> IsFish = new List<tbReffrenceSaveLevelMoalefe>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title = workbook.Sheets[0].Rows[0].Cells;
                                var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var list = db.tbReffrenceSaveLevelMoalefe.Where(p => p.FK_LevelID == Basteh && p.Deleted2 != true).ToList();
                                foreach (var it in list)
                                {
                                    it.Deleted2 = true;
                                    db.SaveChanges();
                                }
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (int.TryParse(Name.Value.ToString(), out int result))
                                    {
                                        id = result;
                                    } // Assuming Name.Value is convertible to an integer.                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                        //var c2 = moalf.Where(p => p.md_ID==id).FirstOrDefault();
                                   
                                        tbReffrenceSaveLevelMoalefe tbm = new tbReffrenceSaveLevelMoalefe();
                                        tbm.FK_LevelID = Basteh;
                                        tbm.FK_MoalefeID = id;
                                        IsFish.Add(tbm);
                                    

                                    
                                  


                                }
                                db.tbReffrenceSaveLevelMoalefe.AddRange(IsFish);
                                db.SaveChanges();
                                break;
                            }
                        }
                        else
                        {
                            return "این شیت فاقد سطر می باشد";
                        }
                    }
                    else
                    {
                        return "هیچ شیتی در این اکسل وجود ندارد";
                    }
                    transaction.Commit();
                    return "با موفقیت انجام شد";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;

                }
            }
        }

        //public ActionResult ImportExcel_Zaribmalyat(HttpPostedFileBase MyExcelStream)
        //{
        //    if (MyExcelStream != null)
        //    {

        //        string Message = zaribmalyat(MyExcelStream);
        //        //TempData["Message"] = Message;
        //        //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
        //        return Content(Message);

        //    }
        //    else
        //    {
        //        return Content("فایل بطور صحیح بارگذاری نشده است");

        //    }

        //}

        //public string zaribmalyat(HttpPostedFileBase MyExcelStream)
        //{
        //    int mlfvlfsh_Year, mlfvlfsh_Month, malefe;
        //    float mlfvlfsh_Value;
        //    string name = "";
        //    List<tbMoalefeValueFish> IsFish = new List<tbMoalefeValueFish>();
        //    var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
        //    using (DbContextTransaction transaction = db.Database.BeginTransaction())
        //    {
        //        try
        //        {
        //            var Count_Sheets = workbook.Sheets.Count;
        //            if (Count_Sheets >= 1)
        //            {
        //                var sheet = workbook.Sheets[0];
        //                var Rows = sheet.Rows;
        //                if (Rows.Count >= 1)
        //                {
        //                    foreach (var item in Rows)
        //                    {
        //                        if (item.Cells.Count != sheet.Rows[0].Cells.Count)
        //                        {
        //                            return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
        //                            ;
        //                        }
        //                        var title = workbook.Sheets[0].Rows[0].Cells;
        //                        var List = workbook.Sheets[0].Rows;
        //                        var count = workbook.Sheets[0].Rows.Count();
        //                        if (count == 1)
        //                        {
        //                            return "فایل اکسل فاقد اطلاعات می باشد";
        //                        }
                             
        //                        for (int i = 1; i < count; i++)
        //                        {
        //                            var row = workbook.Sheets[0].Rows[i];
        //                            var Name = row.Cells[0];
        //                            if (Name.Value != null)
        //                            {
        //                                name = ((string)Name.Value).TrimEnd('\n');
        //                            } // Assuming Name.Value is convertible to an integer.                                    }
        //                            else
        //                            {
        //                                return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
        //                            }

        //                            var c2 = db.tbCities.Where(p => p.Name.Contains(name)).FirstOrDefault();
        //                            if (c2 != null)
        //                            {

        //                                c2.zaribmalyat = ;
        //                                tbm.FK_MoalefeID = c2.md_ID;
        //                                db.tbReffrenceSaveLevelMoalefe.Add(tbm);
        //                                db.SaveChanges();

        //                            }



        //                        }

        //                        break;
        //                    }
        //                }
        //                else
        //                {
        //                    return "این شیت فاقد سطر می باشد";
        //                }
        //            }
        //            else
        //            {
        //                return "هیچ شیتی در این اکسل وجود ندارد";
        //            }
        //            transaction.Commit();
        //            return "با موفقیت انجام شد";
        //        }
        //        catch (Exception ex)
        //        {
        //            transaction.Rollback();
        //            return ex.Message;

        //        }
        //    }

        //}
        public ActionResult ExportMoalefeExcelformoalfefish(int PeymanID)
        {
            var OutPutFile = SetDataExcel_Moalefeforfish(PeymanID);
            string extension = ".xlsx";
            var x = DateTime.Now;
            var c2 = db.tbReffrenceSaveLevel.Select(s => s.Data_time_create).FirstOrDefault();


            // اختلاف زمانی بین DateTime.Now و c2
            TimeSpan timeDifference = DateTime.Now.TimeOfDay - c2.TimeOfDay; ;

            // نمایش اختلاف زمانی به صورت تایمی (ساعت، دقیقه، ثانیه)
            string timeDifferenceString = $"{timeDifference.Hours}:{timeDifference.Minutes}:{timeDifference.Seconds}";

            // یا اگر می‌خواهید اختلاف را به عنوان یک مقدار ثانیه نیز داشته باشید:
            int totalSecondsDifference = (int)timeDifference.TotalSeconds;

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            return File(stream.ToArray(), mimeType, "MoalefeFish" + extension);
        }

        public Workbook SetDataExcel_Moalefeforfish(int PeymanID)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/MoalefeahFish.xlsx"));
            Row Row;
            var usr = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == PeymanID && p.Status == true).ToList();
            var x = db.tbUsers.ToList();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();

            int counter = 1;
            int counter2 = 1;

            foreach (var item in usr)
            {
                if (item.tbUsers.usr_Personal_ID == null)
                {

                }
                else
                {
                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.tbUsers.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                        {
                            Value =  item.tbUsers.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },

                    });

                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }

            }
            foreach(var it in moalfe)
            {
                Row = new Row() { Height = 20, Index = counter2 };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = it.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                           

                    });

                    counter2++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }
            }

            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe5()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/estandard.xlsx"));
            Row Row;
            var x = db.tbContractMoalefeDastmozdi.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                var x2 = db.tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {

                           new Cell()
                        {
                            Value = item.md_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                            new Cell()
                        {
                            Value = x2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }


        public Workbook SetDataExcel_Moalefe6()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/kolmoalfe.xlsx"));
            Row Row;
            var x = db.tbContractMoalefeDastmozdi.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                var x2 = db.tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.md_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                            new Cell()
                        {
                            Value = x2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }


        public Workbook SetDataExcel_Moalefe6forsabt()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/kolmoalfe.xlsx"));
            Row Row;
            var x = db.tbContractMoalefeDastmozdi.Where(p=>p.md_Type==10).ToList();
            var tbCategories = db.tbCategories.ToList();
            int counter = 1;
            foreach (var item in x)
            {
                var x2 = tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.md_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                            new Cell()
                        {
                            Value = x2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe6karkardi()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/kolmoalfe.xlsx"));
            Row Row;
            var x = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type ==8).ToList();
            var tbCategories = db.tbCategories.ToList();
            int counter = 1;
            foreach (var item in x)
            {
                var x2 = tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.md_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                            new Cell()
                        {
                            Value = x2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe6sabttaeed(int id=0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/sabttaeed.xlsx"));
            if (id != 1)
            {
                Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/SABTTAEEDCONTROL.xlsx"));

            }
            Row Row;
            var x = db.tbmarahlsabtbastedit2.Where(p => p.Has == true).ToList();
            var usr = db.tbUsers.ToList();
            var tbCategories = db.tbCategories.ToList();
            int counter = 1;
            var finddd = db.tbmarhesabt3.FirstOrDefault();
            if (id == 1)
            {
                finddd = db.tbmarhesabt3.Where(p => p.marahelsabt != null).FirstOrDefault();

            }
           else if (id == 2)
            {
                finddd = db.tbmarhesabt3.Where(p => p.taeed != null).FirstOrDefault();

            }
            else if (id == 3)
            {
                finddd = db.tbmarhesabt3.Where(p => p.controlmarahel != null).FirstOrDefault();

            }
            foreach (var item in x)
            {
                int t = 1;
                if (id == 1)
                {
                    for (var it = 3; it < finddd.marahelsabt + 3; it++)
                    {
                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                                  new Cell()
                        {
                            Value = item.FK_pymn,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                                    new Cell()
                        {
                            Value = item.FK_City,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                      new Cell()
                        {
                            Value = item.FK_namebAST,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                           new Cell()
                        {
                            Value = item.tbPeymanContracts.pec_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                        new Cell()
                        {
                            Value = item.tbCities.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                            new Cell()
                        {
                            Value = item.tbmarahlsabtbastedit1.naem,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                                new Cell()
                        {
                            Value = "ثبت",
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },
                                      new Cell()
                        {
                            Value = t,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        },

                    });
                            t++;
                            counter++;

                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                }
                else if (id == 2)
                {
                    for (var it = 3; it < finddd.taeed + 3; it++)
                    {
                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                                  new Cell()
                        {
                            Value = item.FK_pymn,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                                    new Cell()
                        {
                            Value = item.FK_City,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                      new Cell()
                        {
                            Value = item.FK_namebAST,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                           new Cell()
                        {
                            Value = item.tbPeymanContracts.pec_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                        new Cell()
                        {
                            Value = item.tbCities.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                            new Cell()
                        {
                            Value = item.tbmarahlsabtbastedit1.naem,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                                new Cell()
                        {
                            Value = "تایید",
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },
                                      new Cell()
                        {
                            Value = t,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        },

                    });
                            t++;
                            counter++;

                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                }
                else if (id == 3)
                {
                    for (var it = 3; it < finddd.controlmarahel + 3; it++)
                    {
                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                                  new Cell()
                        {
                            Value = item.FK_pymn,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                                    new Cell()
                        {
                            Value = item.FK_City,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                      new Cell()
                        {
                            Value = item.FK_namebAST,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                           new Cell()
                        {
                            Value = item.tbPeymanContracts.pec_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                        new Cell()
                        {
                            Value = item.tbCities.Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                            new Cell()
                        {
                            Value = item.tbmarahlsabtbastedit1.naem,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                                new Cell()
                        {
                            Value = "کنترل",
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },
                                      new Cell()
                        {
                            Value = t,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        },

                    });
                            t++;
                            counter++;

                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                }

                int t222 = 1;
               




            }
             counter = 1;

            foreach (var item in usr)
            {
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                          

                    });
                    counter++;

                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }
            }


                return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe6soratvaziat()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/kolmoalfe.xlsx"));
            Row Row;
            var x = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 11).ToList();
            var tbCategories = db.tbCategories.ToList();
            int counter = 1;
            foreach (var item in x)
            {
                var x2 = tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.md_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                            new Cell()
                        {
                            Value = x2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe6taghiz()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/kolmoalfe.xlsx"));
            Row Row;
            var x = db.tbEquipmentBunch.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                //var x2 = db.tbCategories.Where(p => p.Category_ID == item.FK_Category_ID).Select(s => s.Category_Name).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                           new Cell()
                        {
                            Value = item.Eqpbnch_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.Eqpbnch_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        //    new Cell()
                        //{
                        //    Value = x2,
                        //    FontFamily = "B Nazanin",
                        //    Bold = false,
                        //    Enable = true,
                        //    Wrap = false,
                        //    FontSize = 12,
                        //    Italic = false,
                        //    Underline = false,
                        //    Index = 2
                        //},

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe172(int MoalefeID, int Month, int Year)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/gozareshHolow.xlsx"));
            Row Row;
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var py = db.Link_User_And_Peyman.ToList();
            var fish = db.tbMoalefeValueFish.Where(p=>p.mlfvlfsh_Year== Year && p.mlfvlfsh_Month==Month && p.tbUsers.usr_amani != true).ToList();
            var distinctFishByUserID = fish.GroupBy(f => f.FK_User)
                                .Select(group => group.FirstOrDefault())
                                .ToList();
            var usr = db.tbUsers.ToList();
            var moadl = db.tbMoadelPadashJarimeAyab.ToList();

            int counter = 1;
            int counter2 = 0;
            //var hogog = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)").FirstOrDefault();
            //var malyat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "مالیات سهم کارمند (ریال)").FirstOrDefault();
            //var car = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").FirstOrDefault();

            //var ezafecar = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").FirstOrDefault();
            //var mashmolbemeh = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل مشمول بیمه (ریال)").FirstOrDefault();
            //var mashmolmalyat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل مشمول مالیات (ریال)").FirstOrDefault();
            //var bemehsahmkarmand = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "حق بیمه سهم کارمند (ریال)").FirstOrDefault();
            //var ezaafehkarireyal = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "اضافه کاری (ریال)").FirstOrDefault();

            //var kolkosarat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل کسورات (ریال)").FirstOrDefault();
            //var kalesgabeldaryaft = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)").FirstOrDefault();


            try
            {
                foreach (var item in distinctFishByUserID)
                {
                    var xx = usr.Where(p => p.usr_ID == item.FK_User&&p.usr_amani!=true).FirstOrDefault();
                    var hogog = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var malyat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "مالیات سهم کارمند (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var car = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();

                    var ezafecar = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var mashmolbemeh = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل مشمول بیمه (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var mashmolmalyat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل مشمول مالیات (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var bemehsahmkarmand = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "حق بیمه سهم کارمند (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var ezaafehkarireyal = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "اضافه کاری (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();

                    var kolkosarat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل کسورات (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var kalesgabeldaryaft = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع ناخالص حقوق و مزایا (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var Zagherh = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var Zagherh2 = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "ذخیره کار مازاد (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var mozd = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "مزد سنوات (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var rozkarkard = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز کارکرد " && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var tablet = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه(ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var abzar = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var moadel = moadl.Where(p => p.UserID == xx.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                    var omer = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "بیمه عمر و حادثه" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var Atash = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "بیمه آتش سوزی" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var kasr = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "کسر اقساط مواد غذایی" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var takmili = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "بیمه تکمیلی(ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    float x = 0;
                    if(takmili != null)
                    {
                        x = (float)takmili.mlfvlfsh_Value;
                    }
                    var saat = db.tbMaxkarkardMonth.Where(p => p.FKUser == xx.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                    var pymn= py.Where(p=>p.FK_User_ID==xx.usr_ID&&p.Status==true).FirstOrDefault();
                    var eydi = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "عیدی و پاداش ( ماهانه-ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var sanavat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "سنوات خدمت (ماهانه-ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();

                    var mamoriat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var mamoriatrial = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "ماموریت (ریال)" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();
                    var estekagi = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی" && p.FK_User == xx.usr_ID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month).FirstOrDefault();

                    if (hogog != null)
                    {
                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {

                            Value = xx.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                        {
                            Value = xx.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                new Cell()
                        {
                            Value = hogog.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                                new Cell()
                        {
                            Value = malyat.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                               new Cell()
                        {
                            Value = car.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                                new Cell()
                        {
                            Value = ezafecar.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                                new Cell()
                        {
                            Value = mashmolbemeh.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },
                                new Cell()
                        {
                            Value = mashmolmalyat.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        },
                                new Cell()
                        {
                            Value = bemehsahmkarmand.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        },
                                new Cell()
                        {
                            Value = ezaafehkarireyal.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 9
                        },
                                new Cell()
                        {
                            Value = kolkosarat.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 10
                        },
                                   new Cell()
                        {
                            Value = kalesgabeldaryaft.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 11
                        },
                                      new Cell()
                        {
                            Value = Zagherh.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 12
                        },
                                         new Cell()
                        {
                            Value = Zagherh2.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 13
                        },
                                                        new Cell()
                        {
                            Value = mozd.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 14
                        },
                                                                       new Cell()
                        {
                            Value = rozkarkard.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 15
                        },                           new Cell()
                        {
                            Value = tablet.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 16
                        },                           new Cell()
                        {
                            Value = abzar.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 17
                        },
                                                                                  new Cell()
                        {
                            Value =(double)moadel.Moadel,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 18
                        },
                                                                                             new Cell()
                        {
                            Value = omer.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 19
                        },
                                                                                                        new Cell()
                        {
                            Value = Atash.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 20
                        },
                                                                                                                   new Cell()
                        {
                            Value = kasr.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 21
                        },
                                                                                                                              new Cell()
                        {
                            Value = x,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 22
                        },                                                                                                         new Cell()
                        {
                            Value = pymn.tbPeymanContracts.pec_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 24
                        },
                                                                                                                               new Cell()
                        {
                            Value = estekagi.mlfvlfsh_Value??0,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 25
                        },
                                                                                                                                new Cell()
                        {
                                                        Value = mamoriat.mlfvlfsh_Value??0,

                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 26
                        },
                                                                                                                                 new Cell()
                        {
                                                      Value = mamoriatrial.mlfvlfsh_Value??0,

                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 27
                        },

 new Cell()
                        {
                                                      Value = eydi.mlfvlfsh_Value??0,

                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 28
                        },
  new Cell()
                        {
                                                      Value = sanavat.mlfvlfsh_Value??0,

                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 29
                        },
  new Cell()
                        {
                                                      Value = saat.RemainValue??0,

                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 30
                        },
                    });

                            counter++;
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                 




                }
            }catch(Exception ex)
            {
                
            }
       
            return Moalefeexcelfile;
        }


        public Workbook SetDataExcel_Moalefe17(int MoalefeID, int Month, int Year)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            var y = db.tbMoalefeValueFish.Where(p => p.FK_Moalefe == MoalefeID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.tbUsers.usr_amani != true).OrderByDescending(p => p.mlfvlfsh_Value).ToList();
            int counter = 1;
            int counter2 = 0;
            var t = db.tbContractMoalefeDastmozdi.Where(p => p.md_ID == MoalefeID).FirstOrDefault();
            Row = new Row() { Height = 20, Index = counter2 };
            Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = t.md_Title,
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 2
    }
});
            Moalefeexcelfile.Sheets[0].AddRow(Row);

            foreach (var item in y)
            {
                var xx = db.tbUsers.Where(p => p.usr_ID == item.FK_User).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {

                            Value = xx.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                        {
                            Value = xx.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                new Cell()
                        {
                            Value = item.mlfvlfsh_Value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },


                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe1754(List<string> MoalefeID2, int Month, int Year)
        {
            List<int> MoalefeID = new List<int>();

            foreach (var it in MoalefeID2)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.tbUsers.usr_amani != true&&p.Finalaccept==true)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;
            var t = db.tbContractMoalefeDastmozdi.ToList();

            foreach (var item in y)
            {
                foreach (var item2 in MoalefeID)
                {
                    var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();
                    if (ex != null)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = ex.md_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counter2
                    },
                });
                    }

                    var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.FK_User).FirstOrDefault();
                    double ty = 0;
                    if (exx != null)
                    {
                        ty = (double)exx.mlfvlfsh_Value;
                    }
                    var xx = db.tbUsers.Where(p => p.usr_ID == item.FK_User).FirstOrDefault();
                    Row = new Row() { Height = 20, Index = counter };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = xx.FullName,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 0
                },
                new Cell()
                {
                    Value = xx.usr_Personal_ID,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
                new Cell()
                {
                    Value = ty,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 2
                },
            });

                    counter2++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }
                counter++;

            }

            return Moalefeexcelfile;
        }




        public Workbook ExportMoalefeExcelcontrolmahdodatdown(List<string> MoalefeID2, int Month, int Year)
        {
            List<int> MoalefeID = new List<int>();

            foreach (var it in MoalefeID2)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.tbUsers.usr_amani != true && p.Finalaccept == true)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;
            var t = db.tbContractMoalefeDastmozdi.ToList();

            foreach (var item in y)
            {
                foreach (var item2 in MoalefeID)
                {
                    var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();
                    if (ex != null)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = ex.md_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counter2
                    },
                });
                    }

                    var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.FK_User).FirstOrDefault();
                    double ty = 0;
                    if (exx != null)
                    {
                        ty = (double)exx.mlfvlfsh_Value;
                    }
                    var xx = db.tbUsers.Where(p => p.usr_ID == item.FK_User).FirstOrDefault();
                    Row = new Row() { Height = 20, Index = counter };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = xx.FullName,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 0
                },
                new Cell()
                {
                    Value = xx.usr_Personal_ID,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
                new Cell()
                {
                    Value = ty,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 2
                },
            });

                    counter2++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }
                counter++;

            }

            return Moalefeexcelfile;
        }

        
        private string GetSqlConnectionString()
        {
            try
            {
                // روش 1: استخراج از Entity Connection موجود
                var entityConnection = new EntityConnection(ConfigurationManager.ConnectionStrings["SaabEntities"].ConnectionString);
                return entityConnection.StoreConnection.ConnectionString;
            }
            catch
            {
                // روش 2: استفاده از connection string مستقیم (اگر روش اول کار نکرد)
                return "REPLACE_WITH_CONNECTION_STRING";
            }
        }
        public Workbook SetDataExcel_Moalefe17542forcontrol(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            foreach (var it in moalefeIDList)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1965 && p.Finalaccept == true && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel != null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel != null && p.Finalaccept == true)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 9;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                   ,


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        }
                    });


                        //مقادیر مولفه ها
                        int index = 9;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = 0;
                            if (exx != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe17542forcontrol_moadel(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            foreach (var it in moalefeIDList)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoadelPadashJarimeAyab.Where(p => p.Month == Month && p.Year == Year).ToList())
            {
                var finus = db.tbUsers.Where(s => s.usr_ID == item.UserID).FirstOrDefault();
                if (finus != null)
                {
                    tyyy.Add(finus);

                }
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel != null && p.Finalaccept == true)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 9;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                   ,


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        }
                    });


                        //مقادیر مولفه ها
                        int index = 9;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = 0;
                            if (exx != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe17542forcontrolversionasli(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            foreach (var it in moalefeIDList)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1961 && p.Finalaccept == true && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel != null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel != null && p.Finalaccept == true)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 8;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();

                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{it.MoalfeVal_Value} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = amani,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                    });


                        //مقادیر مولفه ها
                        int index = 8;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = 0;
                            if (exx != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe17542EDIT(string MoalefeID2, int Month, int Year)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            for (int i = Month; i <= Month; i++)
            {
                Month = i;
                // Split the input string into a list of strings
                var moalefeIDList = MoalefeID2.Split(',').ToList();
                List<int> MoalefeID = new List<int>();

                //Convert each string in the list to an integer
                foreach (var it in moalefeIDList)
                {
                    if (int.TryParse(it.Trim(), out int moalefeId))
                    {
                        MoalefeID.Add(moalefeId);
                    }
                    else
                    {
                        // Handle the invalid format case, you can log it or throw an exception
                        throw new FormatException($"Input string '{it}' was not in a correct format.");
                    }
                }
                //foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
                //{
                //    MoalefeID.Add((int)item.Key);

                //}

                Row Row;
                var x = db.tbMoadelPadashJarimeAyab.ToList();
                List<tbUsers> tyyy = new List<tbUsers>();
                foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList())
                {
                    tyyy.Add(item.tbUsers);
                }
                //var tyyy = db.tbUsers.ToList();
                var caran = db.tbCaranSettings.ToList();
                var asnad = db.tbfkfinancial.ToList();
                var FinancialDocuments = db.FinancialDocuments.ToList();

                List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
                foreach (var it in asnad)
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                    int persianYear = persianCalendar.GetYear(dataDocument);
                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                    if (persianYear == Year && persianMonth == Month)
                    {
                        asnadstr.Add(it);

                    }
                }
                var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

                var y = db.tbMoalefeValueFish
                            .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month)
                            .ToList();
                int counter = 1;
                int counter2 = 2;






                var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
                var t = db.tbContractMoalefeDastmozdi.ToList();
                var count = 10;
                foreach (var item in MoalefeID)
                {
                    var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                foreach (var item in asnadstr)
                {

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }






                foreach (var item in tyyy)
                {

                    string job = "";
                    if (item.tbjob != null)
                    {
                        job = item.tbjob.Name;
                    }
                    bool amani = false;
                    if (item.usr_amani == true)
                    {
                        amani = true;
                    }
                    var tb = noalf
    .Where(p => p.MoalfeVal_FKUser == item.usr_ID

                && p.MoalfeVal_Value != 0
                )
    .ToList();
                    List<string> mode = new List<string>();

                    foreach (var it in tb)
                    {
                        var yuuu = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                            .FirstOrDefault();

                        if (yuuu == null)
                        {
                            var u = caran
                                .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                                .FirstOrDefault();

                            if (u != null)
                            {
                                mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{it.MoalfeVal_Value} )----");
                            }
                        }
                        else
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
                        }
                    }
                    double moalde = 0;
                    double clock = 0;
                    var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                    if (moad != null && moad.Moadel != null)
                    {
                        moalde = moad.Moadel;
                    }
                    var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                    if (exi != null && exi.RemainValue != null)
                    {
                        clock = (double)exi.RemainValue;
                    }
                    if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                    {

                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                   ,


                                                   new Cell()
                        {
                            Value = Month,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        }
                                                           ,


                                                   new Cell()
                        {
                            Value = Year,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 9
                        }
                    });


                            //مقادیر مولفه ها
                            int index = 10;
                            foreach (var item2 in MoalefeID)
                            {
                                var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                double ty = 0;
                                if (exx != null)
                                {
                                    ty = (double)exx.mlfvlfsh_Value;
                                }
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            foreach (var item2 in asnadstr)
                            {
                                long value = 0;
                                var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                                foreach (var item1 in t3)
                                {
                                    if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                    {
                                        PersianCalendar persianCalendar = new PersianCalendar();
                                        DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                        int persianYear = persianCalendar.GetYear(dataDocument);
                                        int persianMonth = persianCalendar.GetMonth(dataDocument);
                                        if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                        {
                                            if (item1.Debtore != 0 && item1.Creditor != 0)
                                            {
                                                value = (long)(item1.Creditor - item1.Debtore);
                                            }


                                            else if (item1.Debtore != 0 && item1.Debtore != null)
                                            {
                                                value = (long)(item1.Debtore ?? 0);
                                            }
                                            else if (item1.Creditor != 0 && item1.Creditor != null)
                                            {
                                                value = (long)(item1.Creditor ?? 0);

                                            }

                                        }

                                    }
                                }

                                //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                //double ty = 0;
                                //if (exx != null)
                                //{
                                //    ty = (double)exx.mlfvlfsh_Value;
                                //}
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            counter++;
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                    //نام ، نام خانوادگی و کد پرسنلی





                }



























            }

            return Moalefeexcelfile;
        }



        //ارجاع
        public ActionResult DownloadCSV()
        {
            string fileName = "ExportedData.csv";

            // حافظه موقت برای CSV
            var stream = new MemoryStream();
            using (var writer = new StreamWriter(stream, Encoding.UTF8, 1024, true))
            {
                // --- هدر ---
                var moalefeIDs = db.erja_sabt_usr
                    .Where(s => s.FK_omomi != null &&
                                (s.value_cheek != null || s.value_int != null || s.value_str != null))
                    .Select(s => s.FK_erja_omomi_edit)
                    .Distinct()
                    .ToList();

                var editData = db.erja_omomi_edit.ToDictionary(x => x.ID, x => x.Nam);

                writer.Write("نام کاربر,شماره پرسنلی,شهر,");

                foreach (var id in moalefeIDs)
                    writer.Write(editData[(long)id] + ",");

                var fk1Data = db.dbergharelimsertfk1_7.ToList(); // کل دیتا به حافظه
                var distinctFk1Ids = fk1Data.Select(x => x.FK_db2).Distinct().ToList();

                foreach (var id in distinctFk1Ids)
                    writer.Write(id + ",");

                writer.WriteLine();

                // --- داده‌ها با batch ---
                int batchSize = 10000;
                int page = 0;
                bool hasMore = true;

                while (hasMore)
                {
                    // گرفتن batch کاربران
                    var likeUsersBatch = db.erja_like
                        .Include(x => x.tbUsers)
                        .Include(x => x.erja_like_link)
                        .OrderBy(x => x.ID)
                        .Skip(page * batchSize)
                        .Take(batchSize)
                        .ToList();

                    if (!likeUsersBatch.Any())
                    {
                        hasMore = false;
                        break;
                    }

                    var userIDs = likeUsersBatch.Select(x => x.ID).ToList();

                    // گرفتن sabtData مربوط به این batch
                    var sabtDataBatch = db.erja_sabt_usr
                        .Where(x => x.FK_omomi.HasValue && userIDs.Contains(x.FK_omomi.Value))
                        .ToList();

                    foreach (var user in likeUsersBatch)
                    {
                        var fullName = user.tbUsers.usr_Name + " " + user.tbUsers.usr_Family;
                        var pid = user.tbUsers.usr_Personal_ID;
                        var city = db.tbCities
                            .Where(c => c.ID == user.tbUsers.usr_City_Dutysystem)
                            .Select(c => c.Name)
                            .FirstOrDefault();

                        writer.Write($"{fullName},{pid},{city},");

                        // مقادیر moalefeIDs
                        foreach (var id in moalefeIDs)
                        {
                            var exx = sabtDataBatch
                                .FirstOrDefault(x => x.FK_omomi == user.ID && x.FK_erja_omomi_edit == id);

                            string val = exx?.value_cheek?.ToString() ??
                                         exx?.value_int?.ToString() ??
                                         exx?.value_str ?? string.Empty;

                            writer.Write(val + ",");
                        }

                        // مقادیر fk1Data با استفاده از حافظه
                        var userLinks = user.erja_like_link.Select(l => l.value).ToList();
                        foreach (var id in distinctFk1Ids)
                        {
                            var t3 = fk1Data.FirstOrDefault(p => userLinks.Contains(p.valuunic) && p.FK_db2 == id);
                            writer.Write((t3?.valueFK_db2 ?? "") + ",");
                        }

                        writer.WriteLine();
                    }

                    page++;
                }
            }

            stream.Position = 0;
            return File(stream, "text/csv", fileName);
        }


        public Workbook SetDataExcel_Moalefe17541112()
        {
            string MoalefeID2 = "";
            int Month = 1;
            int Year = 0;
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));

            for (int i = Month; i <= Month; i++)
            {
                Month = i;

                // لیست شناسه‌ها
                var moalefeIDList = MoalefeID2.Split(',').ToList();
                List<long> MoalefeID1 = new List<long>();
                List<long> MoalefeID3 = new List<long>();
                List<erja_like> erja_like = new List<erja_like>();

                // گروه‌بندی ارجاعات
                var find22 = db.erja_sabt_usr.Where(s => s.FK_omomi != null)
                                             .GroupBy(s => s.FK_omomi)
                                             .ToList();

                foreach (var it in find22)
                {
                    var find2 = db.erja_sabt_usr
                                  .Where(s => s.FK_omomi == it.Key && (s.value_cheek != null || s.value_int != null || s.value_str != null))
                                  .GroupBy(s => s.FK_erja_omomi_edit)
                                  .ToList();

                    foreach (var it1 in find2)
                        MoalefeID1.Add((long)it1.Key);

                    foreach (var it2 in db.erja_like.Where(s => s.ID == it.Key).ToList())
                    {
                        erja_like.Add(it2);

                        // تبدیل لینک‌ها به primitive
                        var likeValues = it2.erja_like_link.Select(x => x.value).ToList();

                        var find3 = db.dbergharelimsertfk1_7
                                      .Where(s => likeValues.Contains(s.valuunic))
                                      .GroupBy(s => s.FK_db2)
                                      .ToList();

                        foreach (var it1 in find3)
                            MoalefeID3.Add((long)it1.Key);
                    }
                }

                Row Row;

                int counter = 1;
                int count33 = 3;

                var t11 = db.erja_omomi_edit.ToList();
                var t12 = db.dbergha2.ToList();

                // اضافه کردن نام مولفه‌ها
                foreach (var item in MoalefeID1)
                {
                    var ex = t11.FirstOrDefault(p => p.ID == item);
                    string name = ex?.Nam ?? "";

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = name,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = count33
                },
            });
                    count33++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                foreach (var item in MoalefeID3)
                {
                    var ex = t12.FirstOrDefault(p => p.ID == item);
                    string name = ex?.name ?? "";

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = name,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = count33
                },
            });
                    count33++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                // اضافه کردن اطلاعات کاربران
                foreach (var item in erja_like)
                {
                    Row = new Row() { Height = 20, Index = counter };

                    string fullName = item.tbUsers?.FullName ?? "";
                    int personalID =(int) item.tbUsers?.usr_Personal_ID;
                    string cityName = item.tbUsers != null ?
                                      db.tbCities.Where(p => p.ID == item.tbUsers.usr_City_Dutysystem)
                                                 .Select(s => s.Name)
                                                 .FirstOrDefault() ?? "" : "";

                    Row.AddCells(new List<Cell>()
            {
                new Cell() { Value = fullName, FontFamily = "B Nazanin", Index = 0, FontSize=12 },
                new Cell() { Value = personalID, FontFamily = "B Nazanin", Index = 1, FontSize=12 },
                new Cell() { Value = cityName, FontFamily = "B Nazanin", Index = 2, FontSize=12 },
            });

                    int index = 3;

                    // مقادیر MoalefeID1
                    foreach (var item2 in MoalefeID1)
                    {
                        var exx = db.erja_sabt_usr
                                    .Where(p => p.FK_omomi == item.ID && p.FK_erja_omomi_edit == item2 &&
                                               (p.value_cheek != null || p.value_int != null || p.value_str != null))
                                    .FirstOrDefault();

                        string ty = "";
                        if (exx != null)
                        {
                            if (exx.value_cheek != null) ty = exx.value_cheek.ToString();
                            else if (exx.value_int != null) ty = exx.value_int.ToString();
                            else if (exx.value_str != null) ty = exx.value_str.ToString();
                        }

                        Row.AddCells(new List<Cell>()
                {
                    new Cell() { Value = ty, FontFamily = "B Nazanin", Index = index, FontSize=12, Wrap=true }
                });
                        index++;
                    }

                    // مقادیر MoalefeID3
                    foreach (var item2 in MoalefeID3)
                    {
                        var ids = item.erja_like_link.Select(s => s.value).ToList();

                        // بعد query روی دیتابیس
                        var t3List = db.dbergharelimsertfk1_7
                                       .Where(p => ids.Contains(p.valuunic))
                                       .ToList(); // materialize in memory

                        // سپس فیلتر نهایی در memory
                        var t3 = t3List.FirstOrDefault(p => p.FK_db2 == item2);
                        string valueFK_db2 = t3?.valueFK_db2 ?? "";

                        Row.AddCells(new List<Cell>()
                {
                    new Cell() { Value = valueFK_db2, FontFamily = "B Nazanin", Index = index, FontSize=12 }
                });
                        index++;
                    }

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }
            }

            return Moalefeexcelfile;
        }



        public ActionResult SetDataExcel_Moalefe175411132(DateTime fromDate, DateTime toDate)
        {
            var workbook = SetDataExcel_Moalefe175411132_Internal(fromDate, toDate);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            workbook.Save(stream, extension);
            stream.Position = 0;

            var mimeType = MimeTypes.ByExtension[extension];
            return File(stream.ToArray(), mimeType, "ارجاع" + extension);
        }




        //public ActionResult SetDataExcel_Moalefe175411132()
        //{
        //    var workbook = SetDataExcel_Moalefe175411132_Internal();
        //    string extension = ".xlsx";

        //    var stream = new MemoryStream();
        //    workbook.Save(stream, extension);
        //    stream.Position = 0;

        //    var mimeType = MimeTypes.ByExtension[extension];
        //    return File(stream.ToArray(), mimeType, "ارجاع" + extension);
        //}

        public Workbook SetDataExcel_Moalefe175411132_Internal(DateTime fromDate, DateTime toDate)
        {
            db.Configuration.AutoDetectChangesEnabled = false;
            db.Configuration.LazyLoadingEnabled = false;

            var workbook = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            var sheet = workbook.Sheets[0];

            /* ===================== Base Data ===================== */

            var erjaSabtUsr = db.erja_sabt_usr
               .Where(s =>
                   s.FK_omomi != null &&
                   s.timesabt_vaz0 != null &&
                   s.timesabt_vaz0 >= fromDate &&
                   s.timesabt_vaz0 <= toDate && s.vaziat == 2 &&

                   (s.value_cheek != null || s.value_int != null || s.value_str != null )
               )
               .ToList();

            var erjaLikeList = db.erja_like
                .Include("erja_like_link")
                .Include("tbUsers")
                .Where(s=>s.vaziat==2  &&s.erja_sabt_usr.Any(t=>t.timesabt_vaz0 != null &&
                   t.timesabt_vaz0 >= fromDate &&
                   t.timesabt_vaz0 <= toDate && t.vaziat == 2 &&

                   (t.value_cheek != null || t.value_int != null || t.value_str != null)))
                .ToList();

            var erjaOmomiEdit = db.erja_omomi_edit.ToList();
            var dbergha2 = db.dbergha2.ToList();
            var cities = db.tbCities.ToList();

            /* ===================== Dictionaries ===================== */

            var erjaOmomiDict = erjaOmomiEdit.ToDictionary(x => x.ID, x => x.Nam);
            var dberghaDict2 = dbergha2.ToDictionary(x => x.ID, x => x.name);
            var cityDict = cities.ToDictionary(x => x.ID, x => x.Name);

            /* ===================== IDs ===================== */

            var MoalefeID1 = erjaSabtUsr
                .Select(x => (long)x.FK_erja_omomi_edit)
                .Distinct()
                .ToList();

            var allValueUnics = erjaLikeList
                .SelectMany(x => x.erja_like_link.Select(l => l.value))
                .Distinct()
                .ToList();

            var dberghaData = db.dbergharelimsertfk1_7
                .Where(x => allValueUnics.Contains(x.valuunic) && x.FK_db2 != null)
                .Select(x => new
                {
                    x.valuunic,
                    FK_db2 = x.FK_db2.Value,
                    x.valueFK_db2
                })
                .ToList();

            var MoalefeID3 = dberghaData
                .Select(x => x.FK_db2)
                .Distinct()
                .ToList();

            var dberghaDict = dberghaData
                .GroupBy(x => (x.valuunic, x.FK_db2))
                .ToDictionary(g => g.Key, g => g.First().valueFK_db2);

            /* ===================== Header Row ===================== */

            var headerRow = new Row
            {
                Index = 0,
                Height = 25,
                Cells = new List<Cell>
        {
            new Cell { Index = 0, Value = "نام", FontFamily = "B Nazanin", FontSize = 12 },
            new Cell { Index = 1, Value = "کد پرسنلی", FontFamily = "B Nazanin", FontSize = 12 },
            new Cell { Index = 2, Value = "شهر", FontFamily = "B Nazanin", FontSize = 12 },
                        new Cell { Index = 3, Value = "شماره شناسایی", FontFamily = "B Nazanin", FontSize = 12 }

        }
            };

            int headerCol = 4;

            foreach (var id in MoalefeID1)
            {
                headerRow.Cells.Add(new Cell
                {
                    Index = headerCol++,
                    Value = erjaOmomiDict.TryGetValue(id, out var name) ? name : "",
                    FontFamily = "B Nazanin",
                    FontSize = 12
                });
            }

            foreach (var id in MoalefeID3)
            {
                headerRow.Cells.Add(new Cell
                {
                    Index = headerCol++,
                    Value = dberghaDict2.TryGetValue(id, out var name) ? name : "",
                    FontFamily = "B Nazanin",
                    FontSize = 12
                });
            }

            sheet.AddRow(headerRow);

            /* ===================== Data Rows ===================== */

            int rowIndex = 1;
            var likeValuesDict = erjaLikeList
    .ToDictionary(
        x => x.ID,
        x => x.erja_like_link.Select(l => l.value).ToList()
    );
            var likeValueDict = erjaLikeList
    .ToDictionary(
        x => x.ID,
        x => x.erja_like_link.Select(l => l.value).FirstOrDefault()
    );


            //       var excel = db.erja_like
            //.Include(x => x.erja_like_link).ToList();

            foreach (var item in erjaLikeList)
            {
                //var findunic = db.erja_like_link.Where(s => s.FK_erja_like == item.ID).FirstOrDefault();

                var row = new Row
                {
                    Index = rowIndex++,
                    Height = 20,
                    Cells = new List<Cell>
            {
                new Cell { Index = 0, Value = item.tbUsers?.FullName ?? "", FontFamily="B Nazanin", FontSize=12 },
                new Cell { Index = 1, Value = item.tbUsers?.usr_Personal_ID ?? 0, FontFamily="B Nazanin", FontSize=12 },
                new Cell
                {
                    Index = 2,
                    Value = cityDict.TryGetValue(item.tbUsers?.usr_City_Dutysystem ?? 0, out var city)
                            ? city : "",
                    FontFamily="B Nazanin",
                    FontSize=12
                },     
               
            new Cell
            {
                Index = 3,
                Value = likeValueDict.TryGetValue(item.ID, out var val)
                        ? val
                        : "",
                FontFamily = "B Nazanin",
                FontSize = 12
            }
            }
                };

                int colIndex = 4;

                foreach (var m1 in MoalefeID1)
                {
                    var exx = erjaSabtUsr.FirstOrDefault(p =>
                        p.FK_omomi == item.ID && p.FK_erja_omomi_edit == m1);

                    row.Cells.Add(new Cell
                    {
                        Index = colIndex++,
                        Value = exx?.value_cheek?.ToString()
                                ?? exx?.value_int?.ToString()
                                ?? exx?.value_str
                                ?? "",
                        FontFamily = "B Nazanin",
                        FontSize = 12,
                        Wrap = true
                    });
                }

                foreach (var m3 in MoalefeID3)
                {
                    string result = "";

                    foreach (var v in item.erja_like_link.Select(s => s.value))
                    {
                        if (dberghaDict.TryGetValue((v, m3), out var r))
                        {
                            result = r;
                            break;
                        }
                    }

                    row.Cells.Add(new Cell
                    {
                        Index = colIndex++,
                        Value = result,
                        FontFamily = "B Nazanin",
                        FontSize = 12
                    });
                }

                sheet.AddRow(row);
            }

            return workbook;
        }


        //public Workbook SetDataExcel_Moalefe175411132_Internal()
        //{
        //    db.Configuration.AutoDetectChangesEnabled = false;
        //    db.Configuration.LazyLoadingEnabled = false;

        //    var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));


        //    // ------------------- Base data
        //    var erjaSabtUsr = db.erja_sabt_usr
        //        .Where(s => s.FK_omomi != null &&
        //               (s.value_cheek != null || s.value_int != null || s.value_str != null))
        //        .ToList();

        //    var erjaLikeList = db.erja_like
        //        .Include("erja_like_link")
        //        .Include("tbUsers")
        //        .ToList();

        //    var erjaOmomiEdit = db.erja_omomi_edit.ToList();
        //    var dbergha2 = db.dbergha2.ToList();
        //    var cities = db.tbCities.ToList();

        //    // ------------------- MoalefeID1
        //    var MoalefeID1 = erjaSabtUsr
        //        .Select(x => (long)x.FK_erja_omomi_edit)
        //        .Distinct()
        //        .ToList();

        //    // ------------------- valuunic
        //    var allValueUnics = erjaLikeList
        //        .SelectMany(x => x.erja_like_link.Select(l => l.value))
        //        .Distinct()
        //        .ToList();

        //    // ------------------- Big table (filtered)
        //    var dberghaData = db.dbergharelimsertfk1_7
        //        .Where(x => allValueUnics.Contains(x.valuunic) && x.FK_db2 != null)
        //        .Select(x => new
        //        {
        //            x.valuunic,
        //            FK_db2 = x.FK_db2.Value,
        //            x.valueFK_db2
        //        })
        //        .ToList();

        //    // ------------------- MoalefeID3
        //    var MoalefeID3 = dberghaData
        //        .Select(x => x.FK_db2)
        //        .Distinct()
        //        .ToList();

        //    // ------------------- Dictionary SAFE
        //    var dberghaDict = dberghaData
        //        .GroupBy(x => (x.valuunic, x.FK_db2))
        //        .ToDictionary(
        //            g => g.Key,
        //            g => g.First().valueFK_db2
        //        );

        //    // ------------------- Headers
        //    int headerCol = 3;

        //    foreach (var id in MoalefeID1)
        //    {
        //        var ex = erjaOmomiEdit.FirstOrDefault(p => p.ID == id);
        //        var row = new Row() { Height = 20 };
        //        row.AddCells(new List<Cell>() {
        //    new Cell { Value = ex?.Nam ?? "", FontFamily="B Nazanin", FontSize=12, Index=headerCol }
        //});
        //        Moalefeexcelfile.Sheets[0].AddRow(row);
        //        headerCol++;
        //    }

        //    foreach (var id in MoalefeID3)
        //    {
        //        var ex = dbergha2.FirstOrDefault(p => p.ID == id);
        //        var row = new Row() { Height = 20 };
        //        row.AddCells(new List<Cell>() {
        //    new Cell { Value = ex?.name ?? "", FontFamily="B Nazanin", FontSize=12, Index=headerCol }
        //});
        //        Moalefeexcelfile.Sheets[0].AddRow(row);
        //        headerCol++;
        //    }

        //    // ------------------- Rows
        //    int rowIndex = 1;

        //    foreach (var item in erjaLikeList)
        //    {
        //        var row = new Row() { Height = 20, Index = rowIndex };

        //        row.AddCells(new List<Cell>() {
        //    new Cell { Value = item.tbUsers?.FullName ?? "", Index=0, FontFamily="B Nazanin", FontSize=12 },
        //    new Cell { Value = item.tbUsers?.usr_Personal_ID ?? 0, Index=1, FontFamily="B Nazanin", FontSize=12 },
        //    new Cell { Value = cities.FirstOrDefault(c => c.ID == item.tbUsers?.usr_City_Dutysystem)?.Name ?? "", Index=2, FontFamily="B Nazanin", FontSize=12 }
        //});

        //        int colIndex = 3;

        //        foreach (var m1 in MoalefeID1)
        //        {
        //            var exx = erjaSabtUsr.FirstOrDefault(p => p.FK_omomi == item.ID && p.FK_erja_omomi_edit == m1);
        //            string val = exx?.value_cheek?.ToString()
        //                         ?? exx?.value_int?.ToString()
        //                         ?? exx?.value_str
        //                         ?? "";
        //            row.AddCells(new List<Cell>() {
        //        new Cell { Value = val, Index=colIndex, FontFamily="B Nazanin", FontSize=12, Wrap=true }
        //    });
        //            colIndex++;
        //        }

        //        var values = item.erja_like_link.Select(s => s.value).ToList();

        //        foreach (var m3 in MoalefeID3)
        //        {
        //            string result = "";
        //            foreach (var v in values)
        //            {
        //                if (dberghaDict.TryGetValue((v, m3), out var r))
        //                {
        //                    result = r;
        //                    break;
        //                }
        //            }

        //            row.AddCells(new List<Cell>() {
        //        new Cell { Value = result, Index=colIndex, FontFamily="B Nazanin", FontSize=12 }
        //    });
        //            colIndex++;
        //        }

        //        Moalefeexcelfile.Sheets[0].AddRow(row);
        //        rowIndex++;
        //    }

        //    return Moalefeexcelfile;
        //}



        public Workbook SetDataExcel_Moalefe17542(string MoalefeID2, int Month, int Year)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            for (int i = Month; i <= Month; i++)
            {
                Month = i;
                // Split the input string into a list of strings
                var moalefeIDList = MoalefeID2.Split(',').ToList();
                List<int> MoalefeID = new List<int>();

                // Convert each string in the list to an integer
                //foreach (var it in moalefeIDList)
                //{
                //    if (int.TryParse(it.Trim(), out int moalefeId))
                //    {
                //        MoalefeID.Add(moalefeId);
                //    }
                //    else
                //    {
                //        // Handle the invalid format case, you can log it or throw an exception
                //        throw new FormatException($"Input string '{it}' was not in a correct format.");
                //    }
                //}
                foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
                {
                    MoalefeID.Add((int)item.Key);

                }

                Row Row;
                var x = db.tbMoadelPadashJarimeAyab.ToList();
                List<tbUsers> tyyy = new List<tbUsers>();
                foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList())
                {
                    tyyy.Add(item.tbUsers);
                }
                //var tyyy = db.tbUsers.ToList();
                var caran = db.tbCaranSettings.ToList();
                var asnad = db.tbfkfinancial.ToList();
                var FinancialDocuments = db.FinancialDocuments.ToList();

                List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
                foreach (var it in asnad)
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                    int persianYear = persianCalendar.GetYear(dataDocument);
                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                    if (persianYear == Year && persianMonth == Month)
                    {
                        asnadstr.Add(it);

                    }
                }
                var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

                var y = db.tbMoalefeValueFish
                            .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month)
                            .ToList();
                int counter = 1;
                int counter2 = 2;






                var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
                var t = db.tbContractMoalefeDastmozdi.ToList();
                var count = 10;
                foreach (var item in MoalefeID)
                {
                    var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                foreach (var item in asnadstr)
                {

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }






                foreach (var item in tyyy)
                {

                    string job = "";
                    if (item.tbjob != null)
                    {
                        job = item.tbjob.Name;
                    }
                    bool amani = false;
                    if (item.usr_amani == true)
                    {
                        amani = true;
                    }
                    var tb = noalf
    .Where(p => p.MoalfeVal_FKUser == item.usr_ID

                && p.MoalfeVal_Value != 0
                )
    .ToList();
                    List<string> mode = new List<string>();

                    foreach (var it in tb)
                    {
                        var yuuu = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                            .FirstOrDefault();

                        if (yuuu == null)
                        {
                            var u = caran
                                .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                                .FirstOrDefault();

                            if (u != null)
                            {
                                mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{it.MoalfeVal_Value} )----");
                            }
                        }
                        else
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
                        }
                    }
                    double moalde = 0;
                    double clock = 0;
                    var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                    if (moad != null && moad.Moadel != null)
                    {
                        moalde = moad.Moadel;
                    }
                    var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                    if (exi != null && exi.RemainValue != null)
                    {
                        clock = (double)exi.RemainValue;
                    }
                    if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                    {

                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                   ,


                                                   new Cell()
                        {
                            Value = Month,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        }
                                                           ,


                                                   new Cell()
                        {
                            Value = Year,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 9
                        }
                    });


                            //مقادیر مولفه ها
                            int index = 10;
                            foreach (var item2 in MoalefeID)
                            {
                                var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                double ty = 0;
                                if (exx != null)
                                {
                                    ty = (double)exx.mlfvlfsh_Value;
                                }
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            foreach (var item2 in asnadstr)
                            {
                                long value = 0;
                                var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                                foreach (var item1 in t3)
                                {
                                    if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                    {
                                        PersianCalendar persianCalendar = new PersianCalendar();
                                        DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                        int persianYear = persianCalendar.GetYear(dataDocument);
                                        int persianMonth = persianCalendar.GetMonth(dataDocument);
                                        if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                        {
                                            if (item1.Debtore != 0 && item1.Creditor != 0)
                                            {
                                                value = (long)(item1.Creditor - item1.Debtore);
                                            }


                                            else if (item1.Debtore != 0 && item1.Debtore != null)
                                            {
                                                value = (long)(item1.Debtore ?? 0);
                                            }
                                            else if (item1.Creditor != 0 && item1.Creditor != null)
                                            {
                                                value = (long)(item1.Creditor ?? 0);

                                            }

                                        }

                                    }
                                }

                                //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                //double ty = 0;
                                //if (exx != null)
                                //{
                                //    ty = (double)exx.mlfvlfsh_Value;
                                //}
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            counter++;
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                    //نام ، نام خانوادگی و کد پرسنلی





                }



























            }

            return Moalefeexcelfile;
        }
        //public string ExportMoalefeCsv5()
        //{
        //    int year = 1404;
        //    string folderPath = Server.MapPath("~/Content/ExcelFiles/");
        //    if (!Directory.Exists(folderPath))
        //        Directory.CreateDirectory(folderPath);

        //    string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        //    string csvPath = Path.Combine(folderPath, fileName);

        //    var months = Enumerable.Range(1, 8).ToList();

        //    var allowedMoalefeIds = db.tbMoalefeDastmozdiValueFromExcel
        //        .Where(s => s.MoalfeVal_Year == year)
        //        .Select(s => s.MoalfeVal_FKMoalafeDastmozdi.Value)
        //        .Distinct()
        //        .ToList();

        //    var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
        //        .Where(p => p.MoalfeVal_Year == year &&
        //                    allowedMoalefeIds.Contains((int)p.MoalfeVal_FKMoalafeDastmozdi))
        //        .ToList();

        //    var moalefeLookup = moalefeValues.ToLookup(v => new
        //    {
        //        UserID = v.MoalfeVal_FKUser,
        //        MoalefeID = v.MoalfeVal_FKMoalafeDastmozdi,
        //        Month = v.MoalfeVal_Month
        //    });

        //    var moalefeContracts = db.tbContractMoalefeDastmozdi
        //        .Where(p => allowedMoalefeIds.Contains((int)p.md_ID))
        //        .ToList();

        //    var caranSettingsDict = db.tbCaranSettings
        //        .Where(c => allowedMoalefeIds.Contains((int)c.FK_Moalefe_ID))
        //        .ToDictionary(c => c.FK_Moalefe_ID.Value);

        //    var moadDict = db.tbMoadelPadashJarimeAyab
        //        .Where(p => p.Year == year)
        //        .ToLookup(p => new { UserID = (int?)p.UserID, Month = (int?)p.Month });

        //    var cityDict = db.tbCities.ToDictionary(c => c.ID, c => c.Name);

        //    using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        //    {
        //        foreach (var month in months)
        //        {
        //            // ▬▬▬▬▬▬▬  Row 1 - AyabOZahab ▬▬▬▬▬▬▬
        //            var rowAyab = new List<string> { "", "", "", "", "", "", "" };

        //            foreach (var contract in moalefeContracts)
        //            {
        //                caranSettingsDict.TryGetValue(contract.md_ID, out var caran);
        //                rowAyab.Add(caran?.CaranAyabOZahab?.ToString() ?? "0");
        //            }
        //            writer.WriteLine(string.Join(",", rowAyab));

        //            // ▬▬▬▬▬▬▬  Row 2 - Standard ▬▬▬▬▬▬▬
        //            var rowStandard = new List<string> { "", "", "", "", "", "", "" };

        //            foreach (var contract in moalefeContracts)
        //            {
        //                caranSettingsDict.TryGetValue(contract.md_ID, out var caran);
        //                rowStandard.Add(caran?.CaranStandard?.ToString() ?? "0");
        //            }
        //            writer.WriteLine(string.Join(",", rowStandard));

        //            // ▬▬▬▬▬▬▬  Row 3 - Column Titles ▬▬▬▬▬▬▬
        //            var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde" };

        //            foreach (var contract in moalefeContracts)
        //                header.Add(contract.md_Title);

        //            writer.WriteLine(string.Join(",", header));

        //            // کاربران همان ماه
        //            var usersInMonth = moalefeValues
        //                .Where(p => p.MoalfeVal_Month == month)
        //                .Select(p => p.tbUsers)
        //                .Distinct()
        //                .ToList();

        //            foreach (var user in usersInMonth)
        //            {
        //                cityDict.TryGetValue((int)user.usr_City_Dutysystem, out string cityName);

        //                var row = new List<string>
        //        {
        //            user.FullName,
        //            user.usr_Personal_ID.ToString(),
        //            user.tbjob?.Name ?? "",
        //            cityName ?? "",
        //            month.ToString(),
        //            year.ToString()
        //        };

        //                // Moalde
        //                var moadKey = new { UserID = (int?)user.usr_ID, Month = (int?)month };
        //                var moad = moadDict[moadKey].FirstOrDefault();
        //                double moalde = moad?.Moadel ?? 0;
        //                row.Add(moalde.ToString());

        //                // مقادیر مؤلفه‌ها
        //                foreach (var contract in moalefeContracts)
        //                {
        //                    var key = new { UserID = user.usr_ID, MoalefeID = contract.md_ID, Month = month };
        //                    var val = moalefeLookup[key].FirstOrDefault()?.MoalfeVal_Value ?? 0;
        //                    row.Add(val.ToString());
        //                }

        //                writer.WriteLine(string.Join(",", row));
        //            }
        //        }
        //    }

        //    return csvPath;
        //}
        public string ExportMoalefeCsv6()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 9).ToList();

            // Load data
            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year)
                .ToList();

            var allowedIds = moalefeValues
                .Select(p => p.MoalfeVal_FKMoalafeDastmozdi ?? 0)
                .Distinct()
                .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi
                .Where(p => allowedIds.Contains(p.md_ID))
                .Select(p => new { p.md_ID, p.md_Title })
                .ToList();

            var caranDict = db.tbCaranSettings
                .Where(c => allowedIds.Contains((int)c.FK_Moalefe_ID))
                .ToDictionary(c => c.FK_Moalefe_ID, c => c);

            var usersDict = db.tbUsers.ToDictionary(u => u.usr_ID);
            var citiesDict = db.tbCities.ToDictionary(c => c.ID);

            var moadDict = db.tbMoadelPadashJarimeAyab
              .Where(m => m.Year == year)
              .ToList()
              .GroupBy(m => Tuple.Create(m.UserID.Value, m.Month.Value))
              .ToDictionary(g => g.Key, g => g.First().Moadel);



            // very fast lookup ready
            var valueLookup = moalefeValues
                .GroupBy(v => Tuple.Create(
                    v.MoalfeVal_FKUser ?? 0,
                    v.MoalfeVal_FKMoalafeDastmozdi ?? 0,
                    v.MoalfeVal_Month ?? 0
                ))
                .ToDictionary(g => g.Key, g => g.First().MoalfeVal_Value ?? 0);

            var utf8NoBom = new UTF8Encoding(false);

            using (var writer = new StreamWriter(csvPath, false, utf8NoBom))
            {
                foreach (var month in months)
                {
                    // Row 1 — Ayab
                    writer.WriteLine(
                        string.Join(",", Enumerable.Repeat("", 7)
                        .Concat(moalefeContracts.Select(c =>
                            caranDict.TryGetValue(c.md_ID, out var set)
                                ? (set.CaranAyabOZahab ?? 0).ToString()
                                : "0"
                        )))
                    );

                    // Row 2 — Standard
                    writer.WriteLine(
                        string.Join(",", Enumerable.Repeat("", 7)
                        .Concat(moalefeContracts.Select(c =>
                            caranDict.TryGetValue(c.md_ID, out var set)
                                ? (set.CaranStandard ?? 0).ToString()
                                : "0"
                        )))
                    );

                    // Header
                    writer.WriteLine(
                        string.Join(",", new[]
                        {
                    "FullName","PersonalID","Job","City","Month","Year","Moalde"
                        }.Concat(moalefeContracts.Select(c => c.md_Title)))
                    );

                    // Get users once per month
                    var usersInMonth = moalefeValues
                        .Where(v => v.MoalfeVal_Month == month)
                        .Select(v => v.MoalfeVal_FKUser ?? 0)
                        .Distinct()
                        .ToList();

                    foreach (var uid in usersInMonth)
                    {
                        if (!usersDict.TryGetValue(uid, out var user)) continue;

                        var cityName = citiesDict.TryGetValue(user.usr_City_Dutysystem ?? 0, out var ct) ? ct.Name : "";

                        moadDict.TryGetValue(Tuple.Create(uid, month), out double moalde);

                        var row = new List<string>
                {
                    user.FullName,
                    user.usr_Personal_ID?.ToString() ?? "",
                    user.tbjob?.Name ?? "",
                    cityName,
                    month.ToString(),
                    year.ToString(),
                    moalde.ToString()
                };

                        foreach (var c in moalefeContracts)
                        {
                            double val = 0;
                            valueLookup.TryGetValue(Tuple.Create(uid, c.md_ID, month), out val);
                            row.Add(val.ToString());
                        }


                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }
        public string ExportMoalefeCsv7()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 9).ToList();

            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year)
                .ToList();

            var allowedMoalefeIds = moalefeValues
                .Select(p => p.MoalfeVal_FKMoalafeDastmozdi.Value)
                .Distinct()
                .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi
                .Where(p => allowedMoalefeIds.Contains(p.md_ID))
                .ToList();

            var caranSettings = db.tbCaranSettings.ToDictionary(c => c.FK_Moalefe_ID);
            var users = db.tbUsers.ToDictionary(c => c.usr_ID);
            var cities = db.tbCities.ToDictionary(c => c.ID);
            var moadData = db.tbMoadelPadashJarimeAyab
                .Where(p => p.Year == year)
                .ToList()
                .GroupBy(p => (p.UserID, p.Month))
                .ToDictionary(g => g.Key, g => g.First().Moadel);

            // Lookup برای مقادیر کاربران
            var valueLookup = moalefeValues
                .GroupBy(v => (v.MoalfeVal_FKUser, v.MoalfeVal_FKMoalafeDastmozdi, v.MoalfeVal_Month))
                .ToDictionary(g => g.Key, g => g.First().MoalfeVal_Value);

            int baseColumns = 7; // FullName, PersonalID, Job, City, Month, Year, Moalde
            int contractColumns = moalefeContracts.Count;
            int totalColumns = baseColumns + contractColumns;

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // Row 1 AyabOZahab
                    var rowAyab = new List<string>(new string[totalColumns]);
                    for (int i = 0; i < moalefeContracts.Count; i++)
                    {
                        var contract = moalefeContracts[i];
                        if (caranSettings.TryGetValue(contract.md_ID, out var caran))
                            rowAyab[baseColumns + i] = caran.CaranAyabOZahab?.ToString() ?? "0";
                        else
                            rowAyab[baseColumns + i] = "0";
                    }
                    writer.WriteLine(string.Join(",", rowAyab));

                    // Row 2 Standard
                    var rowStandard = new List<string>(new string[totalColumns]);
                    for (int i = 0; i < moalefeContracts.Count; i++)
                    {
                        var contract = moalefeContracts[i];
                        if (caranSettings.TryGetValue(contract.md_ID, out var caran))
                            rowStandard[baseColumns + i] = caran.CaranStandard?.ToString() ?? "0";
                        else
                            rowStandard[baseColumns + i] = "0";
                    }
                    writer.WriteLine(string.Join(",", rowStandard));

                    // Header line
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde" };
                    header.AddRange(moalefeContracts.Select(c => c.md_Title));
                    writer.WriteLine(string.Join(",", header));

                    // Users in month
                    var usersInMonth = moalefeValues
                        .Where(v => v.MoalfeVal_Month == month)
                        .Select(v => v.MoalfeVal_FKUser.Value)
                        .Distinct()
                        .ToList();

                    foreach (var userId in usersInMonth)
                    {
                        var user = users[userId];
                        var cityName = cities.TryGetValue((int)user.usr_City_Dutysystem, out var city) ? city.Name : "";

                        var row = new List<string>(new string[totalColumns])
                        {
                            [0] = user.FullName,
                            [1] = user.usr_Personal_ID.ToString(),
                            [2] = user.tbjob?.Name ?? "",
                            [3] = cityName,
                            [4] = month.ToString(),
                            [5] = year.ToString()
                        };

                        moadData.TryGetValue((userId, month), out double moalde);
                        row[6] = moalde.ToString();

                        for (int i = 0; i < moalefeContracts.Count; i++)
                        {
                            var contract = moalefeContracts[i];
                            valueLookup.TryGetValue((userId, contract.md_ID, month), out double? value);
                            row[baseColumns + i] = (value ?? 0).ToString();
                        }

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }

        public string ExportMoalefeCsv_Final()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 9).ToList();

            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year)
                .ToList();

            var allowedMoalefeIds = moalefeValues
                .Select(p => p.MoalfeVal_FKMoalafeDastmozdi ?? 0)
                .Distinct()
                .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi
                .Where(p => allowedMoalefeIds.Contains(p.md_ID))
                .ToList();

            var caranSettings = db.tbCaranSettings.ToDictionary(c => c.FK_Moalefe_ID);
            var usersDict = db.tbUsers.ToDictionary(u => u.usr_ID);
            var citiesDict = db.tbCities.ToDictionary(c => c.ID);

            var moadDict = db.tbMoadelPadashJarimeAyab
    .Where(m => m.Year == year)
    .ToList()
    .GroupBy(m => Tuple.Create(m.UserID.Value, m.Month.Value))
    .ToDictionary(g => g.Key, g => g.First().Moadel);

            var valueLookup = moalefeValues
                .GroupBy(v => Tuple.Create(
                    v.MoalfeVal_FKUser ?? 0,
                    v.MoalfeVal_FKMoalafeDastmozdi ?? 0,
                    v.MoalfeVal_Month ?? 0
                ))
                .ToDictionary(g => g.Key, g => g.First().MoalfeVal_Value ?? 0);

            // -------- CSV SAFE QUOTE 🔥
            string Q(object x) => $"\"{(x?.ToString() ?? "").Replace("\"", "\"\"")}\"";

            // -------- UTF8 with BOM 👌 مهم برای فارسی
            var utf8Bom = new UTF8Encoding(true);

            using (var writer = new StreamWriter(csvPath, false, utf8Bom))
            {
                foreach (var month in months)
                {
                    // Row 1 Ayab
                    var rowAyab = Enumerable.Repeat("", 7).Select(Q).ToList();
                    rowAyab.AddRange(moalefeContracts.Select(c =>
                        caranSettings.TryGetValue(c.md_ID, out var cs)
                            ? Q(cs.CaranAyabOZahab ?? 0)
                            : Q(0)));
                    writer.WriteLine(string.Join(",", rowAyab));

                    // Row 2 Standard
                    var rowStd = Enumerable.Repeat("", 7).Select(Q).ToList();
                    rowStd.AddRange(moalefeContracts.Select(c =>
                        caranSettings.TryGetValue(c.md_ID, out var cs)
                            ? Q(cs.CaranStandard ?? 0)
                            : Q(0)));
                    writer.WriteLine(string.Join(",", rowStd));

                    // Header
                    var header = new List<string>
            {
                Q("FullName"), Q("PersonalID"), Q("Job"), Q("City"), Q("Month"), Q("Year"), Q("Moalde")
            };
                    header.AddRange(moalefeContracts.Select(c => Q(c.md_Title)));
                    writer.WriteLine(string.Join(",", header));

                    // Users
                    var usersInMonth = moalefeValues
                        .Where(v => v.MoalfeVal_Month == month)
                        .Select(v => v.MoalfeVal_FKUser ?? 0)
                        .Distinct()
                        .ToList();

                    foreach (var uid in usersInMonth)
                    {
                        if (!usersDict.TryGetValue(uid, out var user)) continue;

                        var cityName = citiesDict.TryGetValue(user.usr_City_Dutysystem ?? 0, out var ct) ? ct.Name : "";

                        moadDict.TryGetValue(Tuple.Create(uid, month), out double moalde);

                        var row = new List<string>
                {
                    Q(user.FullName),
                    Q(user.usr_Personal_ID),
                    Q(user.tbjob?.Name),
                    Q(cityName),
                    Q(month),
                    Q(year),
                    Q(moalde)
                };

                        foreach (var c in moalefeContracts)
                        {
                            valueLookup.TryGetValue(Tuple.Create(uid, c.md_ID, month), out double val);
                            row.Add(Q(val));
                        }

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }

        public string ExportMoalefeCsv5()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 9).ToList();

            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year)
                .ToList();

            var allowedMoalefeIds = moalefeValues
                .Select(p => p.MoalfeVal_FKMoalafeDastmozdi.Value)
                .Distinct()
                .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi
                .Where(p => allowedMoalefeIds.Contains(p.md_ID))
                .ToList();

            var caranSettings = db.tbCaranSettings.ToDictionary(c => c.FK_Moalefe_ID);
            var users = db.tbUsers.ToDictionary(c => c.usr_ID);
            var cities = db.tbCities.ToDictionary(c => c.ID);
            var moadData = db.tbMoadelPadashJarimeAyab
                .Where(p => p.Year == year)
                .ToList()
                .GroupBy(p => (p.UserID, p.Month))
                .ToDictionary(g => g.Key, g => g.First().Moadel);

            // ساخت Lookup برای دسترسی سریع به مقادیر کاربران برای هر مؤلفه و ماه
            var valueLookup = moalefeValues
                .GroupBy(v => (v.MoalfeVal_FKUser, v.MoalfeVal_FKMoalafeDastmozdi, v.MoalfeVal_Month))
                .ToDictionary(g => g.Key, g => g.First().MoalfeVal_Value);

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // Row 1 AyabOZahab
                    var rowAyab = Enumerable.Repeat("", 7).ToList();
                    rowAyab.AddRange(moalefeContracts.Select(c =>
                        caranSettings.TryGetValue(c.md_ID, out var caran) ?
                        caran.CaranAyabOZahab?.ToString() ?? "0" : "0"));
                    writer.WriteLine(string.Join(",", rowAyab));

                    // Row 2 Standard
                    var rowStandard = Enumerable.Repeat("", 7).ToList();
                    rowStandard.AddRange(moalefeContracts.Select(c =>
                        caranSettings.TryGetValue(c.md_ID, out var caran) ?
                        caran.CaranStandard?.ToString() ?? "0" : "0"));
                    writer.WriteLine(string.Join(",", rowStandard));

                    // Header line
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde" };
                    header.AddRange(moalefeContracts.Select(c => c.md_Title));
                    writer.WriteLine(string.Join(",", header));

                    // Users in month
                    var usersInMonth = moalefeValues
                        .Where(v => v.MoalfeVal_Month == month)
                        .Select(v => v.MoalfeVal_FKUser.Value)
                        .Distinct()
                        .ToList();

                    foreach (var userId in usersInMonth)
                    {
                        var user = users[userId];
                        var cityName = cities.TryGetValue((int)user.usr_City_Dutysystem, out var city) ? city.Name : "";

                        var row = new List<string>
                {
                    user.FullName,
                    user.usr_Personal_ID.ToString(),
                    user.tbjob?.Name ?? "",
                    cityName,
                    month.ToString(),
                    year.ToString()
                };

                        moadData.TryGetValue((userId, month), out double moalde);
                        row.Add(moalde.ToString());

                        foreach (var contract in moalefeContracts)
                        {
                            valueLookup.TryGetValue((userId, contract.md_ID, month), out double? value);
                            row.Add((value ?? 0).ToString());
                        }


                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }



        public string ExportMoalefeCsv4()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 9).ToList();

            var allowedMoalefeIds = db.tbMoalefeDastmozdiValueFromExcel
                .Where(s => s.MoalfeVal_Year == year)
                .GroupBy(s => s.MoalfeVal_FKMoalafeDastmozdi)
                .Select(g => g.Key)
                .ToList();

            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year &&
                            allowedMoalefeIds.Contains((int)p.MoalfeVal_FKMoalafeDastmozdi))
                .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi
                .Where(p => allowedMoalefeIds.Contains((int)p.md_ID))
                .ToList();

            var caranSettings = db.tbCaranSettings.ToList();
            var users = db.tbUsers.ToList();
            var maxkarkardMonth = db.tbMaxkarkardMonth.ToList();
            var moadData = db.tbMoadelPadashJarimeAyab.ToList();
            var cities = db.tbCities.ToList();

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // ردیف 1 - مقادیر AyabOZahab
                    var rowAyab = new List<string> { "", "", "", "", "", "", "" };
                    foreach (var contract in moalefeContracts)
                    {
                        var caran = caranSettings.FirstOrDefault(c => c.FK_Moalefe_ID == contract.md_ID);
                        rowAyab.Add(caran?.CaranAyabOZahab?.ToString() ?? "0");
                    }
                    writer.WriteLine(string.Join(",", rowAyab));

                    // ردیف 2 - مقادیر Standard
                    var rowStandard = new List<string> { "", "", "", "", "", "", "" };
                    foreach (var contract in moalefeContracts)
                    {
                        var caran = caranSettings.FirstOrDefault(c => c.FK_Moalefe_ID == contract.md_ID);
                        rowStandard.Add(caran?.CaranStandard?.ToString() ?? "0");
                    }
                    writer.WriteLine(string.Join(",", rowStandard));

                    // ردیف 3 - نام مؤلفه‌ها
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde" };
                    foreach (var contract in moalefeContracts)
                        header.Add(contract.md_Title);
                    writer.WriteLine(string.Join(",", header));

                    // لیست کاربران همان ماه
                    var usersInMonth = moalefeValues
                        .Where(p => p.MoalfeVal_Month == month)
                        .Select(p => p.tbUsers)
                        .Distinct()
                        .ToList();

                    foreach (var user in usersInMonth)
                    {
                        var row = new List<string>
                {
                    user.FullName,
                    user.usr_Personal_ID.ToString(),
                    user.tbjob?.Name ?? "",
                    cities.Where(c => c.ID == user.usr_City_Dutysystem).Select(c => c.Name).FirstOrDefault() ?? "",
                    month.ToString(),
                    year.ToString()
                };

                        // Moalde
                        var moad = moadData.FirstOrDefault(p =>
                            p.UserID == user.usr_ID && p.Month == month && p.Year == year);
                        double moalde = moad?.Moadel ?? 0;
                        row.Add(moalde.ToString());

                        // مقادیر مؤلفه‌ها
                        foreach (var contract in moalefeContracts)
                        {
                            var val = moalefeValues
                                .Where(v => v.MoalfeVal_FKUser == user.usr_ID &&
                                            v.MoalfeVal_FKMoalafeDastmozdi == contract.md_ID &&
                                            v.MoalfeVal_Month == month)
                                .Select(v => v.MoalfeVal_Value)
                                .FirstOrDefault();

                            row.Add(val?.ToString() ?? "0");
                        }

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }


        public string ExportMoalefeCsv3()
        {
            int year = 1404;
            // مسیر پوشه CSV
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // نام فایل CSV با تاریخ و ساعت برای یکتا بودن
            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 8).ToList(); // ماه‌ها 1 تا 8

            // بارگذاری داده‌ها یک‌بار از دیتابیس
            //var moalefeValues = db.tbMoalefeValueFish
            //                      .Where(p => p.mlfvlfsh_Year == year && p.FK_EXCel == null)
            //                      .ToList();
            //var allowedMoalefeIds = new List<int> { 1953, 1981, 1964, 1965, 1973, 1982, 1961, 1979 };
            var allowedMoalefeIds = db.tbMoalefeDastmozdiValueFromExcel
    .Where(s => s.MoalfeVal_Year == year)
    .GroupBy(s => s.MoalfeVal_FKMoalafeDastmozdi)
    .Select(g => g.Key)
    .ToList();

            var moalefeValues = db.tbMoalefeDastmozdiValueFromExcel
                .Where(p => p.MoalfeVal_Year == year
                            && allowedMoalefeIds.Contains((int)p.MoalfeVal_FKMoalafeDastmozdi))
                .ToList();

            //var specialMoalefeIds = new List<int> { 1953, 1964, 1965, 5637 };
            //var moalefeValuesSpecial = db.tbMoalefeValueFish
            //                            .Where(p => p.mlfvlfsh_Year == year
            //                                        && p.Finalaccept == true
            //                                        && specialMoalefeIds.Contains((int)p.FK_Moalefe))
            //                            .ToList();
            var moalefeContracts = db.tbContractMoalefeDastmozdi.Where(p => allowedMoalefeIds.Contains((int)p.md_ID)).ToList();
            //var moalefeContractsspical = db.tbContractMoalefeDastmozdi.Where(p => specialMoalefeIds.Contains((int)p.md_ID)).ToList();
            var moalefeContractsstandard = db.tbContractMoalefeDastmozdi.Where(p => allowedMoalefeIds.Contains((int)p.md_ID)).ToList();

            var financialDocs = db.FinancialDocuments.ToList();
            var caranSettings = db.tbCaranSettings.ToList();
            var users = db.tbUsers.ToList();
            var maxkarkardMonth = db.tbMaxkarkardMonth.ToList();
            var moadData = db.tbMoadelPadashJarimeAyab.ToList();
            var allAsnad = db.tbfkfinancial
                             .Where(f => f.DataDocument.HasValue)
                             .ToList();

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // فیلتر Asnad روی حافظه با PersianCalendar
                    var asnadstr = allAsnad
                        .Where(f =>
                        {
                            var pc = new PersianCalendar();
                            var docDate = f.DataDocument.Value;
                            return pc.GetYear(docDate) == year && pc.GetMonth(docDate) == month;
                        })
                        .ToList();

                    // ایجاد هدر CSV
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde", "Clock" };
                    foreach (var contract in moalefeContracts)
                        header.Add(contract.md_Title);
                    //foreach (var contract in moalefeContractsspical)
                    //    header.Add(contract.md_Title);
                    //foreach (var asnad in asnadstr)
                    //    header.Add(asnad.Title);

                    writer.WriteLine(string.Join(",", header));

                    // کاربران ماه جاری
                    var usersInMonth = moalefeValues
                                       .Where(p => p.MoalfeVal_Month == month)
                                       .Select(p => p.tbUsers)
                                       .Distinct()
                                       .ToList();

                    foreach (var user in usersInMonth)
                    {
                        var row = new List<string>
                {
                    user.FullName,
                    user.usr_Personal_ID.ToString(),
                    user.tbjob?.Name ?? "",
                    db.tbCities.Where(c => c.ID == user.usr_City_Dutysystem).Select(c => c.Name).FirstOrDefault() ?? "",
                    month.ToString(),
                    year.ToString()
                };

                        // مقادیر moalde و clock
                        double moalde = 0;
                        double clock = 0;

                        var moad = moadData
                            .Where(p => p.UserID == user.usr_ID && p.Month == month && p.Year == year)
                            .FirstOrDefault();
                        if (moad != null && moad.Moadel != null)
                            moalde = moad.Moadel;

                        var clockData = maxkarkardMonth
                            .Where(p => p.FKUser == user.usr_ID && p.Month == month && p.Year == year)
                            .FirstOrDefault();
                        if (clockData != null && clockData.RemainValue != null)
                            clock = (double)clockData.RemainValue;

                        row.Add(moalde.ToString());
                        //row.Add(clock.ToString());

                        // مقادیر Moalefe
                        foreach (var contract in moalefeContracts)
                        {
                            var val = moalefeValues
                                        .Where(v => v.MoalfeVal_FKUser == user.usr_ID && v.MoalfeVal_FKMoalafeDastmozdi == contract.md_ID && v.MoalfeVal_FKMoalafeDastmozdi == month)
                                        .Select(v => v.MoalfeVal_Value)
                                        .FirstOrDefault();
                            row.Add(val?.ToString() ?? "0");
                        }
                        //foreach (var contract in moalefeContractsspical)
                        //{
                        //    var val = moalefeValuesSpecial
                        //                .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == contract.md_ID && v.mlfvlfsh_Month == month)
                        //                .Select(v => v.mlfvlfsh_Value)
                        //                .FirstOrDefault();
                        //    row.Add(val?.ToString() ?? "0");
                        //}
                        // مقادیر Asnad
                        //foreach (var asnad in asnadstr)
                        //{
                        //    long value = 0;

                        //    var docs = financialDocs
                        //        .Where(f => f.User_ID == user.usr_ID && f.FK_final == asnad.ID && f.DataDocument.HasValue)
                        //        .ToList();

                        //    foreach (var doc in docs)
                        //    {
                        //        var pc = new PersianCalendar();
                        //        int docYear = pc.GetYear(doc.DataDocument.Value);
                        //        int docMonth = pc.GetMonth(doc.DataDocument.Value);

                        //        if (docYear == year && docMonth == month)
                        //        {
                        //            if (doc.Creditor != 0 && doc.Debtore != 0)
                        //                value = (long)(doc.Creditor - doc.Debtore);
                        //            else if (doc.Debtore != 0 && doc.Debtore != null)
                        //                value = (long)(doc.Debtore ?? 0);
                        //            else if (doc.Creditor != 0 && doc.Creditor != null)
                        //                value = (long)(doc.Creditor ?? 0);
                        //        }
                        //    }

                        //    row.Add(value.ToString());
                        //}

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }


        public string ExportMoalefeCsv()
        {
            int year = 1404;
            // مسیر پوشه CSV
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // نام فایل CSV با تاریخ و ساعت برای یکتا بودن
            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 8).ToList(); // ماه‌ها 1 تا 8

            // بارگذاری داده‌ها یک‌بار از دیتابیس
            //var moalefeValues = db.tbMoalefeValueFish
            //                      .Where(p => p.mlfvlfsh_Year == year && p.FK_EXCel == null)
            //                      .ToList();
            var allowedMoalefeIds = new List<int> { 1960,1953, 1981, 1964, 1965, 1973, 1982, 1961,1979 };

            var moalefeValues = db.tbMoalefeValueFish
                                  .Where(p => p.mlfvlfsh_Year == year
                                              && p.FK_EXCel==null
                                              && allowedMoalefeIds.Contains((int)p.FK_Moalefe))
                                  .ToList();
            var specialMoalefeIds = new List<int> { 1953, 1964, 1965, 5637 };
            var moalefeValuesSpecial = db.tbMoalefeValueFish
                                        .Where(p => p.mlfvlfsh_Year == year
                                                    && p.Finalaccept == true
                                                    && specialMoalefeIds.Contains((int)p.FK_Moalefe))
                                        .ToList();
            var moalefeContracts = db.tbContractMoalefeDastmozdi.Where(p=> allowedMoalefeIds.Contains((int)p.md_ID)).ToList();
            var moalefeContractsspical = db.tbContractMoalefeDastmozdi.Where(p => specialMoalefeIds.Contains((int)p.md_ID)).ToList();

            var financialDocs = db.FinancialDocuments.ToList();
            var caranSettings = db.tbCaranSettings.ToList();
            var users = db.tbUsers.ToList();
            var maxkarkardMonth = db.tbMaxkarkardMonth.ToList();
            var moadData = db.tbMoadelPadashJarimeAyab.ToList();
            var allAsnad = db.tbfkfinancial
                             .Where(f => f.DataDocument.HasValue)
                             .ToList();

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // فیلتر Asnad روی حافظه با PersianCalendar
                    var asnadstr = allAsnad
                        .Where(f =>
                        {
                            var pc = new PersianCalendar();
                            var docDate = f.DataDocument.Value;
                            return pc.GetYear(docDate) == year && pc.GetMonth(docDate) == month;
                        })
                        .ToList();

                    // ایجاد هدر CSV
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde", "Clock" };
                    foreach (var contract in moalefeContracts)
                        header.Add(contract.md_Title);
                    foreach (var contract in moalefeContractsspical)
                        header.Add(contract.md_Title);
                    foreach (var asnad in asnadstr)
                        header.Add(asnad.Title);

                    writer.WriteLine(string.Join(",", header));

                    // کاربران ماه جاری
                    var usersInMonth = moalefeValues
                                       .Where(p => p.mlfvlfsh_Month == month &&p.FK_Moalefe==1979)
                                       .Select(p => p.tbUsers)
                                       .Distinct()
                                       .ToList();

                    foreach (var user in usersInMonth)
                    {
                        var row = new List<string>
                {
                    user.FullName,
                    user.usr_Personal_ID.ToString(),
                    user.tbjob?.Name ?? "",
                    db.tbCities.Where(c => c.ID == user.usr_City_Dutysystem).Select(c => c.Name).FirstOrDefault() ?? "",
                    month.ToString(),
                    year.ToString()
                };

                        // مقادیر moalde و clock
                        double moalde = 0;
                        double clock = 0;

                        var moad = moadData
                            .Where(p => p.UserID == user.usr_ID && p.Month == month && p.Year == year)
                            .FirstOrDefault();
                        if (moad != null && moad.Moadel != null)
                            moalde = moad.Moadel;

                        var clockData = maxkarkardMonth
                            .Where(p => p.FKUser == user.usr_ID && p.Month == month && p.Year == year)
                            .FirstOrDefault();
                        if (clockData != null && clockData.RemainValue != null)
                            clock = (double)clockData.RemainValue;

                        row.Add(moalde.ToString());
                        row.Add(clock.ToString());

                        // مقادیر Moalefe
                        foreach (var contract in moalefeContracts)
                        {
                            var val = moalefeValues
                                        .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == contract.md_ID && v.mlfvlfsh_Month == month)
                                        .Select(v => v.mlfvlfsh_Value)
                                        .FirstOrDefault();
                            row.Add(val?.ToString() ?? "0");
                        }
                        foreach (var contract in moalefeContractsspical)
                        {
                            var val = moalefeValuesSpecial
                                        .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == contract.md_ID && v.mlfvlfsh_Month == month)
                                        .Select(v => v.mlfvlfsh_Value)
                                        .FirstOrDefault();
                            row.Add(val?.ToString() ?? "0");
                        }
                        // مقادیر Asnad
                        foreach (var asnad in asnadstr)
                        {
                            long value = 0;

                            var docs = financialDocs
                                .Where(f => f.User_ID == user.usr_ID && f.FK_final == asnad.ID && f.DataDocument.HasValue)
                                .ToList();

                            foreach (var doc in docs)
                            {
                                var pc = new PersianCalendar();
                                int docYear = pc.GetYear(doc.DataDocument.Value);
                                int docMonth = pc.GetMonth(doc.DataDocument.Value);

                                if (docYear == year && docMonth == month)
                                {
                                    if (doc.Creditor != 0 && doc.Debtore != 0)
                                        value = (long)(doc.Creditor - doc.Debtore);
                                    else if (doc.Debtore != 0 && doc.Debtore != null)
                                        value = (long)(doc.Debtore ?? 0);
                                    else if (doc.Creditor != 0 && doc.Creditor != null)
                                        value = (long)(doc.Creditor ?? 0);
                                }
                            }

                            row.Add(value.ToString());
                        }

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }


        public string ExportMoalefeCsv2()
        {
            int year = 1404;
            string folderPath = Server.MapPath("~/Content/ExcelFiles/");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string fileName = $"MoalefeOutput_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string csvPath = Path.Combine(folderPath, fileName);

            var months = Enumerable.Range(1, 8).ToList(); // ماه‌ها 1 تا 8

            var allowedMoalefeIds = new List<int> { 1960,1953, 1981, 1964, 1965, 1973, 1982, 1961, 1979 };
            var specialMoalefeIds = new List<int> { 1953, 1964, 1965, 5637 };

            var moalefeValues = db.tbMoalefeValueFish
                                  .Where(p => p.mlfvlfsh_Year == year && p.FK_EXCel == null && allowedMoalefeIds.Contains((int)p.FK_Moalefe))
                                  .ToList();

            var moalefeValuesSpecial = db.tbMoalefeValueFish
                                         .Where(p => p.mlfvlfsh_Year == year && p.Finalaccept == true && specialMoalefeIds.Contains((int)p.FK_Moalefe))
                                         .ToList();

            var moalefeContracts = db.tbContractMoalefeDastmozdi.Where(p => allowedMoalefeIds.Contains((int)p.md_ID)).ToList();
            var moalefeContractsspical = db.tbContractMoalefeDastmozdi.Where(p => specialMoalefeIds.Contains((int)p.md_ID)).ToList();

            var financialDocs = db.FinancialDocuments.Where(s=>s.tbfkfinancial.numbershomar==5001).ToList();
            var maxkarkardMonth = db.tbMaxkarkardMonth.ToList();
            var moadData = db.tbMoadelPadashJarimeAyab.ToList();
            var allAsnad = db.tbfkfinancial.Where(f =>f.numbershomar==5001&& f.DataDocument.HasValue).ToList();

            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                foreach (var month in months)
                {
                    // فیلتر Asnad برای ماه جاری
                    var asnadstr = allAsnad
                        .Where(f =>
                        {
                            var pc = new PersianCalendar();
                            var docDate = f.DataDocument.Value;
                            return pc.GetYear(docDate) == year && pc.GetMonth(docDate) == month;
                        })
                        .ToList();

                    // هدر CSV
                    var header = new List<string> { "FullName", "PersonalID", "Job", "City", "Month", "Year", "Moalde", "Clock" };
                    header.AddRange(moalefeContracts.Select(c => c.md_Title));
                    header.AddRange(moalefeContractsspical.Select(c => c.md_Title));
                    header.AddRange(asnadstr.Select(a => a.Title));

                    writer.WriteLine(string.Join(",", header));

                    // کاربران ماه جاری (از هر دو لیست)
                    var usersInMonth = moalefeValues
                                       .Where(p => p.mlfvlfsh_Month == month &&p.FK_Moalefe==1979)
                                       .Select(p => p.tbUsers)
                                       //.Concat(moalefeValuesSpecial.Where(p => p.mlfvlfsh_Month == month).Select(p => p.tbUsers))
                                       .Distinct()
                                       .ToList();

                    foreach (var user in usersInMonth)
                    {
                        var row = new List<string>
                {
                    $"\"{user.FullName?.Replace("\"","\"\"") ?? ""}\"",
                    user.usr_Personal_ID.ToString(),
                    $"\"{user.tbjob?.Name?.Replace("\"","\"\"") ?? ""}\"",
                    $"\"{db.tbCities.Where(c => c.ID == user.usr_City_Dutysystem).Select(c => c.Name).FirstOrDefault()?.Replace("\"","\"\"") ?? ""}\"",
                    month.ToString(),
                    year.ToString()
                };

                        // moalde و clock
                        var moad = moadData.FirstOrDefault(p => p.UserID == user.usr_ID && p.Month == month && p.Year == year);
                        double moalde = moad?.Moadel ?? 0;

                        var clockData = maxkarkardMonth.FirstOrDefault(p => p.FKUser == user.usr_ID && p.Month == month && p.Year == year);
                        double clock = clockData?.RemainValue ?? 0;

                        row.Add(moalde.ToString());
                        row.Add(clock.ToString());

                        // مقادیر Moalefe
                        foreach (var contract in moalefeContracts)
                        {
                            var val = moalefeValues
                                      .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == contract.md_ID && v.mlfvlfsh_Month == month)
                                      .Select(v => v.mlfvlfsh_Value)
                                      .FirstOrDefault();
                            row.Add(val?.ToString() ?? "0");
                        }

                        foreach (var contract in moalefeContractsspical)
                        {
                            var val = moalefeValuesSpecial
                                      .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == contract.md_ID && v.mlfvlfsh_Month == month)
                                      .Select(v => v.mlfvlfsh_Value)
                                      .FirstOrDefault();
                            row.Add(val?.ToString() ?? "0");
                        }

                        // مقادیر Asnad
                        foreach (var asnad in asnadstr)
                        {
                            long value = 0;
                            var docs = financialDocs.Where(f => f.User_ID == user.usr_ID && f.FK_final == asnad.ID && f.DataDocument.HasValue).ToList();

                            foreach (var doc in docs)
                            {
                                var pc = new PersianCalendar();
                                int docYear = pc.GetYear(doc.DataDocument.Value);
                                int docMonth = pc.GetMonth(doc.DataDocument.Value);
                                if (docYear == year && docMonth == month)
                                {
                                    if (doc.Creditor != 0 && doc.Debtore != 0)
                                        value = (long)(doc.Creditor - doc.Debtore);
                                    else if (doc.Debtore != 0)
                                        value = (long)(doc.Debtore);
                                    else if (doc.Creditor != 0)
                                        value = (long)(doc.Creditor);
                                }
                            }

                            row.Add(value.ToString());
                        }

                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }

            return csvPath;
        }

        public Workbook SetDataExcel_Moalefe175422()
        {
            int Month=0; int Year = 1404;
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelFIS_jonobi.xlsx"));
            for (int i = 1; i <= 9; i++)
            {
                Month = i;
                // Split the input string into a list of strings
                //var moalefeIDList = MoalefeID2.Split(',').ToList();
                List<int> MoalefeID = new List<int>();

                // Convert each string in the list to an integer
                //foreach (var it in moalefeIDList)
                //{
                //    if (int.TryParse(it.Trim(), out int moalefeId))
                //    {
                //        MoalefeID.Add(moalefeId);
                //    }
                //    else
                //    {
                //        // Handle the invalid format case, you can log it or throw an exception
                //        throw new FormatException($"Input string '{it}' was not in a correct format.");
                //    }
                //}
                foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
                {
                    MoalefeID.Add((int)item.Key);

                }

                Row Row;
                var x = db.tbMoadelPadashJarimeAyab.ToList();
                List<tbUsers> tyyy = new List<tbUsers>();
                foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList())
                {
                    tyyy.Add(item.tbUsers);
                }
                //var tyyy = db.tbUsers.ToList();
                var caran = db.tbCaranSettings.ToList();
                var asnad = db.tbfkfinancial.ToList();
                var FinancialDocuments = db.FinancialDocuments.ToList();

                List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
                foreach (var it in asnad)
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                    int persianYear = persianCalendar.GetYear(dataDocument);
                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                    if (persianYear == Year && persianMonth == Month)
                    {
                        asnadstr.Add(it);

                    }
                }
                var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

                var y = db.tbMoalefeValueFish
                            .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month)
                            .ToList();
                int counter = 1;
                int counter2 = 2;






                var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
                var t = db.tbContractMoalefeDastmozdi.ToList();
                var count = 10;
                foreach (var item in MoalefeID)
                {
                    var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[i].AddRow(Row);
                }

                foreach (var item in asnadstr)
                {

                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                    count++;
                    Moalefeexcelfile.Sheets[i].AddRow(Row);
                }






                foreach (var item in tyyy)
                {

                    string job = "";
                    if (item.tbjob != null)
                    {
                        job = item.tbjob.Name;
                    }
                    bool amani = false;
                    if (item.usr_amani == true)
                    {
                        amani = true;
                    }
                    var tb = noalf
    .Where(p => p.MoalfeVal_FKUser == item.usr_ID

                && p.MoalfeVal_Value != 0
                )
    .ToList();
                    List<string> mode = new List<string>();

                    foreach (var it in tb)
                    {
                        var yuuu = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                            .FirstOrDefault();

                        if (yuuu == null)
                        {
                            var u = caran
                                .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                                .FirstOrDefault();

                            if (u != null)
                            {
                                mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{it.MoalfeVal_Value} )----");
                            }
                        }
                        else
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
                        }
                    }
                    double moalde = 0;
                    double clock = 0;
                    var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                    if (moad != null && moad.Moadel != null)
                    {
                        moalde = moad.Moadel;
                    }
                    var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                    if (exi != null && exi.RemainValue != null)
                    {
                        clock = (double)exi.RemainValue;
                    }
                    if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                    {

                        Row = new Row() { Height = 20, Index = counter };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                   ,


                                                   new Cell()
                        {
                            Value = Month,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        }
                                                           ,


                                                   new Cell()
                        {
                            Value = Year,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 9
                        }
                    });


                            //مقادیر مولفه ها
                            int index = 10;
                            foreach (var item2 in MoalefeID)
                            {
                                var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                double ty = 0;
                                if (exx != null)
                                {
                                    ty = (double)exx.mlfvlfsh_Value;
                                }
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            foreach (var item2 in asnadstr)
                            {
                                long value = 0;
                                var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                                foreach (var item1 in t3)
                                {
                                    if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                    {
                                        PersianCalendar persianCalendar = new PersianCalendar();
                                        DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                        int persianYear = persianCalendar.GetYear(dataDocument);
                                        int persianMonth = persianCalendar.GetMonth(dataDocument);
                                        if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                        {
                                            if (item1.Debtore != 0 && item1.Creditor != 0)
                                            {
                                                value = (long)(item1.Creditor - item1.Debtore);
                                            }


                                            else if (item1.Debtore != 0 && item1.Debtore != null)
                                            {
                                                value = (long)(item1.Debtore ?? 0);
                                            }
                                            else if (item1.Creditor != 0 && item1.Creditor != null)
                                            {
                                                value = (long)(item1.Creditor ?? 0);

                                            }

                                        }

                                    }
                                }

                                //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                                //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                                //double ty = 0;
                                //if (exx != null)
                                //{
                                //    ty = (double)exx.mlfvlfsh_Value;
                                //}
                                Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                                index++;
                            }
                            counter++;
                            Moalefeexcelfile.Sheets[i].AddRow(Row);
                        }
                    }
                    //نام ، نام خانوادگی و کد پرسنلی





                }



























            }

            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe_Final()
        {
            int year = 1404;
            var workbook = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelFIS_jonobi.xlsx"));
            var months = Enumerable.Range(1, 9).ToList();

            // Preload data
            var moalefeValues = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == year).ToList();
            var contracts = db.tbContractMoalefeDastmozdi.ToDictionary(c => c.md_ID);
            var caranSettings = db.tbCaranSettings.ToDictionary(c => c.FK_Moalefe_ID);
            var usersDict = db.tbUsers.ToDictionary(u => u.usr_ID);
            var citiesDict = db.tbCities.ToDictionary(c => c.ID);
            var moadelData = db.tbMoadelPadashJarimeAyab
     .Where(m => m.Year == year)
     .ToList();  // <- انتقال به حافظه

            // سپس روی حافظه GroupBy و Dictionary بسازیم
            var moadelDict = moadelData
                .GroupBy(m => Tuple.Create(m.UserID.Value, m.Month.Value))
                .ToDictionary(g => g.Key, g => g.First().Moadel);

            var financialDocs = db.FinancialDocuments.ToList();
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Year == year).ToList();

            foreach (var month in months)
            {
                var sheet = workbook.Sheets[month];
                var monthValues = moalefeValues.Where(v => v.mlfvlfsh_Month == month).ToList();

                // Unique Moalefe IDs
                var moalefeIds = monthValues
                    .Where(v => v.FK_EXCel == null)
                    .GroupBy(v => v.FK_Moalefe)
                    .Select(g => g.Key.Value)
                    .ToList();

                int count = 10;

                // Add Moalefe Titles
                foreach (var id in moalefeIds)
                {
                    if (!contracts.TryGetValue(id, out var contract)) continue;

                    var row = new Row { Height = 20, Index = 0 };
                    row.AddCells(new List<Cell>
            {
                new Cell
                {
                    Value = contract.md_Title,
                    FontFamily = "B Nazanin",
                    FontSize = 12,
                    Index = count
                }
            });
                    count++;
                    sheet.AddRow(row);
                }

                // Add Financial Document Titles
                foreach (var doc in db.tbfkfinancial.Where(f => f.DataDocument.HasValue
                                                                && f.DataDocument.Value.Year == year
                                                                && f.DataDocument.Value.Month == month))
                {
                    var row = new Row { Height = 20, Index = 0 };
                    row.AddCells(new List<Cell>
            {
                new Cell
                {
                    Value = doc.Title,
                    FontFamily = "B Nazanin",
                    FontSize = 12,
                    Index = count
                }
            });
                    count++;
                    sheet.AddRow(row);
                }

                // Users in month
                var usersInMonth = monthValues.Select(v => v.tbUsers).Distinct().ToList();
                int rowIndex = 1;

                foreach (var user in usersInMonth)
                {
                    var row = new Row { Height = 20, Index = rowIndex };

                    double moalde = 0;
                    moadelDict.TryGetValue(Tuple.Create(user.usr_ID, month), out moalde);

                    double clock = db.tbMaxkarkardMonth
                        .Where(c => c.FKUser == user.usr_ID && c.Month == month && c.Year == year)
                        .Select(c => c.RemainValue ?? 0)
                        .FirstOrDefault();

                    // Base info
                    row.AddCells(new List<Cell>
            {
                new Cell { Value = user.FullName, FontFamily = "B Nazanin", FontSize = 12, Index = 0 },
                new Cell { Value = user.usr_Personal_ID, FontFamily = "B Nazanin", FontSize = 12, Index = 1 },
                new Cell { Value = user.tbjob?.Name, FontFamily = "B Nazanin", FontSize = 12, Index = 2 },
                new Cell { Value = citiesDict.TryGetValue(user.usr_City_Dutysystem ?? 0, out var city) ? city.Name : "", FontFamily = "B Nazanin", FontSize = 12, Index = 3 },
                new Cell { Value = month, FontFamily = "B Nazanin", FontSize = 12, Index = 4 },
                new Cell { Value = year, FontFamily = "B Nazanin", FontSize = 12, Index = 5 },
                new Cell { Value = moalde, FontFamily = "B Nazanin", FontSize = 12, Index = 6 },
                new Cell { Value = clock, FontFamily = "B Nazanin", FontSize = 12, Index = 7 }
            });

                    // Moalefe values
                    int cellIndex = 10;
                    foreach (var mid in moalefeIds)
                    {
                        double val = monthValues
                            .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == mid)
                            .Select(v => v.mlfvlfsh_Value)
                            .FirstOrDefault() ?? 0;  // اگر null بود 0 قرار بده


                        row.AddCells(new List<Cell>
                {
                    new Cell { Value = val, FontFamily = "B Nazanin", FontSize = 12, Index = cellIndex }
                });
                        cellIndex++;
                    }

                    sheet.AddRow(row);
                    rowIndex++;
                }
            }

            return workbook;
        }

        public Workbook SetDataExcel_Moalefe_Final_csv()
        {
            int year = 1404;
            var workbook = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelFIS_jonobi.xlsx"));
            var months = Enumerable.Range(1, 9).ToList();

            // Preload data
            var moalefeValues = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == year).ToList();
            var contracts = db.tbContractMoalefeDastmozdi.ToDictionary(c => c.md_ID);
            var caranSettings = db.tbCaranSettings.ToDictionary(c => c.FK_Moalefe_ID);
            var usersDict = db.tbUsers.ToDictionary(u => u.usr_ID);
            var citiesDict = db.tbCities.ToDictionary(c => c.ID);
            var moadelData = db.tbMoadelPadashJarimeAyab
         .Where(m => m.Year == year)
         .ToList();  // <- انتقال به حافظه

            // سپس روی حافظه GroupBy و Dictionary بسازیم
            var moadelDict = moadelData
                .GroupBy(m => Tuple.Create(m.UserID.Value, m.Month.Value))
                .ToDictionary(g => g.Key, g => g.First().Moadel);

            var financialDocs = db.FinancialDocuments.ToList();
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Year == year).ToList();

            foreach (var month in months)
            {
                var sheet = workbook.Sheets[month];
                var monthValues = moalefeValues.Where(v => v.mlfvlfsh_Month == month).ToList();

                // Unique Moalefe IDs
                var moalefeIds = monthValues
                    .Where(v => v.FK_EXCel == null)
                    .GroupBy(v => v.FK_Moalefe)
                    .Select(g => g.Key.Value)
                    .ToList();

                int count = 10;

                // Add Moalefe Titles
                foreach (var id in moalefeIds)
                {
                    if (!contracts.TryGetValue(id, out var contract)) continue;

                    var row = new Row { Height = 20, Index = 0 };
                    row.AddCells(new List<Cell>
            {
                new Cell
                {
                    Value = contract.md_Title,
                    FontFamily = "B Nazanin",
                    FontSize = 12,
                    Index = count
                }
            });
                    count++;
                    sheet.AddRow(row);
                }

                // Add Financial Document Titles
                foreach (var doc in db.tbfkfinancial.Where(f => f.DataDocument.HasValue
                                                                && f.DataDocument.Value.Year == year
                                                                && f.DataDocument.Value.Month == month))
                {
                    var row = new Row { Height = 20, Index = 0 };
                    row.AddCells(new List<Cell>
            {
                new Cell
                {
                    Value = doc.Title,
                    FontFamily = "B Nazanin",
                    FontSize = 12,
                    Index = count
                }
            });
                    count++;
                    sheet.AddRow(row);
                }

                // Users in month
                var usersInMonth = monthValues.Select(v => v.tbUsers).Distinct().ToList();
                int rowIndex = 1;

                foreach (var user in usersInMonth)
                {
                    var row = new Row { Height = 20, Index = rowIndex };

                    double moalde = 0;
                    moadelDict.TryGetValue(Tuple.Create(user.usr_ID, month), out moalde);

                    double clock = db.tbMaxkarkardMonth
                        .Where(c => c.FKUser == user.usr_ID && c.Month == month && c.Year == year)
                        .Select(c => c.RemainValue ?? 0)
                        .FirstOrDefault();

                    // Base info
                    row.AddCells(new List<Cell>
            {
                new Cell { Value = user.FullName, FontFamily = "B Nazanin", FontSize = 12, Index = 0 },
                new Cell { Value = user.usr_Personal_ID, FontFamily = "B Nazanin", FontSize = 12, Index = 1 },
                new Cell { Value = user.tbjob?.Name, FontFamily = "B Nazanin", FontSize = 12, Index = 2 },
                new Cell { Value = citiesDict.TryGetValue(user.usr_City_Dutysystem ?? 0, out var city) ? city.Name : "", FontFamily = "B Nazanin", FontSize = 12, Index = 3 },
                new Cell { Value = month, FontFamily = "B Nazanin", FontSize = 12, Index = 4 },
                new Cell { Value = year, FontFamily = "B Nazanin", FontSize = 12, Index = 5 },
                new Cell { Value = moalde, FontFamily = "B Nazanin", FontSize = 12, Index = 6 },
                new Cell { Value = clock, FontFamily = "B Nazanin", FontSize = 12, Index = 7 }
            });

                    // Moalefe values
                    int cellIndex = 10;
                    foreach (var mid in moalefeIds)
                    {
                        double val = monthValues
        .Where(v => v.FK_User == user.usr_ID && v.FK_Moalefe == mid)
        .Select(v => v.mlfvlfsh_Value)
        .FirstOrDefault() ?? 0;  // اگر null بود 0 قرار بده

                        row.AddCells(new List<Cell>
                {
                    new Cell { Value = val, FontFamily = "B Nazanin", FontSize = 12, Index = cellIndex }
                });
                        cellIndex++;
                    }

                    sheet.AddRow(row);
                    rowIndex++;
                }
            }

            return workbook;
        }

        private Cell CellText(object value, int index)
        {
            return new Cell
            {
                Value = value,
                FontFamily = "B Nazanin",
                FontSize = 12,
                Index = index
            };
        }

        private Cell CellNumber(double value, int index)
        {
            return new Cell
            {
                Value = value,
                FontFamily = "B Nazanin",
                FontSize = 12,
                Index = index
            };
        }

        //    public ActionResult ExportMoalefeCsv()
        //    {
        //        Response.Clear();
        //        Response.Buffer = false;
        //        Response.AddHeader("Content-Disposition", "attachment;filename=Moalefe_1404.csv");
        //        Response.ContentType = "text/csv";

        //        string connStr = GetSqlConnectionString();

        //        var sql = @"
        //WITH RankedData AS (
        //    SELECT 
        //        u.usr_ID,
        //        u.usr_Name + ' ' + u.usr_Family AS FullName,
        //        u.usr_Personal_ID,
        //        c.Name AS CityName,
        //        v.FK_Moalefe,
        //        v.mlfvlfsh_Value,
        //        v.mlfvlfsh_Month,
        //        ROW_NUMBER() OVER(PARTITION BY u.usr_ID, v.FK_Moalefe, v.mlfvlfsh_Month
        //                          ORDER BY v.mlfvlfsh_ID DESC) AS rn
        //    FROM Salary.tbMoalefeValueFish v
        //    INNER JOIN Usser.tbUsers u ON v.FK_User = u.usr_ID
        //    LEFT JOIN dbo.tbCities c ON u.usr_City_Dutysystem = c.ID
        //    WHERE v.mlfvlfsh_Year = 1404 AND v.mlfvlfsh_Month BETWEEN 1 AND 8
        //)
        //SELECT *
        //FROM RankedData
        //WHERE rn = 1
        //ORDER BY usr_ID, FK_Moalefe, mlfvlfsh_Month;
        //";

        //        var matrix = new Dictionary<int, List<double>>();
        //        var users = new Dictionary<int, (string FullName, string PersID, string City)>();

        //        using (var conn = new SqlConnection(connStr))
        //        using (var cmd = new SqlCommand(sql, conn))
        //        {
        //            conn.Open();
        //            using (var reader = cmd.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    int userId = Convert.ToInt32(reader["usr_ID"]);
        //                    int month = Convert.ToInt32(reader["mlfvlfsh_Month"]);
        //                    double v = Convert.ToDouble(reader["mlfvlfsh_Value"]);

        //                    if (!matrix.ContainsKey(userId))
        //                        matrix[userId] = Enumerable.Repeat(0.0, 8).ToList();

        //                    matrix[userId][month - 1] += v;

        //                    if (!users.ContainsKey(userId))
        //                    {
        //                        users[userId] = (
        //                            reader["FullName"].ToString(),
        //                            reader["usr_Personal_ID"]?.ToString(),
        //                            reader["CityName"]?.ToString()
        //                        );
        //                    }
        //                }
        //            }
        //        }

        //        using (var writer = new StreamWriter(Response.OutputStream, Encoding.UTF8))
        //        {
        //            writer.WriteLine("FullName,PersonalID,City,Month1,Month2,Month3,Month4,Month5,Month6,Month7,Month8,Total");

        //            foreach (var u in users)
        //            {
        //                var d = matrix[u.Key];
        //                writer.Write($"{u.Value.FullName},{u.Value.PersID},{u.Value.City}");
        //                foreach (var item in d)
        //                    writer.Write($",{item}");
        //                writer.Write($",{d.Sum()}");
        //                writer.WriteLine();
        //            }
        //        }

        //        Response.End();
        //        return new EmptyResult();
        //    }

        public Workbook SetDataExcel_Moalefe17542versionasli(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            //foreach (var it in moalefeIDList)
            //{
            //    if (int.TryParse(it.Trim(), out int moalefeId))
            //    {
            //        MoalefeID.Add(moalefeId);
            //    }
            //    else
            //    {
            //        // Handle the invalid format case, you can log it or throw an exception
            //        throw new FormatException($"Input string '{it}' was not in a correct format.");
            //    }
            //}
            foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
            {
                MoalefeID.Add((int)item.Key);

            }
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year&&p.FK_EXCel==null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month&&p.FK_EXCel==null)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 9;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                                              ,


                           new Cell()
{
    Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
    FontFamily = "B Nazanin",
    Bold = false,
    Enable = true,
    Wrap = false,
    FontSize = 12,
    Italic = false,
    Underline = false,
    Index = 8
}
                    });


                        //مقادیر مولفه ها
                        int index = 9;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = -2;
                            if (exx != null &&exx.mlfvlfsh_Value!=null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe17542versionasli_modd(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            //foreach (var it in moalefeIDList)
            //{
            //    if (int.TryParse(it.Trim(), out int moalefeId))
            //    {
            //        MoalefeID.Add(moalefeId);
            //    }
            //    else
            //    {
            //        // Handle the invalid format case, you can log it or throw an exception
            //        throw new FormatException($"Input string '{it}' was not in a correct format.");
            //    }
            //}
            //foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
            //{
            //    MoalefeID.Add((int)item.Key);

            //}
            int[] moalefeList = { 1961, 1960, 1964, 1965, 1975, 1979 };

            foreach (var item in db.tbMoalefeValueFish
              .Where(s =>
    s.mlfvlfsh_Month == Month &&
    s.mlfvlfsh_Year == Year &&
    s.FK_EXCel == null &&
    (s.FK_Moalefe == 1961 ||
     s.FK_Moalefe == 1960 ||
     s.FK_Moalefe == 1964 ||
     s.FK_Moalefe == 1965 ||
     s.FK_Moalefe == 1975 ||
     s.FK_Moalefe == 1979))

                .GroupBy(s => s.FK_Moalefe)
                .ToList())
            {
                MoalefeID.Add((int)item.Key);
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish1.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();
            List<tbfkfinancial> asnadstr1 = new List<tbfkfinancial>();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month-1)
                {
                    asnadstr1.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel == null)
                        .ToList();
            var y1 = db.tbMoalefeValueFish
                .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month-1 && p.FK_EXCel == null)
                .ToList();
            int counter = 1;
            int counter2 = 2;
            
            var find20 = db.tbkarkard_notsystem.Where(s => s.moadel_kol != null && s.month == Month && s.year == Year).ToList();

            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 11;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            count = count + 2;
            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList(); 
                var fidnn1 =  find20
.Where(p => p.FK_usr == item.usr_ID

            && p.moadel != 0
            )
.OrderByDescending(s=>s.ID)
.FirstOrDefault();
                var tb1 = find20
.Where(p => p.FK_usr == item.usr_ID

            && p.moadel != 0
            )
.ToList();
                if (fidnn1 != null)
                {
                    tb1 = find20
.Where(p => p.FK_usr == item.usr_ID
&& p.usrsabt == fidnn1.usrsabt
           && p.moadel != 0
           )
.ToList();
                }
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                List<string> mode2 = new List<string>();

                foreach (var it in tb1)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.FK_moaldeh && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.moadel ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.FK_moaldeh && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode2.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode2.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                var exx1 = y.Where(p => p.FK_Moalefe == 1979 && p.FK_User == item.usr_ID).FirstOrDefault();
                double ty1 = 0;
                if (exx1 != null && exx1.mlfvlfsh_Value != null)
                {
                    ty1 = (double)exx1.mlfvlfsh_Value;
                }
                var exx11 = y1.Where(p => p.FK_Moalefe == 1979 && p.FK_User == item.usr_ID).FirstOrDefault();
                double ty11 = 0;
                if (exx11 != null && exx11.mlfvlfsh_Value != null)
                {
                    ty11 = (double)exx11.mlfvlfsh_Value;
                }
                var exx2 = y.Where(p => p.FK_Moalefe == 1975 && p.FK_User == item.usr_ID).FirstOrDefault();
                double ty2 = 0;
                if (exx2 != null && exx2.mlfvlfsh_Value != null)
                {
                    ty2 = (double)exx2.mlfvlfsh_Value;
                }
                foreach (var item2 in asnadstr)
                {
                    long value = 0;
                    var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                    foreach (var item1 in t3)
                    {
                        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                        {
                            PersianCalendar persianCalendar = new PersianCalendar();
                            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                            int persianYear = persianCalendar.GetYear(dataDocument);
                            int persianMonth = persianCalendar.GetMonth(dataDocument);
                            if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                            {
                                if (item1.Debtore != 0 && item1.Creditor != 0)
                                {
                                    ty1+= (long)(item1.Creditor - item1.Debtore);
                                }


                                else if (item1.Debtore != 0 && item1.Debtore != null)
                                {
                                    ty1-= (long)(item1.Debtore ?? 0);
                                }
                                else if (item1.Creditor != 0 && item1.Creditor != null)
                                {
                                    ty1+= (long)(item1.Creditor ?? 0);

                                }

                            }

                        }
                    }

                    //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                    //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                    //double ty = 0;
                    //if (exx != null)
                    //{
                    //    ty = (double)exx.mlfvlfsh_Value;
                    //}
          
                }
                foreach (var item2 in asnadstr1)
                {
                    long value = 0;
                    var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                    foreach (var item1 in t3)
                    {
                        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                        {
                            PersianCalendar persianCalendar = new PersianCalendar();
                            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                            int persianYear = persianCalendar.GetYear(dataDocument);
                            int persianMonth = persianCalendar.GetMonth(dataDocument);
                            if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                            {
                                if (item1.Debtore != 0 && item1.Creditor != 0)
                                {
                                    ty11 += (long)(item1.Creditor - item1.Debtore);
                                }


                                else if (item1.Debtore != 0 && item1.Debtore != null)
                                {
                                    ty11 -= (long)(item1.Debtore ?? 0);
                                }
                                else if (item1.Creditor != 0 && item1.Creditor != null)
                                {
                                    ty11 += (long)(item1.Creditor ?? 0);

                                }

                            }

                        }
                    }

                    //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                    //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                    //double ty = 0;
                    //if (exx != null)
                    //{
                    //    ty = (double)exx.mlfvlfsh_Value;
                    //}

                }

                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }

                double moalde_notsystem = 0;
                double moalde_notsystem_kol = 0;

                var moad122 = find20.Where(p => p.FK_usr == item.usr_ID ).OrderByDescending(s=>s.ID).FirstOrDefault();
                if (moad122 != null && moad122.moadel_kol != null)
                {
                    moalde_notsystem = (double)moad122.moadel_kol;
                }
                if (moalde_notsystem != 0)
                {
                    moalde_notsystem_kol = (double)( moalde_notsystem - moalde);
                }
                else
                {
                    moalde_notsystem = (double)moalde;
                }
                double moalde1 = 0;
                double clock1 = 0;
                var moad1 = x.Where(p => p.UserID == item.usr_ID && p.Month == Month-1 && p.Year == Year).FirstOrDefault();
                if (moad1 != null && moad1.Moadel != null)
                {
                    moalde1 = moad1.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },          new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode2),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 8
                        },         new Cell()
                        {
                            Value = moalde_notsystem_kol,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 9
                        },
                                        new Cell()
                        {
                            Value = moalde_notsystem,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 10
                        },
                                   

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        }
                                                                              ,


                           new Cell()
{
    Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
    FontFamily = "B Nazanin",
    Bold = false,
    Enable = true,
    Wrap = false,
    FontSize = 12,
    Italic = false,
    Underline = false,
    Index = 5
}                 ,



                          

                   });


                        //مقادیر مولفه ها
                        int index = 11;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = -2;
                            if (exx != null && exx.mlfvlfsh_Value != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        index++;

                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty1,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        index++;



                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }

                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = moalde1,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        index++;

                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty11,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe17542versionasli_saghf(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();
            List<long> MoalefeID3 = new List<long>();

            // Convert each string in the list to an integer
            //foreach (var it in moalefeIDList)
            //{
            //    if (int.TryParse(it.Trim(), out int moalefeId))
            //    {
            //        MoalefeID.Add(moalefeId);
            //    }
            //    else
            //    {
            //        // Handle the invalid format case, you can log it or throw an exception
            //        throw new FormatException($"Input string '{it}' was not in a correct format.");
            //    }
            //}
            foreach (var item in db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_EXCel == null).GroupBy(s => s.FK_Moalefe).ToList())
            {
                MoalefeID.Add((int)item.Key);

            }
            foreach (var item in db.dbtarifmahdodayt1.GroupBy(s => s.FK_moalfe).ToList())
            {
                MoalefeID3.Add((int)item.Key);

            }
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel == null)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}



            var dbtarifmahdodayt2 = db.dbtarifmahdodayt2.ToList();
            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 9;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }



            foreach (var item in MoalefeID3)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }


            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                                              ,


                           new Cell()
{
    Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
    FontFamily = "B Nazanin",
    Bold = false,
    Enable = true,
    Wrap = false,
    FontSize = 12,
    Italic = false,
    Underline = false,
    Index = 8
}
                    });


                        //مقادیر مولفه ها
                        int index = 9;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = -2;
                            if (exx != null && exx.mlfvlfsh_Value != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in MoalefeID3)
                        {
                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = dbtarifmahdodayt2.Where(p => p.dbtarifmahdodayt1.FK_moalfe == item2 && p.FK_usr == item.usr_ID).FirstOrDefault();
                            double ty = -3;
                            if (exx != null && exx.value != null)
                            {
                                ty = (double)exx.value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe17542versionaslitest(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            foreach (var it in moalefeIDList)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_ID == 1979 && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel == null)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 9;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();
                    var roundedValue = Math.Round(it.MoalfeVal_Value ?? 0);
                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{roundedValue} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{roundedValue} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = db.tbCities.Where(p=>p.ID==item.usr_City_Dutysystem).Select(s=>s.Name).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                                                                              ,


                           new Cell()
{
    Value = db.tbCities.Where(s=>s.ID==item.usr_City_Dutysystem).Select(s=>s.ZaribSharestan).FirstOrDefault(),
    FontFamily = "B Nazanin",
    Bold = false,
    Enable = true,
    Wrap = false,
    FontSize = 12,
    Italic = false,
    Underline = false,
    Index = 8
}
                    });


                        //مقادیر مولفه ها
                        int index = 9;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = -2;
                            if (exx != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }


        public Workbook ExportMoalefeExcelcontrolmahdodatdowncorrect(string MoalefeID2, int Month, int Year)
        {
            // Split the input string into a list of strings
            var moalefeIDList = MoalefeID2.Split(',').ToList();
            List<int> MoalefeID = new List<int>();

            // Convert each string in the list to an integer
            foreach (var it in moalefeIDList)
            {
                if (int.TryParse(it.Trim(), out int moalefeId))
                {
                    MoalefeID.Add(moalefeId);
                }
                else
                {
                    // Handle the invalid format case, you can log it or throw an exception
                    throw new FormatException($"Input string '{it}' was not in a correct format.");
                }
            }

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            List<tbUsers> tyyy = new List<tbUsers>();
            foreach (var item in db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار" && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel != null&&p.Finalaccept==true).ToList())
            {
                tyyy.Add(item.tbUsers);
            }
            //var tyyy = db.tbUsers.ToList();
            var caran = db.tbCaranSettings.ToList();
            var asnad = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();

            List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
            foreach (var it in asnad)
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                int persianYear = persianCalendar.GetYear(dataDocument);
                int persianMonth = persianCalendar.GetMonth(dataDocument);
                if (persianYear == Year && persianMonth == Month)
                {
                    asnadstr.Add(it);

                }
            }
            var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();

            var y = db.tbMoalefeValueFish
                        .Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_EXCel == null)
                        .OrderByDescending(p => p.mlfvlfsh_ID)
                        .ToList();
            int counter = 1;
            int counter2 = 2;


            //foreach (var itemmm in yyyyyyy3)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == itemmm.FK_User_ID).ToList();
            //    foreach (var item1 in t)
            //    {
            //        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
            //        {
            //            PersianCalendar persianCalendar = new PersianCalendar();
            //            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
            //            int persianYear = persianCalendar.GetYear(dataDocument);
            //            int persianMonth = persianCalendar.GetMonth(dataDocument);

            //            if (persianYear == year && persianMonth == month)
            //            {
            //                if (item1.Debtore != 0 && item1.Debtore != null)
            //                {
            //                    traz -= (long)item1.Debtore;
            //                    traz3 -= (long)item1.Debtore;

            //                }
            //                else if (item1.Creditor != 0 && item1.Creditor != null)
            //                {
            //                    traz += (long)item1.Creditor;
            //                    traz2 += (long)item1.Creditor;


            //                }
            //            }

            //        }
            //    }

            //}




            var clov = db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.Year == Year).ToList();
            var t = db.tbContractMoalefeDastmozdi.ToList();
            var count = 8;
            foreach (var item in MoalefeID)
            {
                var ex = t.Where(p => p.md_ID == item).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = ex.md_Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            foreach (var item in asnadstr)
            {

                Row = new Row() { Height = 20, Index = 0 };
                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Title,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = count
            },
        });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }






            foreach (var item in tyyy)
            {

                string job = "";
                if (item.tbjob != null)
                {
                    job = item.tbjob.Name;
                }
                bool amani = false;
                if (item.usr_amani == true)
                {
                    amani = true;
                }
                var tb = noalf
.Where(p => p.MoalfeVal_FKUser == item.usr_ID

            && p.MoalfeVal_Value != 0
            )
.ToList();
                List<string> mode = new List<string>();

                foreach (var it in tb)
                {
                    var yuuu = caran
                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
                        .FirstOrDefault();

                    if (yuuu == null)
                    {
                        var u = caran
                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
                            .FirstOrDefault();

                        if (u != null)
                        {
                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {u.CaranAyabOZahab}/ 0 /{it.MoalfeVal_Value} )----");
                        }
                    }
                    else
                    {
                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {yuuu.CaranAyabOZahab}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
                    }
                }
                double moalde = 0;
                double clock = 0;
                var moad = x.Where(p => p.UserID == item.usr_ID && p.Month == Month && p.Year == Year).FirstOrDefault();
                if (moad != null && moad.Moadel != null)
                {
                    moalde = moad.Moadel;
                }
                var exi = clov.Where(p => p.FKUser == item.usr_ID).FirstOrDefault();
                if (exi != null && exi.RemainValue != null)
                {
                    clock = (double)exi.RemainValue;
                }
                if (y.Where(p => p.FK_User == item.usr_ID).FirstOrDefault() != null)


                {

                    Row = new Row() { Height = 20, Index = counter };
                    {
                        Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },

                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                             new Cell()
                        {
  Value = string.Join(Environment.NewLine, mode),
                                 FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                               new Cell()
                        {
                            Value = moalde,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                    new Cell()
                        {
                            Value = clock,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },

                                                   new Cell()
                        {
                            Value = db.Link_User_And_Peyman.Where(p=>p.FK_User_ID==item.usr_ID&&p.Status==true).Select(s=>s.tbPeymanContracts.pec_Title).FirstOrDefault(),
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },


                                                   new Cell()
                        {
                            Value = job,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },


                                                   new Cell()
                        {
                            Value = amani,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }
                    });


                        //مقادیر مولفه ها
                        int index = 8;
                        foreach (var item2 in MoalefeID)
                        {
                            var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            double ty = 0;
                            if (exx != null)
                            {
                                ty = (double)exx.mlfvlfsh_Value;
                            }
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = true,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        foreach (var item2 in asnadstr)
                        {
                            long value = 0;
                            var t3 = FinancialDocuments.Where(p => p.User_ID == item.usr_ID && p.FK_final == item2.ID).ToList();

                            foreach (var item1 in t3)
                            {
                                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                                {
                                    PersianCalendar persianCalendar = new PersianCalendar();
                                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                                    int persianYear = persianCalendar.GetYear(dataDocument);
                                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                                    if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
                                    {
                                        if (item1.Debtore != 0 && item1.Creditor != 0)
                                        {
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }

                            //var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();

                            //var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.usr_ID).FirstOrDefault();
                            //double ty = 0;
                            //if (exx != null)
                            //{
                            //    ty = (double)exx.mlfvlfsh_Value;
                            //}
                            Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = value,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                            index++;
                        }
                        counter++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);
                    }
                }
                //نام ، نام خانوادگی و کد پرسنلی





            }



























            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_Moalefe175(List<int> MoalefeID, int Month, int Year)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/seefish.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.ToList();
            var y = db.tbMoalefeValueFish.Where(p =>  p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.tbUsers.usr_amani != true).OrderByDescending(p => p.mlfvlfsh_Value).ToList();
            int counter = 1;
            int counter2 = 2;
            var t = db.tbContractMoalefeDastmozdi.ToList();
     

            foreach (var item in y)
            {
                foreach (var item2 in MoalefeID)
                {

                    var ex = t.Where(p => p.md_ID == item2).FirstOrDefault();
                    if(ex != null)
                    {
                        Row = new Row() { Height = 20, Index = 0 };
                        {
                            Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {

                            Value = ex.md_Title,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = counter2
                        },
                         


                    });

                            counter2++;
                            Moalefeexcelfile.Sheets[0].AddRow(Row);
                        }
                    }
                    var exx = y.Where(p => p.FK_Moalefe == item2 && p.FK_User == item.FK_User).FirstOrDefault();
                    double ty = 0;
                    if (exx != null)
                    {
                        ty = (double)exx.mlfvlfsh_Value;
                    }
                    var xx = db.tbUsers.Where(p => p.usr_ID == item.FK_User).FirstOrDefault();
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {

                            Value = xx.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                        {
                            Value = xx.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                new Cell()
                        {
                            Value = ty,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },


                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            }
            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe16()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/ayabzohab.xlsx"));
            Row Row;
            var x = db.tbMoadelPadashJarimeAyab.Where(p=>p.Month==3).ToList();

            int counter = 1;
            foreach (var item in x)
            {
                var xx = db.tbUsers.Where(p => p.usr_ID == item.UserID).FirstOrDefault();
                string v = "";
                var xx2 = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == item.UserID &&p.Status==true).FirstOrDefault();
                if (xx2 != null)
                {
                    v = xx2.tbPeymanContracts.pec_Title;
                }
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {

                            Value = xx.FullName,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                            new Cell()
                        {
                            Value = xx.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                                new Cell()
                        {
                            Value = item.Padash,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                                      new Cell()
                        {
                            Value = (double)item.Moadel,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                                            new Cell()
                        {
                            Value = item.AyabOZahab,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                                                  new Cell()
                        {
                            Value = item.Jarime,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                                                                           new Cell()
                        {
                            Value = item.UserID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        },                                                 new Cell()
                        {
                            Value = v,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 7
                        }

                    });

                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }












        public string linkcity(List<citymoalfelink> Filters)
        {

            foreach (var filter in Filters)
            {
                //List<MoalefeUserInfo> UsersMoalefe = filter.innerObjects;
                List<tblink_moalfe_city_valu> lstMoalefeexcel = new List<tblink_moalfe_city_valu>();


                foreach (var userMoalefe in filter.innerObjects)
                {


                    {
                        tblink_moalfe_city_valu dastmozdexcel = new tblink_moalfe_city_valu();


                        dastmozdexcel.number = userMoalefe.number;
                        dastmozdexcel.value = userMoalefe.value;

                        lstMoalefeexcel.Add(dastmozdexcel);
                    }



                }
                //if (link_Maolf_city_valu.Create(lstMoalefeexcel) == "True")
                //{



                //}
                tblink_moalfe_city savedfunctions = new tblink_moalfe_city();

                savedfunctions.FK_moalfe = filter.FK_moalfe;
                savedfunctions.Fk_pymn = filter.Fk_pymn;
                savedfunctions.Fk_City = filter.Fk_City;
                savedfunctions.tblink_moalfe_city_valu = lstMoalefeexcel;
                if (link_Maolf_city.Create(savedfunctions) == "True")
                {
                }





            }
            return "True";
        }

        // و در اینجا تابع go را قرار دهید

        [HttpPost]

        public async Task<ActionResult> send(List<FunctionModel> Filters)
        {

            //var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Moalefe.xlsx"));
            Row Row;
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
            List<int> code = new List<int>();
            List<string> malfe = new List<string>();

            foreach (var filter in Filters)
            {

                List<MoalefeUserInfo> UsersMoalefe = filter.UsersMoalefe;
                Header SabtHeader = filter.SabtHeader;
                int MoalfeVal_Year = SabtHeader.MoalfeVal_Year;
                int MoalfeVal_Month = SabtHeader.MoalfeVal_Month;
                string BastehName = SabtHeader.BastehName;
                string PeymanName = SabtHeader.PeymanName;


                // حالا شما می‌توانید به موارد داخل MoalefeUserInfo دسترسی پیدا کنید
                foreach (var userMoalefe in UsersMoalefe)
                {
                    List<MoalefeInfo> listMoalefe = userMoalefe.listMoalefe;
                    string fullName = userMoalefe.FullName;

                    int PersonalCode = userMoalefe.PersonalCode;
                    code.Add(PersonalCode);
                    var perID = System.Convert.ToInt32(PersonalCode);

                    var us = await db.tbUsers.Where(p => p.usr_Personal_ID == PersonalCode).FirstOrDefaultAsync();
                    if (us != null)
                    {

                        var UserID = us.usr_ID;

                        // انجام عملیات مورد نظر بر روی fullName

                        //foreach (var ismalfe in listMoalefe)
                        //{

                        //    string MoalefeTitle = ismalfe.MoalefeTitle.TrimEnd('\n'); ;
                        //    malfe.Add(MoalefeTitle);
                        //    string MoalefeValue = ismalfe.MoalefeValue.TrimEnd('\n'); ;
                        //    var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(MoalefeTitle)).FirstOrDefault().md_ID;
                        //    var v = System.Convert.ToInt32(MoalfeVal_Month);
                        //    var t = System.Convert.ToInt32(MoalfeVal_Year);
                        //    var d = System.Convert.ToDouble(MoalefeValue);
                        //    var x = lstMoalefeexcel.Where(p => p.MoalfeVal_Year == t && p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == v && p.MoalfeVal_Value == d && p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID).FirstOrDefault();
                        //    if (x == null)
                        //    {
                        //        tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                        //        {
                        //            MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                        //            MoalfeVal_FKUser = UserID,
                        //            MoalfeVal_Month = System.Convert.ToInt32(MoalfeVal_Month),
                        //            MoalfeVal_Year = System.Convert.ToInt32(MoalfeVal_Year),
                        //            MoalfeVal_Value = System.Convert.ToDouble(MoalefeValue),


                        //        };
                        //        lstMoalefeexcel.Add(dastmozdexcel);

                        //    }





                        //}


                    }
                    break;
                }
                break;
            }

            string BastehName2 = "";
            string PeymanName2 = "";
            DateTime ToDate2 = DateTime.Now;
            DateTime FromDate2 = DateTime.Now;
            foreach (var filter in Filters)
            {

                List<MoalefeUserInfo> UsersMoalefe = filter.UsersMoalefe;
                Header SabtHeader = filter.SabtHeader;
                int MoalfeVal_Year = SabtHeader.MoalfeVal_Year;
                int MoalfeVal_Month = SabtHeader.MoalfeVal_Month;
                BastehName2 = SabtHeader.BastehName;
                PeymanName2 = SabtHeader.PeymanName;
                ToDate2 = SabtHeader.ToDate;
                FromDate2 = SabtHeader.FromDate;
                break;
            }
            var pymm = await db.tbPeymanContracts.Where(p => p.pec_Title == PeymanName2 && p.Inactive != true).FirstOrDefaultAsync();
            var pymnid = pymm.pec_ID;
            var Bas = await db.tbReffrenceSaveLevel.Where(p => p.Title == BastehName2 && p.Deleted != true && p.tbReffrenceSave.FK_PeymanID == pymnid).FirstOrDefaultAsync();
            var Bastena = Bas.ID;



            var us4 = db.tbReffrenceSave
.Where(p => p.FK_PeymanID == pymnid)
.SelectMany(s => s.tbReffrenceSaveLevel)
.Where(s => s.ID == Bastena)
.OrderByDescending(p => p.ID)
.FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
                .Where(p => p.FK_RRSave == us4.FK_RRSave)
                .OrderByDescending(p => p.ID)
                .FirstOrDefault();
            var rus3 = db.tbReffrenceSaveLevelUser
                .Where(p => p.FK_LevelID == us2.ID)
                .OrderByDescending(p => p.ID)
                .Select(p => p.FK_UserID)
                .FirstOrDefault();
            var User2 = new tbUsers();
            var cookie_user2 = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                //if (User2 != null)
                //{

                //    userid = User2.usr_ID;
                //    if (userid == rus3)
                //    {

                //        var matchedRows = db.tbMoalefeValuePishkhan
                //            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان")
                //            .ToList();

                //        foreach (var row1 in matchedRows)
                //        {
                //            row1.mlfval_Value = "1";

                //        }

                //        db.SaveChanges();

                //    }
                //}

            }











            //var us2 = db.tbUsers.Where(p => p.usr_Personal_ID == PersonalCode).FirstOrDefault();
            tbSavedFunctions savedfunctions = new tbSavedFunctions();
            int userid2 = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid2 = User.usr_ID;
                    }
                }
            }

            if (cookie_user2 != null)
            {
                //int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User2 = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User2 != null)
                {

                    int userid4 = User2.usr_ID;
                    if (userid4 == rus3)
                    {
                        savedfunctions.Final_Sabt = true;



                    }
                }
            }


            else
            {

                return Content("در ذخیره سازی مشکلی به وجود آمده است");
            }
            savedfunctions.svdfunc_BastehID = Bastena;
            savedfunctions.svdfunc_FromDate = FromDate2;
            savedfunctions.svdfunc_ToDate = ToDate2;
            savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
            savedfunctions.svdfunc_pymnID = pymnid;
            savedfunctions.svdfunc_UserSaveID = userid2;
            savedfunctions.svdfunc_IsSubmmit = false;
            //var v = System.Convert.ToInt32(lstMoalefeexcel.MoalfeVal_Year);
            //var t = System.Convert.ToInt32(MoalfeVal_Year);
            //var d = System.Convert.ToDouble(MoalefeValue);
            //var x = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Year == t && p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == v && p.MoalfeVal_Value == d && p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID).FirstOrDefault();


            if (savedfunctionRepo.Create2(savedfunctions) != 0)
            {
                //Task.Run(() =>
                //{
                //    ExportMoalefeExcel3(db.tbSavedFunctions.OrderByDescending(p => p.svdfunc_ID).Select(s=>s.svdfunc_ID).FirstOrDefault());
                //});

            }

            foreach (var filter in Filters)
            {

                List<MoalefeUserInfo> UsersMoalefe = filter.UsersMoalefe;
                Header SabtHeader = filter.SabtHeader;
                int MoalfeVal_Year = SabtHeader.MoalfeVal_Year;
                int MoalfeVal_Month = SabtHeader.MoalfeVal_Month;
                string BastehName = SabtHeader.BastehName;
                string PeymanName = SabtHeader.PeymanName;

                List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
                lstmoalafe2 = db.tbContractMoalefeDastmozdi.ToList();
                // حالا شما می‌توانید به موارد داخل MoalefeUserInfo دسترسی پیدا کنید
                foreach (var userMoalefe in UsersMoalefe)
                {
                    List<MoalefeInfo> listMoalefe = userMoalefe.listMoalefe;
                    string fullName = userMoalefe.FullName;

                    int PersonalCode = userMoalefe.PersonalCode;
                    code.Add(PersonalCode);
                    var perID = System.Convert.ToInt32(PersonalCode);

                    var us =await db.tbUsers.Where(p => p.usr_Personal_ID == PersonalCode).FirstOrDefaultAsync();
                    if (us != null)
                    {

                        var UserID = us.usr_ID;

                        // انجام عملیات مورد نظر بر روی fullName

                        foreach (var ismalfe in listMoalefe)
                        {

                            string MoalefeTitle = ismalfe.MoalefeTitle.TrimEnd('\n'); ;

                            string MoalefeValue = ismalfe.MoalefeValue.TrimEnd('\n'); ;
                            var MoalefeID = lstmoalafe2.Where(p => p.md_Title.Contains(MoalefeTitle)).Select(s=>s.md_ID).FirstOrDefault();
                            var v = System.Convert.ToInt32(MoalfeVal_Month);
                            var t = System.Convert.ToInt32(MoalfeVal_Year);
                            var d = System.Convert.ToDouble(MoalefeValue);
                            var x = lstMoalefeexcel.Where(p => p.MoalfeVal_Year == t && p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == v && p.MoalfeVal_Value == d && p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID).FirstOrDefault();

                            if (x == null)
                            {

                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                {
                                    MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                                    MoalfeVal_FKUser = UserID,
                                    MoalfeVal_Month = System.Convert.ToInt32(MoalfeVal_Month),
                                    MoalfeVal_Year = System.Convert.ToInt32(MoalfeVal_Year),
                                    MoalfeVal_Value = System.Convert.ToDouble(MoalefeValue),
                                    FK_SavedFunctionsID = savedfunctions.svdfunc_ID,


                                };
                                lstMoalefeexcel.Add(dastmozdexcel);

                            }





                        }


                       
                    }
                }
            }
            if (moalefeexcelRepo.Create(lstMoalefeexcel) == "True")
            {
            }
            else
            {
                if (await moalefeexcelRepo.Create2Async(lstMoalefeexcel) > 0)
                {
                    // Saving successful, number of saved items is returned by Create2Async
                    // You can use the returned value here, for example:
                    //int savedCount = await moalefeexcelRepo.Create2Async(lstMoalefeexcel);
                    //Console.WriteLine($"Successfully saved {savedCount} items.");
                }
                else
                {
                    // Saving might have failed (0 items saved) or encountered an exception
                    Console.WriteLine("Saving failed.");
                }

            }


            //var exist = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == pymnid && p.svdfunc_BastehID == Bastena && p.svdfunc_FromDate == FromDate2 && p.svdfunc_ToDate == ToDate2).FirstOrDefault();
            //if (exist != null)//update
            //{
            //    tbSavedFunctions savedfunctions2 = new tbSavedFunctions();
            //    if (cookie_user2 != null)
            //    {
            //        //int userid = 0;
            //        var nationalcode = Utility.Base64.Base64Decode(cookie_user2.Value);
            //        using (SaabEntities db = new SaabEntities())
            //        {
            //            User2 =await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
            //        }
            //        if (User2 != null)
            //        {

            //            int userid4 = User2.usr_ID;
            //            if (userid4 == rus3)
            //            {
            //                savedfunctions2.Final_Sabt = true;



            //            }
            //        }
            //    }

            //    savedfunctions2.svdfunc_BastehID = Bastena;

            //    savedfunctions2.svdfunc_FromDate = FromDate2;
            //    savedfunctions2.svdfunc_ToDate = ToDate2;
            //    savedfunctions2.svdfunc_SavedDateTime = DateTime.Now;
            //    savedfunctions2.svdfunc_pymnID = pymnid;
            //    savedfunctions2.svdfunc_UserSaveID = userid2;
            //    savedfunctions2.svdfunc_IsSubmmit = false;
            //    savedfunctions2.svdfunc_ID = exist.svdfunc_ID;

            //    var id2 = savedfunctionRepo.UpdateReturnID(savedfunctions2);

            //    if (id2 != null)
            //    {
            //        foreach (var item in lstMoalefeexcel)
            //        {
            //            var b =await db.tbReffrenceSaveLevel.Where(p => p.ID == Bastena && p.Deleted != true).FirstOrDefaultAsync();
            //            var update =await db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKUser == item.MoalfeVal_FKUser && p.MoalfeVal_Month == item.MoalfeVal_Month && p.MoalfeVal_Year == item.MoalfeVal_Year && p.MoalfeVal_FKMoalafeDastmozdi == item.MoalfeVal_FKMoalafeDastmozdi).FirstOrDefaultAsync();
            //            if (update.MoalfeVal_Value != item.MoalfeVal_Value)
            //            {

            //                //InsertInLog
            //                tbLogFunctions tblogfunc = new tbLogFunctions
            //                {
            //                    lgfunc_BastehID = b.FK_RRSave,
            //                    lgfunc_DateTime = DateTime.Now,
            //                    lgfunc_MoalefeID = item.MoalfeVal_FKMoalafeDastmozdi,
            //                    lgfunc_NewValue = item.MoalfeVal_Value,
            //                    lgfunc_OldValue = update.MoalfeVal_Value,
            //                    lgfunc_pymnID = pymnid,
            //                    lgfunc_UserID = userid2
            //                };
            //                var checkInsertInLogFunction = logfunc.InsertToLogFunctionTable(tblogfunc);
            //                update.MoalfeVal_Value = item.MoalfeVal_Value;
            //                update.FK_SavedFunctionsID = id2;

            //                ;


            //                try
            //                {
            //                    db.SaveChanges();

            //                }
            //                catch (Exception e)
            //                {
            //                }



            //            }


            //        }
            //    }


            //    else
            //    {

            //        return Content("شرایطی برای اجرای عملیات وجود ندارد");
            //    }

            //}
            //else//add
            //{

            //}


            return Content("true");
        }


        //public void StartTimer()
        //{
        //    // ایجاد یک تایمر که هر 5 دقیقه یکبار اجرا شود
        //    _timer = new Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromMinutes(5));
        //}

        //private void TimerCallback(object state)
        //{
        //    // کدی که باید هر 5 دقیقه یکبار اجرا شود
        //    createexcel();
        //}
        //public void StartTimer()
        //{
        //    // ایجاد یک تایمر که هر 3 ساعت یکبار اجرا شود
        //    _timer = new Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromMinutes(5));
        //}

        //private void TimerCallback(object state)
        //{
        //    // کدی که باید هر 3 ساعت یکبار اجرا شود
        //    createexcel();
        //}

        public async Task<ActionResult> createexcel()
      {
            var excel = db.tbSavedFunctions.Where(p => p.svdfunc_FileSystemNameExcel == null).FirstOrDefault();
            
            if (excel != null)
            {
                var fin = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.FK_SavedFunctionsID == excel.svdfunc_ID).FirstOrDefault();
                if (fin != null)
                {
                    Task.Run(() =>
                    {

                        ExportMoalefeExcel3(excel.svdfunc_ID);
                    });
                }
             
            }
        

            return Content("kfd");
        }


        public string creatForzarib2(tbPeymanZaribForFish zarib)
        {
            // Create an instance of tbPeymanZaribForFish

            return ZaribforFish.Create(zarib);

            // Do something with the created instance if needed

            // Return a string if necessary

        }

        public string sabtpavast(IEnumerable<HttpPostedFileBase> files = null)
        {
            tbSavedFunctions savedfunctions = new tbSavedFunctions();
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
                        savedfunctions.svdfunc_FileSystemName = filename;
                        savedfunctions.svdfunc_FileName = file.FileName;

                        return savedfunctionRepo.Create(savedfunctions).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            return "True";
        }








        public string cHEAK(int MoalfeVal_Year, int MoalfeVal_Month, int Basteh, int baste_peyman1)
        {
            int userid = 0;
            var User = new tbUsers();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

            if (cookie_user != null)
            {
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

            List<tbSavedFunctions> time1 = new List<tbSavedFunctions>();
            var ex = db.tbSavedFunctions
                .Where(p => p.svdfunc_pymnID == baste_peyman1 && p.svdfunc_BastehID == Basteh && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == MoalfeVal_Month && s.MoalfeVal_Year == MoalfeVal_Year))
                .ToList();

            if (ex.Count != 0)
            {
                foreach (var item in ex)
                {
                    if (item.Final_Accept == true)
                    {
                        return "False2";
                    }


                }

                // If the loop is not entered, return a default value
                return "true";
            }
            else
            {
                return "true";
            }
        }

        public ActionResult onesabtmolfe(int MoalfeVal_FKMoalafeDastmozdi, int MoalfeVal_FKUser, int MoalfeVal_Year, float MoalfeVal_Value, int MoalfeVal_Month, int Basteh, DateTime FromDate, DateTime ToDate, int baste_peyman1)
        {










            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();
            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
            {
                MoalfeVal_FKMoalafeDastmozdi = MoalfeVal_FKMoalafeDastmozdi,
                MoalfeVal_FKUser = MoalfeVal_FKUser,
                MoalfeVal_Month = System.Convert.ToInt32(MoalfeVal_Month),
                MoalfeVal_Year = System.Convert.ToInt32(MoalfeVal_Year),
                MoalfeVal_Value = System.Convert.ToDouble(MoalfeVal_Value),


            };

            lstMoalefeexcel.Add(dastmozdexcel);
            if (moalefeexcelRepo.Create(dastmozdexcel) == "true")
            {
            }
            int userid2 = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid2 = User.usr_ID;
                    }
                }
            }
            var exist = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == baste_peyman1 && p.svdfunc_BastehID == Basteh && p.svdfunc_FromDate == FromDate && p.svdfunc_ToDate == ToDate).FirstOrDefault();
            if (exist != null)//update
            {
                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                savedfunctions.svdfunc_BastehID = Basteh;
                savedfunctions.svdfunc_FromDate = FromDate;
                savedfunctions.svdfunc_ToDate = ToDate;
                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                savedfunctions.svdfunc_pymnID = baste_peyman1;
                savedfunctions.svdfunc_UserSaveID = userid2;
                savedfunctions.svdfunc_IsSubmmit = false;
                savedfunctions.svdfunc_ID = exist.svdfunc_ID;


                var id2 = savedfunctionRepo.UpdateReturnID(savedfunctions);

                if (id2 != null)
                {
                    foreach (var item in lstMoalefeexcel)
                    {
                        var update = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKUser == item.MoalfeVal_FKUser && p.MoalfeVal_Month == item.MoalfeVal_Month && p.MoalfeVal_Year == item.MoalfeVal_Year && p.MoalfeVal_FKMoalafeDastmozdi == item.MoalfeVal_FKMoalafeDastmozdi).FirstOrDefault();
                        if (update.MoalfeVal_Value != item.MoalfeVal_Value)
                        {
                            //InsertInLog
                            tbLogFunctions tblogfunc = new tbLogFunctions
                            {
                                lgfunc_BastehID = Basteh,
                                lgfunc_DateTime = DateTime.Now,
                                lgfunc_MoalefeID = item.MoalfeVal_FKMoalafeDastmozdi,
                                lgfunc_NewValue = item.MoalfeVal_Value,
                                lgfunc_OldValue = update.MoalfeVal_Value,
                                lgfunc_pymnID = baste_peyman1,
                                lgfunc_UserID = userid2
                            };

                            var checkInsertInLogFunction = logfunc.InsertToLogFunctionTable(tblogfunc);
                            update.MoalfeVal_Value = item.MoalfeVal_Value;
                            update.FK_SavedFunctionsID = id2;
                            db.SaveChanges();

                        }


                    }
                }



                else
                {

                    return Content("شرایطی برای اجرای عملیات وجود ندارد");
                }


            }
            else//add
            {
                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                savedfunctions.svdfunc_BastehID = Basteh;
                savedfunctions.svdfunc_FromDate = FromDate;
                savedfunctions.svdfunc_ToDate = ToDate;
                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                savedfunctions.svdfunc_pymnID = baste_peyman1;
                savedfunctions.svdfunc_UserSaveID = userid2;
                savedfunctions.svdfunc_IsSubmmit = false;
                savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;

                if (savedfunctionRepo.Create(savedfunctions) == "True")
                {
                }
                else
                {

                    return Content("در ذخیره سازی مشکلی به وجود آمده است");
                }
            }

            return Content("true"); ;
        }









        public string sabtpavast1(IEnumerable<HttpPostedFileBase> files = null, DateTime FromDate = default, DateTime ToDate = default, int PeymanID = 0, int Basteh = 0)
        {
            int userid2 = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        userid2 = User.usr_ID;
                    }
                }
            }
            var exist = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == PeymanID && p.svdfunc_BastehID == Basteh && p.svdfunc_FromDate == FromDate && p.svdfunc_ToDate == ToDate).FirstOrDefault();
            if (exist != null)
            {//update
                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                savedfunctions.svdfunc_BastehID = Basteh;
                savedfunctions.svdfunc_FromDate = FromDate;
                savedfunctions.svdfunc_ToDate = ToDate;
                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                savedfunctions.svdfunc_pymnID = PeymanID;
                savedfunctions.svdfunc_UserSaveID = userid2;
                savedfunctions.svdfunc_IsSubmmit = false;
                savedfunctions.svdfunc_ID = exist.svdfunc_ID;
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
                            savedfunctions.svdfunc_FileSystemName = filename;
                            savedfunctions.svdfunc_FileName = file.FileName;
                            savedfunctionRepo.Update2(savedfunctions);

                        }

                        else
                        {
                            savedfunctionRepo.Update2(savedfunctions);

                        }


                    }
                }
;
            }

            else//add
            {
                tbSavedFunctions savedfunctions = new tbSavedFunctions();
                savedfunctions.svdfunc_BastehID = Basteh;
                savedfunctions.svdfunc_FromDate = FromDate;
                savedfunctions.svdfunc_ToDate = ToDate;
                savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                savedfunctions.svdfunc_pymnID = PeymanID;
                savedfunctions.svdfunc_UserSaveID = userid2;
                savedfunctions.svdfunc_IsSubmmit = false;

                foreach (var file in files)
                {
                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/" + filename));
                        savedfunctions.svdfunc_FileSystemName = filename;
                        savedfunctions.svdfunc_FileName = file.FileName;
                        savedfunctionRepo.Create(savedfunctions);

                    }
                    else
                    {
                        savedfunctionRepo.Create(savedfunctions);
                    }

                }


                return "True";
            }
            return "True";

        }





        #endregion



        // GET: Salaries/Moalefe


        [AuthorizeAAA]
        public string CheckFromDate(DateTime FromDate, int Basteh)
        {
            int? TypeBasteh = registerController.ReffrenceSave_GetTypeOfBaste(Basteh);
            var savedfunc = savedfunctionRepo.Update();
            if (savedfunc.Count != 0)
            {
                var result = savedfunc.Where(p => p.svdfunc_BastehID == Basteh).ToList();
                if (TypeBasteh == 1)//تجمعی
                {
                    if (FromDate != result.OrderByDescending(p => p.svdfunc_ToDate).FirstOrDefault().svdfunc_FromDate)
                    {
                        return "بسته مورد نظر به صورت تجمعی تعریف شده است و بایستی تاریخ شروع با دوره های قبل یکسان باشد";
                    }
                }
                else if (TypeBasteh == 2)//متوالی
                {
                    if (FromDate <= result.OrderByDescending(p => p.svdfunc_ToDate).FirstOrDefault().svdfunc_ToDate)
                    {
                        return "بسته مورد نظر به صورت متوالی تعریف شده است و بایستی تاریخ شروع بعد از تاریخ پایان دوره قبلی باشد";
                    }
                }
            }
            return "ok";

        }




    }
}

