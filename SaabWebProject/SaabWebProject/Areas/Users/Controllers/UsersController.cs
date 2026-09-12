using Access.Models.UnitofWorks.Repository;
using Newtonsoft.Json;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Xml;
using Telerik.Web.Spreadsheet;
using static SaabWebProject.Areas.Users.Controllers.MessageBoxController;
using SaabWebProject.Models.ViewModels.Users;
using SaabWebProject.Utility;
using Syncfusion.XlsIO;
using SaabWebProject.Models.Utilitis;
using Microsoft.Office.Interop.Excel;
using SaabWebProject.Models.Repositories.Contracts;

namespace SaabWebProject.Areas.Users.Controllers
{
    public class UsersController : Controller
    {
        #region متغیر ها

        IRepository<tbUsers> UserRepo;
        IUnitofWork unitofwork;
        SaabEntities db;
        tbUsersRepository rep_users;
        tbJobsRepository jobsRepository;
        BusinessSideRepository rep_business;
        tblinkUserAndPeymanRepository linkuserandpeymanRepo;
        public UsersController()
        {
            db = new SaabEntities();
            UserRepo = new SaabWebProject.Models.UnitOfWork.Repositories.Repository<tbUsers>(db);
            unitofwork = new UnitofWork(db);
            rep_users = new tbUsersRepository();
            jobsRepository = new tbJobsRepository();
            rep_business = new BusinessSideRepository(db);
            linkuserandpeymanRepo = new tblinkUserAndPeymanRepository(db);
        }

        #endregion

        #region صفحات


        [AuthorizeAAA]
        public ActionResult _PartialViewUsersListWithSearch()
        {
            return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearch.cshtml", rep_users.Update().OrderByDescending(p => p.usr_ID).ToList());
        }

        [AuthorizeAAA]
        public ActionResult _PartialViewUsersListWithSearchAndMultiple(int positions = 0)
        {
            if (positions == 0)
                return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", rep_users.Update().OrderByDescending(p => p.usr_ID).ToList());
            else
                return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", rep_users.GetAllUsersInSamePosition(positions));
        }

        public ActionResult editafrad(int positions = 0)
        {
            if (positions == 0)
            {
                return PartialView("~/Areas/Users/Views/Users/editafrad.cshtml", rep_users.Update().OrderByDescending(p => p.usr_ID).ToList());

            }
            else
            {
                var fi = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == positions).FirstOrDefault();

                List<tbUsers> us = new List<tbUsers>();
                var findusr = db.tbReffrenceSaveLevelUser.Where(p => p.FK_LevelID == fi.ID).ToList();
                foreach (var item in findusr)
                {
                    us.Add(item.tbUsers);
                }
                return PartialView("~/Areas/Users/Views/Users/editafrad.cshtml", us);
            }

        }
        public ActionResult editafrad2(int positions = 0)
        {

            if (positions == 0)
            {

                return PartialView("~/Areas/Users/Views/Users/editafrad2.cshtml", rep_users.Update().OrderByDescending(p => p.usr_ID).ToList());

            }
            else
            {
                var fi = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == positions).FirstOrDefault();

                List<tbUsers> us = new List<tbUsers>();
                var findusr = db.tbReffrenceSaveLevelUser.Where(p => p.FK_LevelID == fi.ID).ToList();
                foreach (var item in findusr)
                {
                    us.Add(item.tbUsers);
                }
                return PartialView("~/Areas/Users/Views/Users/editafrad2.cshtml", us);
            }

        }
        public ActionResult usrtbpersin()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    return PartialView(fin);

                }
            }
            else
            {
                var fi = db.dbrelatinperson.ToList();
                return PartialView(fi);

            }

        }
        public ActionResult usrfaliatsabt()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    return PartialView(fin);

                }
            }
            else
            {
                var fi = db.dbrelatinperson.ToList();
                return PartialView(fi);

            }

        }
        [AuthorizeAAA]
        public ActionResult _PartialViewWithPersonalyCode()
        {
            return PartialView("~/Areas/Users/Views/Users/_PartialViewWithPersonalyCode.cshtml", rep_users.Update().OrderByDescending(p => p.usr_ID).ToList());
        }
        /// <summary>
        /// پارشیال ویو مربوط به ایجاد کاربر
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _CreateUser()
        {
            return PartialView();
        }

        /// <summary>
        /// پارشیال ویو مربوط به ویرایش کاربر
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _EditUser(int id)
        {
            var user = rep_users.Find(id);
            return PartialView(user);
        }

        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// پارشیال ویو کاربران در صفحه اصلی مدیریت کاربران
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult GetAllUsersList()
        {
            var list = rep_users.Update().Where(p => p.usr_IsUser != false).OrderByDescending(p => p.usr_ID).ToList();

            return PartialView("~/Areas/Users/Views/Users/_PartialViewUserList.cshtml", list);
        }
        public ActionResult GetFilter(int id)
        {
            List<tbUsers> list = new List<tbUsers>();
            var list2 = db.tbUser_link_BusinessSide.Where(p => p.FK_up_ID == id).ToList();
            foreach (var item in list2)
            {
                var list23 = db.tbUsers.Where(p => p.usr_IsUser != false && p.usr_ID == item.FK_u_ID).OrderByDescending(p => p.usr_ID).FirstOrDefault();
                if (list23 != null)
                {
                    list.Add(list23);
                }
            }
            //list=db.tbUsers.Where(p => p.usr_IsUser != false&&p.tbUser_link_BusinessSide.FK_u_ID).OrderByDescending(p => p.usr_ID).ToList();
            return PartialView("~/Areas/Users/Views/Users/_PartialViewUserList.cshtml", list);
        }

        [AuthorizeAAA]
        public ActionResult _Cities_Select()
        {
            var obj = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
            //using(var db=new SaabEntities())
            //{
            //    List<tbCities> mylist = new List<tbCities>();
            //    foreach(var item in obj)
            //    {
            //        tbCities ob = new tbCities();
            //        ob.ID = item.cityId;
            //        ob.Name = item.cityName;
            //        mylist.Add(ob);

            //    }
            //    db.tbCities.AddRange(mylist);
            //    db.SaveChanges();
            //}
            return PartialView(obj);
        }
        //مالتی سلکت همه افراد فعال

        [AuthorizeAAA]
        public ActionResult _GetAllUsers_Select()
        {
            var a = rep_users.ActiveList();
            return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", a);
        }
        public ActionResult viewfish_pyman()
        {
            return PartialView("~/Areas/Users/Views/Users/viewfish_pyman.cshtml");
        }
        public ActionResult _GetAllUsers_Select3()
        {
            var cookieUser = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookieUser == null)
            {
                return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", new List<tbUsers>());
            }

            var nationalCode = Utility.Base64.Base64Decode(cookieUser.Value);

            var userId = db.tbUsers
                .Where(u => u.usr_NationalCode == nationalCode)
                .Select(u => u.usr_ID)
                .FirstOrDefault();

            if (userId == 0)
            {
                return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", new List<tbUsers>());
            }

            var relatedUserIds = db.tbAdamAdam
                .Where(x => x.Name == userId && x.Value == 1)
                .Select(x => x.nam)
                .Distinct()
                .ToList();

            var users = db.tbUsers
                .Where(u => relatedUserIds.Contains(u.usr_ID))
                .ToList();

            return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", users);
        }

        public ActionResult _GetAllUsers_Select2()
        {
            var a = rep_users.ActiveList();
            return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", a);
        }
        //مالتی سلکت همه افراد  حقوقی 

        [AuthorizeAAA]
        public ActionResult _HoghoghiUserList()
        {
            var Model = rep_users.HoghooghiUser();
            if (Model != null)
            {
                return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", new List<tbUsers>());
            }
        }
        public ActionResult Getuser(int month, int year, int month2)
        {

            var today = DateTime.Now.GetShamsiDayOfMonth();
            var currentTime = DateTime.Now.TimeOfDay;

            var x2 = db.tbUserContracts
     .Where(p => p.usc_StartTime.HasValue)
     .Select(s => s.usc_StartTime)
     .ToList();



            var ex = x2
                .Where(date => /*date?.GetShamsiMonth() >= month  && */date?.GetShamsYear() == year)
            .ToList();
            List<tbUserContracts> usr2 = new List<tbUserContracts>();
            List<tbUserContracts> usr3 = new List<tbUserContracts>();

            List<tbUsers> usr4 = new List<tbUsers>();
            var uniqueIds2 = new List<int>();

            foreach (var item in ex)
            {
                var ec = db.tbUserContracts.Where(p => p.usc_StartTime == item && !uniqueIds2.Contains(p.usc_ID)).FirstOrDefault();
                if (ec != null)
                {
                    uniqueIds2.Add(ec.usc_ID);
                    usr2.Add(ec);
                }
            }
            var c = usr2.Select(s => s.usc_EndTime).ToList();





            var ex2 = c
               .Where(date => date?.GetShamsiMonth() >= month2 && date?.GetShamsYear() == year)
           .ToList();
            var uniqueIds = new List<int>();
            foreach (var item in ex2)
            {
                var ec = db.tbUserContracts.Where(p => p.usc_EndTime == item && !uniqueIds.Contains(p.usc_ID)).FirstOrDefault();
                if (ec != null)
                {
                    uniqueIds.Add(ec.usc_ID);
                    usr3.Add(ec);
                }
            }

            var ex4 = usr3.Select(s => s.FK_UserID).ToList();
            foreach (var item in ex4)
            {
                var ec = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                usr4.Add(ec);
            }

            return PartialView("~/Areas/Users/Views/Users/Getuser.cshtml", usr4);

            // Rest of your code...
        }


        //مالتی سلکت همه افراد  حقیقی 
        [AuthorizeAAA]
        public ActionResult _GetAllHaghighiUsers_Select()
        {
            var a = rep_users.HaghighiUser();
            return PartialView("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", a);
        }
        //سینگل سلکت همه افراد  حقوقی 

        [AuthorizeAAA]
        public ActionResult _GetAllHoghooghiUsers_Select()
        {
            var Model = rep_users.HoghooghiUser();
            return PartialView("~/Areas/Users/Views/Users/_PartialViewListUsers.cshtml", Model);
        }



        public ActionResult _GetusersBaseOnPeymanID(int PeymanID)
        {
            var Model = linkuserandpeymanRepo.GetUsers(PeymanID);
            return PartialView("~/Areas/Users/Views/Users/_GetusersBaseOnPeymanID.cshtml", Model);
        }
        #endregion

        #region توابع

        [HttpGet]
        [AuthorizeAAA]
        public ActionResult Export_Users()
        {
            var OutPutFile = SetDataExcel_Users();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Users" + extension);
        }

        private Telerik.Web.Spreadsheet.Workbook SetDataExcel_Users()
        {
            var NamingTagssampleFile = Telerik.Web.Spreadsheet.Workbook.Load(Server.MapPath("~/Content/ExcelFiles/userSample.xlsx"));
            Row Row;
            using (SaabEntities db = new SaabEntities())
            {
                var jobInfo = db.tbBusinessSide.ToList();
                if (jobInfo.Count() > 0)
                {
                    int counter = 1;
                    foreach (var item in jobInfo)
                    {
                        Row = new Row() { Height = 20, Index = counter };

                        Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.up_name,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 0
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }

                int tginf = 1;

                var userInfo = UserRepo.Get(p => p.usr_IsActive == true).ToList();
                if (userInfo.Count > 0)
                {
                    foreach (var item in userInfo)
                    {
                        var t = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == item.usr_ID && p.Status == true).FirstOrDefault();
                        var ty = db.tbUsers.Where(p => p.usr_ID == item.usr_ID).FirstOrDefault();
                        string cit = "";
                        int aman = 0;
                        int ghardad = 0;
                        string job = "";
                        if (item.tbjob != null)
                        {
                            job = item.tbjob.Name;
                        }
                        if (ty != null)
                        {
                            var r = db.tbUserContracts.Where(p => p.FK_UserID == ty.usr_ID).FirstOrDefault();
                            if (r != null)
                            {
                                ghardad = 1;
                            }
                        }
                        if (ty.usr_amani == true)
                        {
                            aman = 1;
                        }
                        if (ty.usr_City_Dutysystem != null)
                        {
                            var city = db.tbCities.Where(p => p.ID == ty.usr_City_Dutysystem).FirstOrDefault();

                            if (city != null)
                            {
                                cit = city.Name;
                            }

                        }
                        string pecTitle = "";
                        if (t == null)
                        {
                            // Skip this iteration if t is null
                            //pecTitle = "به پیمانی متصل نیست";
                        }
                        else
                        {
                            pecTitle = t.tbPeymanContracts?.pec_Title ?? "N/A"; // یا هر مقدار پیش‌فرض دیگری که مناسب باشد

                        }

                        Row = new Row() { Height = 20, Index = tginf };

                        var cells = new List<Cell>
        {
            new Cell
            {
                Value = item.FullName,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 1
            },
            new Cell
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
            },
               new Cell
            {
                Value = item.usr_NationalCode ??"0",
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 3
            },
                     new Cell
            {
                Value =  cit,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 4
            },

            new Cell
            {
                Value = aman,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 5
            },
                  new Cell
            {
                Value = ghardad,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 6
            },
            new Cell
            {
                Value = pecTitle,
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
            new Cell
            {
                Value = job,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = true,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 8
            }
        };
                        Row.AddCells(cells);

                        tginf++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }


                var Cities = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
                if (Cities.Count >= 1)
                {
                    int counter = 1;
                    foreach (var item in Cities)
                    {
                        Row = new Row() { Height = 20, Index = counter };

                        Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.cityName,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 13
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }

                var peymans = db.tbPeymanContracts.Where(p => p.Inactive != true).ToList();
                if (peymans.Count >= 1)
                {
                    int counter = 1;
                    foreach (var item in peymans)
                    {
                        Row = new Row() { Height = 20, Index = counter };

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
                                Index = 14
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }
                var jobasli = db.tbjob.ToList();
                if (jobasli.Count >= 1)
                {
                    int counter = 1;
                    foreach (var item in jobasli)
                    {
                        Row = new Row() { Height = 20, Index = counter };

                        Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.Name+"_"+item.ID,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 16
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }

                var detailjpb = db.tbdetailjob.ToList();
                if (detailjpb.Count >= 1)
                {
                    int counter = 1;
                    foreach (var item in detailjpb)
                    {
                        Row = new Row() { Height = 20, Index = counter };

                        Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.detail_title +"_"+item.ID,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 17
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }
                //int tginf2 = 1;

                //var userInfo2 = UserRepo.Get(p => p.usr_IsActive == true).ToList();
                //if (userInfo2.Count() > 0)
                //{
                //    foreach (var item in userInfo2)
                //    {
                //        Row = new Row() { Height = 20, Index = tginf };

                //        Row.AddCells(new List<Cell>()
                //        {
                //            new Cell()
                //            {
                //                Value = item.usr_Personal_ID,
                //                FontFamily = "B Nazanin",
                //                Bold = false,
                //                Enable = true,
                //                Wrap = false,
                //                FontSize = 12,
                //                Italic = false,
                //                Underline = false,
                //                Index = 9
                //            }
                //        });
                //        tginf2++;
                //        NamingTagssampleFile.Sheets[1].AddRow(Row);
                //    }
                //}


            }

            return NamingTagssampleFile;
        }


        [AuthorizeAAA]
        [HttpPost]
        public ActionResult ImportExel_Users(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream == null)
            {
                return MessageBox.Show("فایل بطور صحیح بارگذاری نشده است", MessageType.Error, false, MessageAlignment.TopRight);
            }

            string Message = GetDataFromExcel_Users(MyExcelStream);
            return Content(Message);
        }
        //تابع دریافت اکسل
        public ActionResult cratausrpassword()
        {
            var finddd = db.tbUsers
                .Where(s => s.tbUser_link_BusinessSide.Any(s1 => s1.FK_up_ID == 2175 && s1.Status == true)
                            && s.usr_NationalCode != null
                            && s.usr_Personal_ID != null)
                .ToList();

            foreach (var it in finddd)
            {
                // usr_NationalCode رشته است
                string nationalCodeStr = it.usr_NationalCode;

                // usr_Personal_ID عدد nullable است → باید ToString()
                string personalIdStr = it.usr_Personal_ID.HasValue
                    ? it.usr_Personal_ID.Value.ToString()
                    : "";

                // گرفتن 4 رقم آخر
                string nationalCodePart = nationalCodeStr.Length >= 4
                    ? nationalCodeStr.Substring(nationalCodeStr.Length - 4)
                    : nationalCodeStr;

                string personalIdPart = personalIdStr.Length >= 4
                    ? personalIdStr.Substring(personalIdStr.Length - 4)
                    : personalIdStr;

                // ساختن کد 8 رقمی
                string rawPassword = nationalCodePart + personalIdPart;

                // هش کردن (MD5)
                it.usr_Password = FormsAuthentication.HashPasswordForStoringInConfigFile(rawPassword, "MD5");
            }

            db.SaveChanges();

            return Content("Passwords updated successfully.");
        }



        public string GetDataFromExcel_Users(HttpPostedFileBase MyExcelStream)
        {
            // var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));

            using (SaabEntities db = new SaabEntities())
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    using (ExcelEngine xl = new ExcelEngine())
                    {
                        try
                        {
                            IApplication app = xl.Excel;
                            var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                            var Count_Sheets = workbook.Worksheets.Count;
                            if (Count_Sheets > 1)
                            {
                                var sheet = workbook.Worksheets[0];
                                var Rows = sheet.Rows;
                                int Cols = 0;
                                if (Rows.Length >= 1)
                                {
                                    Cols = Rows[0].Cells.Length;

                                    //تعداد ستون ها
                                    if (Count_Sheets != 2 || Cols != 28)
                                    {
                                        return "اکسل وارد شده با اکسل نمونه  متفاوت است";
                                    }

                                    var title = workbook.Worksheets[0].Rows[0].Cells;
                                    var count = workbook.Worksheets[0].Rows.Count();
                                    if (count == 1)
                                    {
                                        return "فایل اکسل فاقد اطلاعات می باشد";
                                    }

                                    Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                    int counter = 0;
                                    foreach (var item in title)
                                    {
                                        dicTitles.Add(item.Value.ToString(), counter);
                                        counter++;
                                    }

                                    // List<string> AddedIWNCode = new List<string>();
                                    List<tbUsers> myList = new List<tbUsers>();
                                    for (int i = 1; i < count; i++)
                                    {
                                        tbUsers Us;
                                        tbUsers User = new tbUsers();
                                        var row = workbook.Worksheets[0].Rows[i];
                                        var cell_code = row.Cells[11];

                                        if (cell_code != null)
                                        {
                                            if (cell_code.Value != null || cell_code.Value != "")
                                            {
                                                User.usr_NationalCode = cell_code.Value;
                                                Us = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == User.usr_NationalCode);
                                                if (Us != null)
                                                {
                                                    User = Us;
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون کد ملی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون کد ملی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_family = row.Cells[0];
                                        if (cell_family != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_family.Value))
                                            {
                                                User.usr_Family = cell_family.Value;
                                            }
                                            else
                                            {
                                                return " لطفا ستون نام خانوادگی سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }
                                        }

                                        var cell_name = row.Cells[1];

                                        if (cell_name != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_name.Value))
                                            {
                                                User.usr_Name = cell_name.Value;
                                            }
                                            else
                                            {
                                                return " لطفا ستون نام سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }
                                        }
                                        else
                                        {
                                            return " لطفا ستون نام سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        }

                                        var cell_Shcode = row.Cells[2];
                                        if (cell_Shcode != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Shcode.Value))
                                            {
                                                try
                                                {
                                                    User.usr_SHCode = Convert.ToInt64(cell_Shcode.Value);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون شماره شناسنامه سطر " + " ( " + i + " ) " + "را درست پر کنید ";
                                                }
                                            }
                                        }


                                        var cell_Address = row.Cells[3];

                                        if (cell_Address != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Address.Value))
                                            {
                                                User.usr_Address = cell_Address.Value.ToString();
                                            }
                                        }

                                        var cell_Email = row.Cells[4];

                                        if (cell_Email != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Email.Value))
                                            {
                                                User.usr_Email = cell_Email.Value.ToString();
                                            }
                                        }

                                        var cell_Fathername = row.Cells[5];

                                        if (cell_Fathername != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Fathername.Value))
                                            {
                                                User.usr_FatherName = cell_Fathername.Value;
                                            }
                                        }

                                        var cell_dutysystem = row.Cells[6];

                                        if (cell_dutysystem != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_dutysystem.Value))
                                            {
                                                var Dutysystem = cell_dutysystem.Value.ToString();
                                                if (Dutysystem == "مشمول")
                                                {
                                                    Dutysystem = 1.ToString();
                                                }
                                                else if (Dutysystem == "پایان خدمت")
                                                {
                                                    Dutysystem = 2.ToString();
                                                }
                                                else if (Dutysystem == "معاف")
                                                {
                                                    Dutysystem = 3.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون نظام وظیفه سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_Dutysystem = Convert.ToInt32(Dutysystem);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون نظام وظیفه سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_Gender = row.Cells[7];

                                        if (cell_Gender != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Gender.Value))
                                            {
                                                var gender = cell_Gender.Value;
                                                if (gender == "زن")
                                                {
                                                    gender = 1.ToString();
                                                }
                                                else if (gender == "مرد")
                                                {
                                                    gender = 2.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون جنسیت سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_Gender = Convert.ToInt32(gender);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون جنسیت سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_marial = row.Cells[8];

                                        if (cell_marial != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_marial.Value))
                                            {
                                                var maritaIStatus = cell_marial.Value;
                                                if (maritaIStatus == "متاهل")
                                                {
                                                    maritaIStatus = 1.ToString();
                                                }
                                                else if (maritaIStatus == "مجرد")
                                                {
                                                    maritaIStatus = 2.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون وضعیت تاهل سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_MaritaIStatus = Convert.ToInt32(maritaIStatus);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون وضعیت تاهل سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_phonenumber = row.Cells[9];

                                        if (cell_phonenumber != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_phonenumber.Value))
                                            {
                                                try
                                                {
                                                    User.usr_PhoneNumber = Convert.ToInt64(cell_phonenumber.Value);
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره همراه سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_PlaceOfIssue = row.Cells[10];

                                        if (cell_PlaceOfIssue != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_PlaceOfIssue.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName == cell_PlaceOfIssue.Value.ToString());

                                                if (city != null)
                                                {
                                                    User.usr_PlaceOfIssue = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون محل صدور سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_Pass = row.Cells[12];

                                        if (!string.IsNullOrEmpty(cell_Pass.Value))
                                        {
                                            if (cell_Pass.Value != null)
                                            {
                                                User.usr_Password = FormsAuthentication.HashPasswordForStoringInConfigFile(cell_Pass.Value, "MD5");
                                            }
                                        }



                                        var cell_job = row.Cells[13];

                                        if (cell_job != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_job.Value))

                                            {
                                                var job = rep_business.IfExitsPosition(cell_job.Value);
                                                if (job != 0 || job != null)
                                                {
                                                    if (User.tbUser_link_BusinessSide != null)
                                                    {
                                                        foreach (var item in User.tbUser_link_BusinessSide.ToList())
                                                        {
                                                            item.Status = false;
                                                        }
                                                    }

                                                    tbUser_link_BusinessSide new_link = new tbUser_link_BusinessSide();
                                                    new_link.Status = true;
                                                    new_link.FK_up_ID = job;
                                                    tbLink_UserAndAction l2 = new tbLink_UserAndAction();
                                                    var ex = db.tbActions.FirstOrDefault();
                                                    if (ex != null)
                                                    {
                                                        l2.HasAccess = true;
                                                        l2.FK_Position_ID = job;
                                                        l2.FK_Action_ID = ex.Action_ID;
                                                        User.tbLink_UserAndAction.Add(l2);

                                                    }

                                                    User.tbUser_link_BusinessSide.Add(new_link);

                                                }
                                                else
                                                {
                                                    return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_regnum = row.Cells[14];

                                        if (cell_regnum != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_regnum.Value))
                                            {
                                                try
                                                {
                                                    User.usr_RegistrationNumber = Convert.ToInt64(cell_regnum.Value);
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره ثابت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_brDate = row.Cells[15];

                                        if (cell_brDate != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_brDate.Value))
                                            {
                                                var Date = cell_brDate.Value.ToString().Split('/').ToList();
                                                if (Date != null)
                                                {
                                                    try
                                                    {
                                                        PersianCalendar pc = new PersianCalendar();
                                                        int year = int.Parse(Date[0]);
                                                        int month = int.Parse(Date[1]);
                                                        int day = int.Parse(Date[2]);
                                                        var datetime = pc.ToDateTime(year, month, day, 0, 0, 0, 0);
                                                        User.usr_DateOfBrith = datetime;
                                                    }
                                                    catch
                                                    {
                                                        return "  لطفا ستون تاریخ تولد سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                    }
                                                }
                                            }
                                        }

                                        var cell_placeOfdate = row.Cells[16];

                                        if (cell_placeOfdate != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_placeOfdate.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName == cell_placeOfdate.Value);
                                                if (city != null)
                                                {
                                                    User.usr_PlaceOfBrith = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون محل تولد سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_child = row.Cells[17];

                                        if (cell_child != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_child.Value))
                                            {
                                                try
                                                {
                                                    User.usr_Child_Allowance = Convert.ToInt32(cell_child.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون تعداد حق اولاد مشمول بیمه سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_degg = row.Cells[18];

                                        if (cell_degg != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_degg.Value))
                                            {
                                                var degree = cell_degg.Value.ToString();
                                                switch (degree)
                                                {
                                                    case "بیسواد":
                                                        {
                                                            User.usr_Degree = 1;
                                                            break;
                                                        }
                                                    case "سیکل":
                                                        {
                                                            User.usr_Degree = 2;
                                                            break;
                                                        }
                                                    case "دیپلم":
                                                        {
                                                            User.usr_Degree = 3;
                                                            break;
                                                        }
                                                    case "کاردانی":
                                                        {
                                                            User.usr_Degree = 4;
                                                            break;
                                                        }
                                                    case "کارشناسی":
                                                        {
                                                            User.usr_Degree = 5;
                                                            break;
                                                        }
                                                    case "کارشناسی ارشد":
                                                        {
                                                            User.usr_Degree = 6;
                                                            break;
                                                        }
                                                    case "دکتری":
                                                        {
                                                            User.usr_Degree = 7;
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        var cell_gr = row.Cells[19];

                                        if (cell_gr != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_gr.Value))
                                            {
                                                var grade = cell_gr.Value.ToString();
                                                switch (grade)
                                                {
                                                    case "بیسواد":
                                                        {
                                                            User.usr_Grade = 1;
                                                            break;
                                                        }
                                                    case "سیکل":
                                                        {
                                                            User.usr_Grade = 2;
                                                            break;
                                                        }
                                                    case "دیپلم":
                                                        {
                                                            User.usr_Grade = 3;
                                                            break;
                                                        }
                                                    case "کاردانی":
                                                        {
                                                            User.usr_Grade = 4;
                                                            break;
                                                        }
                                                    case "کارشناسی":
                                                        {
                                                            User.usr_Grade = 5;
                                                            break;
                                                        }
                                                    case "کارشناسی ارشد":
                                                        {
                                                            User.usr_Grade = 6;
                                                            break;
                                                        }
                                                    case "دکتری":
                                                        {
                                                            User.usr_Grade = 7;
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        var cell_postal = row.Cells[20];

                                        if (cell_postal != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_postal.Value))
                                            {
                                                try
                                                {
                                                    User.usr_PostalCode = Convert.ToInt64(cell_postal.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون کد پستی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_City_Duty = row.Cells[21];

                                        if (cell_City_Duty != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_City_Duty.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName.Trim() == cell_City_Duty.Value.Trim());
                                                if (city != null)
                                                {
                                                    User.usr_City_Dutysystem = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون شهر محل خدمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_personal_code = row.Cells[22];

                                        if (cell_personal_code != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_personal_code.Value))
                                            {
                                                try
                                                {
                                                    User.usr_Personal_ID = Convert.ToInt32(cell_personal_code.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_all_child = row.Cells[23];

                                        if (cell_all_child != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_all_child.Value))
                                            {
                                                try
                                                {
                                                    User.usr_All_Child_Allowance = Convert.ToInt32(cell_all_child.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون تعداد کل اولاد سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_peymans = row.Cells[24];
                                        if (cell_peymans != null)
                                        {
                                            if (User.Link_User_And_Peyman != null)
                                            {
                                                foreach (var item in User.Link_User_And_Peyman.ToList())
                                                {
                                                    item.Status = false;
                                                    //db.SaveChanges();
                                                }
                                            }

                                            var cell_peymans_dat = row.Cells[25];

                                            if (cell_peymans_dat != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_peymans_dat.Value))
                                                {
                                                    var Date = cell_peymans_dat.Value.ToString().Split('/').ToList();
                                                    if (Date != null)
                                                    {
                                                        try
                                                        {
                                                            PersianCalendar pc = new PersianCalendar();
                                                            int year = int.Parse(Date[0]);
                                                            int month = int.Parse(Date[1]);
                                                            int day = int.Parse(Date[2]);
                                                            var datetime = pc.ToDateTime(year, month, day, 0, 0, 0, 0);
                                                            var pey = cell_peymans.Value;
                                                            if (!string.IsNullOrEmpty(pey))
                                                            {
                                                                if (pey.ToString().Contains(','))
                                                                {

                                                                    foreach (var item in pey.ToString().Split(',').ToList())
                                                                    {
                                                                        var peyman = db.tbPeymanContracts?.Where(p => p.pec_Title.Trim() == item && p.Inactive != true).FirstOrDefault();
                                                                        if (peyman != null)
                                                                        {
                                                                            Link_User_And_Peyman pyy = new Link_User_And_Peyman();
                                                                            pyy.Status = true;
                                                                            pyy.FK_Peyman_ID = peyman.pec_ID;
                                                                            User.Link_User_And_Peyman.Add(pyy);
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    var peyman = db.tbPeymanContracts.FirstOrDefault(p => p.pec_Title.Trim() == pey.ToString() && p.Inactive != true);
                                                                    if (peyman != null)
                                                                    {
                                                                        Link_User_And_Peyman pyy = new Link_User_And_Peyman();
                                                                        pyy.Status = true;
                                                                        pyy.datatime = datetime;

                                                                        pyy.FK_Peyman_ID = peyman.pec_ID;
                                                                        User.Link_User_And_Peyman.Add(pyy);
                                                                    }
                                                                    else
                                                                    {
                                                                        return "  لطفا ستون پیمان سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            return "  لطفا ستون تاریخ لینک پیمان سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                        var jobalis = row.Cells[26];

                                        if (jobalis != null)
                                        {
                                            if (!string.IsNullOrEmpty(jobalis.Value))
                                            {
                                                try
                                                {
                                                    User.usr_typeshoghl = Convert.ToInt32(jobalis.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }





                                        var jobalisdetail = row.Cells[27];

                                        if (jobalisdetail != null)
                                        {
                                            if (!string.IsNullOrEmpty(jobalisdetail.Value))
                                            {
                                                try
                                                {
                                                    User.usr_job = Convert.ToInt32(jobalisdetail.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون کد شغل بیمه  سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }



                                        if (Us != null)
                                        {
                                            db.Entry(User).State = EntityState.Modified;
                                        }
                                        else
                                        {
                                            User.usr_first_payment = true;
                                            User.usr_IsActive = true;
                                            myList.Add(User);
                                        }
                                    }


                                    db.tbUsers.AddRange(myList);

                                    unitofwork.SaveChange();
                                    db.SaveChanges();
                                    transaction.Commit();
                                }
                                else
                                {
                                    return "این شیت فاقد سطر می باشد";
                                }

                                return "true";
                            }
                            else
                            {
                                return "هیچ شیتی در این اکسل وجود ندارد";
                            }
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            //return "ثبت اطلاعات با خطا مواجه شده است. لطفا فایل اکسل ورودی را بررسی کرده و مجددا تلاش کنید";
                            return ex.Message;
                        }
                    }
                }
            }
        }

        public string GetDataFromExcel_Users2(HttpPostedFileBase MyExcelStream)
        {
            // var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));

            using (SaabEntities db = new SaabEntities())
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    using (ExcelEngine xl = new ExcelEngine())
                    {
                        try
                        {
                            IApplication app = xl.Excel;
                            var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                            var Count_Sheets = workbook.Worksheets.Count;
                            if (Count_Sheets > 1)
                            {
                                var sheet = workbook.Worksheets[0];
                                var Rows = sheet.Rows;
                                int Cols = 0;
                                if (Rows.Length >= 1)
                                {
                                    Cols = Rows[0].Cells.Length;

                                    //تعداد ستون ها
                                    if (Count_Sheets != 2 || Cols != 25)
                                    {
                                        return "اکسل وارد شده با اکسل نمونه  متفاوت است";
                                    }

                                    var title = workbook.Worksheets[0].Rows[0].Cells;
                                    var count = workbook.Worksheets[0].Rows.Count();
                                    if (count == 1)
                                    {
                                        return "فایل اکسل فاقد اطلاعات می باشد";
                                    }

                                    Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                    int counter = 0;
                                    foreach (var item in title)
                                    {
                                        dicTitles.Add(item.Value.ToString(), counter);
                                        counter++;
                                    }

                                    // List<string> AddedIWNCode = new List<string>();
                                    List<tbUsers> myList = new List<tbUsers>();
                                    for (int i = 1; i < count; i++)
                                    {
                                        tbUsers Us;
                                        tbUsers User = new tbUsers();
                                        var row = workbook.Worksheets[0].Rows[i];
                                        var cell_code = row.Cells[11];

                                        if (cell_code != null)
                                        {
                                            if (cell_code.Value != null || cell_code.Value != "")
                                            {
                                                User.usr_NationalCode = cell_code.Value;
                                                Us = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == User.usr_NationalCode);
                                                if (Us != null)
                                                {
                                                    User = Us;
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون کد ملی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون کد ملی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_family = row.Cells[0];
                                        if (cell_family != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_family.Value))
                                            {
                                                User.usr_Family = cell_family.Value;
                                            }
                                            else
                                            {
                                                return " لطفا ستون نام خانوادگی سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }
                                        }

                                        var cell_name = row.Cells[1];

                                        if (cell_name != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_name.Value))
                                            {
                                                User.usr_Name = cell_name.Value;
                                            }
                                            else
                                            {
                                                return " لطفا ستون نام سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }
                                        }
                                        else
                                        {
                                            return " لطفا ستون نام سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        }

                                        var cell_Shcode = row.Cells[2];
                                        if (cell_Shcode != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Shcode.Value))
                                            {
                                                try
                                                {
                                                    User.usr_SHCode = Convert.ToInt64(cell_Shcode.Value);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون شماره شناسنامه سطر " + " ( " + i + " ) " + "را درست پر کنید ";
                                                }
                                            }
                                        }


                                        var cell_Address = row.Cells[3];

                                        if (cell_Address != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Address.Value))
                                            {
                                                User.usr_Address = cell_Address.Value.ToString();
                                            }
                                        }

                                        var cell_Email = row.Cells[4];

                                        if (cell_Email != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Email.Value))
                                            {
                                                User.usr_Email = cell_Email.Value.ToString();
                                            }
                                        }

                                        var cell_Fathername = row.Cells[5];

                                        if (cell_Fathername != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Fathername.Value))
                                            {
                                                User.usr_FatherName = cell_Fathername.Value;
                                            }
                                        }

                                        var cell_dutysystem = row.Cells[6];

                                        if (cell_dutysystem != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_dutysystem.Value))
                                            {
                                                var Dutysystem = cell_dutysystem.Value.ToString();
                                                if (Dutysystem == "مشمول")
                                                {
                                                    Dutysystem = 1.ToString();
                                                }
                                                else if (Dutysystem == "پایان خدمت")
                                                {
                                                    Dutysystem = 2.ToString();
                                                }
                                                else if (Dutysystem == "معاف")
                                                {
                                                    Dutysystem = 3.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون نظام وظیفه سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_Dutysystem = Convert.ToInt32(Dutysystem);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون نظام وظیفه سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_Gender = row.Cells[7];

                                        if (cell_Gender != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_Gender.Value))
                                            {
                                                var gender = cell_Gender.Value;
                                                if (gender == "زن")
                                                {
                                                    gender = 1.ToString();
                                                }
                                                else if (gender == "مرد")
                                                {
                                                    gender = 2.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون جنسیت سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_Gender = Convert.ToInt32(gender);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون جنسیت سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_marial = row.Cells[8];

                                        if (cell_marial != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_marial.Value))
                                            {
                                                var maritaIStatus = cell_marial.Value;
                                                if (maritaIStatus == "متاهل")
                                                {
                                                    maritaIStatus = 1.ToString();
                                                }
                                                else if (maritaIStatus == "مجرد")
                                                {
                                                    maritaIStatus = 2.ToString();
                                                }
                                                else
                                                {
                                                    return " لطفا ستون وضعیت تاهل سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }

                                                try
                                                {
                                                    User.usr_MaritaIStatus = Convert.ToInt32(maritaIStatus);
                                                }
                                                catch
                                                {
                                                    return " لطفا ستون وضعیت تاهل سطر " + " ( " + i + " ) " + "را از جدول راهنما پر کنید ";
                                                }
                                            }
                                        }

                                        var cell_phonenumber = row.Cells[9];

                                        if (cell_phonenumber != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_phonenumber.Value))
                                            {
                                                try
                                                {
                                                    User.usr_PhoneNumber = Convert.ToInt64(cell_phonenumber.Value);
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره همراه سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_PlaceOfIssue = row.Cells[10];

                                        if (cell_PlaceOfIssue != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_PlaceOfIssue.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName == cell_PlaceOfIssue.Value.ToString());

                                                if (city != null)
                                                {
                                                    User.usr_PlaceOfIssue = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون محل صدور سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_Pass = row.Cells[12];

                                        if (!string.IsNullOrEmpty(cell_Pass.Value))
                                        {
                                            if (cell_Pass.Value != null)
                                            {
                                                User.usr_Password = FormsAuthentication.HashPasswordForStoringInConfigFile(cell_Pass.Value, "MD5");
                                            }
                                        }



                                        var cell_job = row.Cells[13];

                                        if (cell_job != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_job.Value))

                                            {
                                                var job = rep_business.IfExitsPosition(cell_job.Value);
                                                if (job != 0 || job != null)
                                                {
                                                    if (User.tbUser_link_BusinessSide != null)
                                                    {
                                                        foreach (var item in User.tbUser_link_BusinessSide.ToList())
                                                        {
                                                            item.Status = false;
                                                        }
                                                    }

                                                    tbUser_link_BusinessSide new_link = new tbUser_link_BusinessSide();
                                                    new_link.Status = true;
                                                    new_link.FK_up_ID = job;
                                                    User.tbUser_link_BusinessSide.Add(new_link);
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون سمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_regnum = row.Cells[14];

                                        if (cell_regnum != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_regnum.Value))
                                            {
                                                try
                                                {
                                                    User.usr_RegistrationNumber = Convert.ToInt64(cell_regnum.Value);
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره ثابت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_brDate = row.Cells[15];

                                        if (cell_brDate != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_brDate.Value))
                                            {
                                                var Date = cell_brDate.Value.ToString().Split('/').ToList();
                                                if (Date != null)
                                                {
                                                    try
                                                    {
                                                        PersianCalendar pc = new PersianCalendar();
                                                        int year = int.Parse(Date[0]);
                                                        int month = int.Parse(Date[1]);
                                                        int day = int.Parse(Date[2]);
                                                        var datetime = pc.ToDateTime(year, month, day, 0, 0, 0, 0);
                                                        User.usr_DateOfBrith = datetime;
                                                    }
                                                    catch
                                                    {
                                                        return "  لطفا ستون تاریخ تولد سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                    }
                                                }
                                            }
                                        }

                                        var cell_placeOfdate = row.Cells[16];

                                        if (cell_placeOfdate != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_placeOfdate.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName == cell_placeOfdate.Value);
                                                if (city != null)
                                                {
                                                    User.usr_PlaceOfBrith = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون محل تولد سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_child = row.Cells[17];

                                        if (cell_child != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_child.Value))
                                            {
                                                try
                                                {
                                                    User.usr_Child_Allowance = Convert.ToInt32(cell_child.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون تعداد حق اولاد مشمول بیمه سطر " + " ( " + i + " ) " + "را درست وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_degg = row.Cells[18];

                                        if (cell_degg != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_degg.Value))
                                            {
                                                var degree = cell_degg.Value.ToString();
                                                switch (degree)
                                                {
                                                    case "بیسواد":
                                                        {
                                                            User.usr_Degree = 1;
                                                            break;
                                                        }
                                                    case "سیکل":
                                                        {
                                                            User.usr_Degree = 2;
                                                            break;
                                                        }
                                                    case "دیپلم":
                                                        {
                                                            User.usr_Degree = 3;
                                                            break;
                                                        }
                                                    case "کاردانی":
                                                        {
                                                            User.usr_Degree = 4;
                                                            break;
                                                        }
                                                    case "کارشناسی":
                                                        {
                                                            User.usr_Degree = 5;
                                                            break;
                                                        }
                                                    case "کارشناسی ارشد":
                                                        {
                                                            User.usr_Degree = 6;
                                                            break;
                                                        }
                                                    case "دکتری":
                                                        {
                                                            User.usr_Degree = 7;
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        var cell_gr = row.Cells[19];

                                        if (cell_gr != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_gr.Value))
                                            {
                                                var grade = cell_gr.Value.ToString();
                                                switch (grade)
                                                {
                                                    case "بیسواد":
                                                        {
                                                            User.usr_Grade = 1;
                                                            break;
                                                        }
                                                    case "سیکل":
                                                        {
                                                            User.usr_Grade = 2;
                                                            break;
                                                        }
                                                    case "دیپلم":
                                                        {
                                                            User.usr_Grade = 3;
                                                            break;
                                                        }
                                                    case "کاردانی":
                                                        {
                                                            User.usr_Grade = 4;
                                                            break;
                                                        }
                                                    case "کارشناسی":
                                                        {
                                                            User.usr_Grade = 5;
                                                            break;
                                                        }
                                                    case "کارشناسی ارشد":
                                                        {
                                                            User.usr_Grade = 6;
                                                            break;
                                                        }
                                                    case "دکتری":
                                                        {
                                                            User.usr_Grade = 7;
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        var cell_postal = row.Cells[20];

                                        if (cell_postal != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_postal.Value))
                                            {
                                                try
                                                {
                                                    User.usr_PostalCode = Convert.ToInt64(cell_postal.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون کد پستی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_City_Duty = row.Cells[21];

                                        if (cell_City_Duty != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_City_Duty.Value))
                                            {
                                                var city = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json")).FirstOrDefault(p => p.cityName.Trim() == cell_City_Duty.Value.Trim());
                                                if (city != null)
                                                {
                                                    User.usr_City_Dutysystem = city.cityId;
                                                }
                                                else
                                                {
                                                    return "  لطفا ستون شهر محل خدمت سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_personal_code = row.Cells[22];

                                        if (cell_personal_code != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_personal_code.Value))
                                            {
                                                try
                                                {
                                                    User.usr_Personal_ID = Convert.ToInt32(cell_personal_code.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                            else
                                            {
                                                return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                            }
                                        }
                                        else
                                        {
                                            return "  لطفا ستون شماره پرسنلی سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                        }

                                        var cell_all_child = row.Cells[23];

                                        if (cell_all_child != null)
                                        {
                                            if (!string.IsNullOrEmpty(cell_all_child.Value))
                                            {
                                                try
                                                {
                                                    User.usr_All_Child_Allowance = Convert.ToInt32(cell_all_child.Value.ToString());
                                                }
                                                catch
                                                {
                                                    return "  لطفا ستون تعداد کل اولاد سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                }
                                            }
                                        }

                                        var cell_peymans = row.Cells[24];
                                        if (cell_peymans != null)
                                        {
                                            var pey = cell_peymans.Value;
                                            if (!string.IsNullOrEmpty(pey))
                                            {
                                                if (pey.ToString().Contains(','))
                                                {
                                                    foreach (var item in pey.ToString().Split(',').ToList())
                                                    {
                                                        var peyman = db.tbPeymanContracts?.Where(p => p.pec_Title.Trim() == item && p.Inactive != true).FirstOrDefault();
                                                        if (peyman != null)
                                                        {
                                                            Link_User_And_Peyman pyy = new Link_User_And_Peyman();
                                                            pyy.Status = true;
                                                            pyy.FK_Peyman_ID = peyman.pec_ID;
                                                            User.Link_User_And_Peyman.Add(pyy);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    var peyman = db.tbPeymanContracts.FirstOrDefault(p => p.pec_Title.Trim() == pey.ToString() && p.Inactive != true);
                                                    if (peyman != null)
                                                    {
                                                        Link_User_And_Peyman pyy = new Link_User_And_Peyman();
                                                        pyy.Status = true;
                                                        pyy.FK_Peyman_ID = peyman.pec_ID;
                                                        User.Link_User_And_Peyman.Add(pyy);
                                                    }
                                                    else
                                                    {
                                                        return "  لطفا ستون پیمان سطر " + " ( " + i + " ) " + "را وارد کنید  ";
                                                    }
                                                }
                                            }
                                        }

                                        if (Us != null)
                                        {
                                            db.Entry(User).State = EntityState.Modified;
                                        }
                                        else
                                        {
                                            User.usr_first_payment = true;
                                            User.usr_IsActive = true;
                                            myList.Add(User);
                                        }
                                    }


                                    db.tbUsers.AddRange(myList);

                                    unitofwork.SaveChange();
                                    db.SaveChanges();
                                    transaction.Commit();
                                }
                                else
                                {
                                    return "این شیت فاقد سطر می باشد";
                                }

                                return "true";
                            }
                            else
                            {
                                return "هیچ شیتی در این اکسل وجود ندارد";
                            }
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            //return "ثبت اطلاعات با خطا مواجه شده است. لطفا فایل اکسل ورودی را بررسی کرده و مجددا تلاش کنید";
                            return ex.Message;
                        }
                    }
                }
            }
        }
        [HttpPost]
        [AuthorizeAAA]
        public ActionResult _CreateUser(tbUsers usr, List<int> positionID, string password, List<int> select_peyman, int year_pymn, int month_pymn, int day_pymn, int year, int month, int day, List<string> MoalefeValue, List<int> select_moalefeKarbari)
        {
            try
            {

                var user = db.tbUsers.Where(p => p.usr_NationalCode == usr.usr_NationalCode).FirstOrDefault();


                if (user == null)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        usr.usr_Password = FormsAuthentication.HashPasswordForStoringInConfigFile(password, "MD5");
                    }

                    usr.usr_IsActive = true;
                    usr.usr_first_payment = false;

                    PersianCalendar pc = new PersianCalendar();
                    var datetime = pc.ToDateTime(year, month, day, 0, 0, 0, 0, 0);
                    usr.usr_DateOfBrith = datetime;
                    if (positionID != null)
                    {
                        foreach (var item in positionID)
                        {
                            tbUser_link_BusinessSide link = new tbUser_link_BusinessSide();
                            link.FK_up_ID = item;
                            link.Status = true;
                            usr.tbUser_link_BusinessSide.Add(link);
                        }
                    }
                    //if (user.tbjob == null)
                    //{
                    //    usr.usr_typeshoghl = null;
                    //}
                    //if (user.usr_deyailjob == null)
                    //{
                    //    usr.usr_deyailjob = null;
                    //}
                    if (positionID != null)
                    {
                        foreach (var item in positionID)
                        {




                            tbLink_UserAndAction link = new tbLink_UserAndAction();
                            var ex = db.tbActions.FirstOrDefault();
                            if (ex != null)
                            {
                                link.HasAccess = true;
                                link.FK_Position_ID = item;
                                link.FK_Action_ID = ex.Action_ID;
                                link.FK_User_ID = usr.usr_ID;

                                usr.tbLink_UserAndAction.Add(link);

                            }

                        }
                    }
                    var datetimepymn = pc.ToDateTime(year_pymn, month_pymn, day_pymn, 0, 0, 0, 0, 0);

                    if (select_peyman != null)
                    {
                        foreach (var item in select_peyman)
                        {
                            Link_User_And_Peyman link_moalefe = new Link_User_And_Peyman();
                            link_moalefe.FK_Peyman_ID = item;
                            link_moalefe.datatime = datetimepymn;
                            link_moalefe.Status = true;
                            usr.Link_User_And_Peyman.Add(link_moalefe);
                        }
                    }

                    if (select_moalefeKarbari != null)
                    {
                        for (var i = 0; i < select_moalefeKarbari.Count; i++)
                        {
                            Link_User_And_Moalefe _User_And_Moalefe = new Link_User_And_Moalefe();
                            _User_And_Moalefe.FK_Moalefe_ID = select_moalefeKarbari[i];
                            _User_And_Moalefe.Moalefe_Value = MoalefeValue[i];
                            _User_And_Moalefe.Status = true;
                            usr.Link_User_And_Moalefe.Add(_User_And_Moalefe);
                        }
                    }
                    var res = rep_users.Create(usr);
                    TempData["response"] = res;

                    return RedirectToAction("Index", "Users");
                }
                else
                {
                    TempData["response"] = "false_NotFoundUser";
                    return RedirectToAction("Index", "Users");
                }
            }
            catch
            {
                TempData["response"] = "catch";
                return RedirectToAction("Index", "Users");
            }
        }
        public ActionResult listussel()
        {
            var t = db.tbUsers.ToList();
            return PartialView(t);

        }

        //[AuthorizeAAA]
        [HttpPost]
        public ActionResult _EditUser(tbUsers usr, List<int> positionID, string password, List<int> select_peyman, int year_pymn_edit, int month_pymn_edit, int day_pymn_edit, int year_edit, int month_edit, int day_edit, List<string> MoalefeValue, List<int> select_moalefeKarbari)
        {
            try
            {
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == usr.usr_NationalCode);
                if (user != null)
                {
                    if (user.usr_ID != usr.usr_ID)
                    {
                        TempData["response"] = "false_NotFoundUser";
                        return RedirectToAction("Index", "Users");
                    }
                }



                var us = rep_users.Find(usr.usr_ID);
                if (!string.IsNullOrEmpty(password))
                {
                    us.usr_Password = FormsAuthentication.HashPasswordForStoringInConfigFile(password, "MD5");
                }
                //var t = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == us.usr_ID && p.Status == true).FirstOrDefault();
                //if (t != null)
                //{
                //    us.Link_User_And_Peyman = t.tbPeymanContracts.pec_Title;
                //}
                PersianCalendar pc = new PersianCalendar();
                var datetime = pc.ToDateTime(year_edit, month_edit, day_edit, 0, 0, 0, 0, 0);
                us.usr_DateOfBrith = datetime;
                us.usr_Name = usr.usr_Name;
                us.usr_Family = usr.usr_Family;
                us.usr_NationalCode = usr.usr_NationalCode;
                us.usr_PostalCode = usr.usr_PostalCode;
                us.usr_PlaceOfBrith = usr.usr_PlaceOfBrith;
                us.usr_MaritaIStatus = usr.usr_MaritaIStatus;
                us.usr_SHCode = usr.usr_SHCode;
                us.usr_PlaceOfIssue = usr.usr_PlaceOfIssue;
                us.usr_FatherName = usr.usr_FatherName;
                us.usr_PhoneNumber = usr.usr_PhoneNumber;
                us.usr_Gender = usr.usr_Gender;
                us.usr_Dutysystem = usr.usr_Dutysystem;
                us.usr_RegistrationNumber = usr.usr_RegistrationNumber;
                us.usr_Address = usr.usr_Address;
                us.usr_All_Child_Allowance = usr.usr_All_Child_Allowance;
                us.usr_City_Dutysystem = usr.usr_City_Dutysystem;
                us.usr_Personal_ID = usr.usr_Personal_ID;
                us.usr_Child_Allowance = usr.usr_Child_Allowance;
                us.usr_Grade = usr.usr_Grade;
                us.usr_Degree = usr.usr_Degree;
                us.usr_Email = usr.usr_Email;
                us.usr_Picture = usr.usr_Picture;
                us.usr_IsActive = usr.usr_IsActive;
                us.usr_first_payment = us.usr_first_payment;
                us.usr_typeshoghl = usr.usr_typeshoghl;

                foreach (var item in us.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
                {
                    item.Status = false;
                }

                if (positionID != null)
                {
                    foreach (var item in positionID)
                    {
                        tbUser_link_BusinessSide l = new tbUser_link_BusinessSide();
                        l.Status = true;
                        l.FK_up_ID = item;
                        tbLink_UserAndAction l2 = new tbLink_UserAndAction();
                        var ex = db.tbActions.FirstOrDefault();
                        if (ex != null)
                        {
                            l2.HasAccess = true;
                            l2.FK_Position_ID = item;
                            l2.FK_Action_ID = ex.Action_ID;
                            us.tbLink_UserAndAction.Add(l2);

                        }

                        us.tbUser_link_BusinessSide.Add(l);
                    }
                }

                foreach (var item in us.Link_User_And_Peyman.Where(p => p.Status == true).ToList())
                {
                    item.Status = false;
                }
                var datetime_pymn = pc.ToDateTime(year_pymn_edit, month_pymn_edit, day_pymn_edit, 0, 0, 0, 0, 0);

                if (select_peyman != null)
                {
                    foreach (var item in select_peyman)
                    {
                        Link_User_And_Peyman l = new Link_User_And_Peyman();
                        l.Status = true;
                        l.datatime = datetime_pymn;
                        l.FK_Peyman_ID = item;
                        us.Link_User_And_Peyman.Add(l);
                    }
                }

                foreach (var item in us.Link_User_And_Moalefe.Where(p => p.Status == true).ToList())
                {
                    item.Status = false;
                }
                if (select_moalefeKarbari != null)
                {
                    for (var i = 0; i < select_moalefeKarbari.Count; i++)
                    {
                        Link_User_And_Moalefe _User_And_Moalefe = new Link_User_And_Moalefe();
                        _User_And_Moalefe.FK_Moalefe_ID = select_moalefeKarbari[i];
                        _User_And_Moalefe.Moalefe_Value = MoalefeValue[i];
                        _User_And_Moalefe.Status = true;
                        us.Link_User_And_Moalefe.Add(_User_And_Moalefe);
                    }
                }

                var res = rep_users.SaveChanges();
                TempData["response"] = res.ToString();
                return RedirectToAction("Index", "Users");
            }
            catch
            {
                TempData["response"] = "catch";
                return RedirectToAction("Index", "Users");
            }
        }

        [AuthorizeAAA]
        public ActionResult _Show_UserProfile()
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







        [AuthorizeAAA]
        public ActionResult _Show_UserProfile2()
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

        [AuthorizeAAA]
        public ActionResult showinformation()
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

        public JsonResult showinformation2340()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var CodeMeli = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode.ToString() == CodeMeli);
                    if (User != null)
                    {
                        return Json(new { UserId = User.usr_ID }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            return Json(new { UserId = 0 }, JsonRequestBehavior.AllowGet);
        }


        public int showinformation234()
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
                        return User.usr_ID;
                    }
                    return 0;

                }
            }
            return 0;

        }



        [AuthorizeAAA]
        public ActionResult _Show_UserProfile_bigger()
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

        [AuthorizeAAA]
        public ActionResult EditPicProfile(HttpPostedFileBase usr_Picture)
        {
            try
            {
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    var nationalcode = Base64.Base64Decode(cookie_user.Value);
                    var User = db.tbUsers.Where(p => p.usr_NationalCode == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        var Pic = usr_Picture.FileName;
                        var path = Server.MapPath("~/Assets/img/user/" + Pic);
                        if (!System.IO.File.Exists(path))
                        {
                            User.usr_Picture = Pic;
                            usr_Picture.SaveAs(path);
                            db.SaveChanges();
                            return RedirectToAction("Index", "Users");
                        }
                        else
                        {
                            User.usr_Picture = Pic;
                            db.SaveChanges();
                        }
                    }
                }

                return RedirectToAction("Index", "Users");
            }
            catch
            {
                return RedirectToAction("Index", "Users");
            }
        }

        /// <summary>
        /// فعال و غیرفعال کردن کاربر
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [AuthorizeAAA]
        public bool DeActiveUser(int id)
        {
            var us = rep_users.Disable(id);
            return us;
        }


        [AuthorizeAAA]
        public ActionResult ChangeActionsParent(int id)
        {
            return PartialView("_ActionAccessLevel", rep_users.GetActionsForController(id));
        }

        [AuthorizeAAA]
        public ActionResult AddNewLevelForCreateUser(int Counter)
        {
            return PartialView("~/Areas/Users/Views/Users/_LevelMoalefeForUser.cshtml", Counter);
        }


        #endregion
    }
}