using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Globalization;
using Microsoft.Ajax.Utilities;
using SaabWebProject.Models.ViewModels.Reporter.shenaght;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.IO;
using System.Data.Entity;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Stimulsoft.Report.StiRecentConnections;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using System.Data.Entity;
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
using DocumentFormat.OpenXml.ExtendedProperties;
using SkiaSharp;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using static SkiaSharp.HarfBuzz.SKShaper;
using static SaabWebProject.Models.ViewModels.Contracts.Function.FunctionModel;

namespace SaabWebProject.Areas.Reporter.Controllers
{
    public class DashboardController : Controller
    {
        // GET: Reporter/Dashboard
        SaabEntities db;
        shenaght shenaght = new shenaght();
        shenaght shenaght2 = new shenaght();

        tblinkUserAndPeymanRepository linkuserPeymanRepo;
        tbUsersRepository UserRepo;
        public DashboardController()
        {
            db = new SaabEntities();
            linkuserPeymanRepo = new tblinkUserAndPeymanRepository(db);
            UserRepo = new tbUsersRepository(db);
        }
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View("~/Areas/Reporter/Views/Dashboard/Index.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult ManageNews()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/ManageNews.cshtml");

            DateTime date1 = new DateTime(2023, 4, 1);
            DateTime date2 = new DateTime(2023, 4, 11);

            TimeSpan duration = date2 - date1;

            Console.WriteLine("Duration: " + duration.TotalDays + " days");
        }
        //public ActionResult Dashboard()
        //{
        //    return PartialView("~/Areas/Reporter/Views/Dashboard/Dashboard.cshtml");
        //}

        [AuthorizeAAA]
        public ActionResult KPI_dashboard()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/KPI_dashboard.cshtml");
        }




    //    public ActionResult Linkafradmoalfeh()
    //    {
    //        using (var db = new SaabEntities())
    //        {
    //            var model = new FunctionModel
    //            {
    //                LinkafradmoalfehData = new FunctionModel.LinkafradmoalfehViewModel
    //                {
    //                    MahdodayList = db.dbtarifmahdodayt1
    //                        .Include(x => x.tbContractMoalefeDastmozdi)
    //                        .Select(x => new FunctionModel.MahdodayViewModel
    //                        {
    //                            ID = x.ID,
    //                            FK_moalfe = x.FK_moalfe ??0,
    //                            Title = x.tbContractMoalefeDastmozdi.md_Title
    //                        }).ToList(),

    //                    TickList = db.dbtarifmahdodayt6
    //                        .Select(x => new FunctionModel.TickMoadelViewModel
    //                        {
    //                            FK_name = x.FK_name ??0,
    //                            FK_moalfeh = x.FK_moalfeh ??0,
    //                            value = x.value ??false
    //                        }).ToList(),

    //                    ContractUsers = db.tbmarahesabt4
    //.Where(s => s.control == true)
    //.GroupBy(s => s.FK_usr)
    //.Select(g => g.FirstOrDefault())
    //.AsEnumerable() // ← بعد از این، تبدیل به LINQ to Objects میشه
    //.Select(c => new FunctionModel.ContractUserViewModel
    //{
    //    FK_usr = c.FK_usr ?? 0,
    //    FullName = c.tbUsers.FullName, // الان مشکلی نیست
    //    PersonalID = (int)c.tbUsers.usr_Personal_ID
    //}).ToList()

    //                }
    //            };

    //            return PartialView("~/Areas/Reporter/Views/Dashboard/Linkafradmoalfeh.cshtml", model);
    //        }
    //    }



        [AuthorizeAAA]

        public ActionResult Linkafradmoalfeh()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/Linkafradmoalfeh.cshtml");
        }
        public ActionResult mahdodayatlinkha()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/mahdodayatlinkha.cshtml");
        }
        public ActionResult listtarifmoalf()
        {
            return View();
        }
        public ActionResult listtarifmoalfedit(int id=0)
        {
            var find = db.dbtarifmahdodayt1.Where(p => p.ID == id).FirstOrDefault();
            if(find != null)
            {
                return View(find);

            }
            return View();

        }
        public ActionResult listtarifmoalfeditcreat(int id = 0,int fk=0)
        {
            var find = db.dbtarifmahdodayt1.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                find.FK_moalfe = fk;
                db.SaveChanges();
                return Content("True");

            }
            return Content("True");

        }
        public ActionResult mahdodayatha()
        {
            return View();
        }
        public ActionResult mahdodayatha_edit()
        {
            return View();
        }

        public string UpdateRecordmahdodat3(List<dbtarifmahdodayt4> modelList)
        {
            try
            {
                List<dbtarifmahdodayt4> dbtarifmahdodayt4 = new List<dbtarifmahdodayt4>();
                using (var context = new SaabEntities())
                {
                    var find2 = context.dbtarifmahdodayt4
                            .Where(p =>
                                p.FK_usr !=null &&
                                p.FK_name  !=null).ToList();
                    foreach (var it in modelList)
                    {
                        var find = find2
                            .FirstOrDefault(p =>
                                p.FK_usr == it.FK_usr &&
                                p.FK_name == it.FK_name);

                        if (find != null)
                        {
                            // Update
                            find.Position = it.Position;
                        }
                        else
                        {
                            // Insert
                            dbtarifmahdodayt4.Add(new dbtarifmahdodayt4
                            {
                                FK_usr = it.FK_usr,
                                FK_name = it.FK_name,
                                Position = it.Position
                            });
                        }
                    }
                    if (dbtarifmahdodayt4.Count > 0)
                    {
                        context.dbtarifmahdodayt4.AddRange(dbtarifmahdodayt4);

                    }
                    // ⭐ فقط یکبار ذخیره — ۱۰ برابر سریع‌تر
                    context.SaveChanges();
                }

                return "true";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public ActionResult marhal4mahdod()
        {
            using (var db = new SaabEntities())
            {
                db.Configuration.LazyLoadingEnabled = false;

                var model = new Marhale4ViewModel
                {
                    Users = db.tbUsers.AsNoTracking().ToList(),

                    Cities = db.tbCities.AsNoTracking().ToList(),

                    UserPeymans = db.Link_User_And_Peyman
                        .Include("tbPeymanContracts")
                        .Where(x => x.Status == true)
                        .AsNoTracking()
                        .ToList(),

                    Marhale3 = db.dbtarifmahdodayt3
                        .Where(p => p.name != null)
                        .AsNoTracking()
                        .ToList(),

                    Values = db.dbtarifmahdodayt4
                        .AsNoTracking()
                        .ToList()
                };

                return View(model);
            }
        }


        public ActionResult creatmarhaleh1(List<int> Fkpym = null)
        {
            List<dbtarifmahdodayt1> dbtarifmahdodayt1 = new List<dbtarifmahdodayt1>();
            foreach (var it in Fkpym)
            {
                dbtarifmahdodayt1 dbrelatinperson = new dbtarifmahdodayt1();
                dbrelatinperson.FK_moalfe = it;
                dbtarifmahdodayt1.Add(dbrelatinperson);
            }
            db.dbtarifmahdodayt1.AddRange(dbtarifmahdodayt1);
            db.SaveChanges();
            return Content("True");

        }
        //public ActionResult viewmarheh2mahdodayat()
        //{
        //    return View();  
        //}
        //public async Task< ActionResult> viewmarheh2mahdodayat2()
        //{
        //    var dbtarifmahdodayt1 = await db.dbtarifmahdodayt1.ToListAsync();
        //    var tbUsers = await db.tbUsers.ToListAsync();

        //    return View();
        //}
        public ActionResult viewmarheh2mahdodayat2()
        {
            // گرفتن لیست موالفه‌ها و کاربران در یک بار کوئری
            var moalefeList = db.dbtarifmahdodayt1
                                .Select(m => new
                                {
                                    m.ID,
                                    Title = m.tbContractMoalefeDastmozdi.md_Title
                                }).ToList();

            var users = db.tbUsers.ToList();

            // دیکشنری شهرها برای دسترسی سریع
            var cityDict = db.tbCities.ToDictionary(c => c.ID, c => c.Name);

            // گرفتن پیمان‌ها با شرط Status == true و ساخت دیکشنری برای جستجوی سریع
            var peymanDict = db.Link_User_And_Peyman
                                .Where(p => (p.Status ?? false) == true)  // اگر Status از نوع bool? هست
                                .Select(p => new { p.FK_User_ID, Title = p.tbPeymanContracts.pec_Title })
                                .ToList()
                                .GroupBy(p => p.FK_User_ID)
                                .ToDictionary(g => g.Key, g => g.First().Title);

            // گرفتن همه مقدارهای موالفه به صورت یکجا
            var userMoalefeValues = db.dbtarifmahdodayt2
                                     .ToList()
                                     .GroupBy(v => v.FK_usr)
                                     .ToDictionary(g => g.Key, g => g.ToList());

            var model = new List<MarhaleMahdodayatViewModel>();

            foreach (var user in users)
            {
                string cityName = "وصل نیست";
                if (user.usr_City_Dutysystem.HasValue && cityDict.ContainsKey(user.usr_City_Dutysystem.Value))
                    cityName = cityDict[user.usr_City_Dutysystem.Value];

                string peymanTitle = "وصل نیست";
                if (peymanDict.ContainsKey(user.usr_ID))
                    peymanTitle = peymanDict[user.usr_ID];

                var moalefeValues = new List<MoalefeValueViewModel>();

                List<dbtarifmahdodayt2> userValues = null;
                userMoalefeValues.TryGetValue(user.usr_ID, out userValues);

                foreach (var moalefe in moalefeList)
                {
                    var valueEntity = userValues?.FirstOrDefault(v => v.FK_moalfe == moalefe.ID);

                    moalefeValues.Add(new MoalefeValueViewModel
                    {
                        MoalefeId = moalefe.ID,
                        MoalefeTitle = moalefe.Title,
                        Value = valueEntity != null ? (long)valueEntity.value : 0
                    });
                }

                model.Add(new MarhaleMahdodayatViewModel
                {
                    FullName = user.FullName,
                    PeymanTitle = peymanTitle,
                    CityName = cityName,
                    MoalefeValues = moalefeValues
                });
            }

            return View(model);
        }
        public string UpdateRecordmahdodat2(List<dbtarifmahdodayt2> modelList)
        {
            try
            {
                using (var context = new SaabEntities())
                {
                    var userIds = modelList.Select(x => x.FK_usr).Distinct().ToList();

                    var existingRecords = context.dbtarifmahdodayt2
                        .Where(x => userIds.Contains(x.FK_usr))
                        .ToList();

                    var newRecords = new List<dbtarifmahdodayt2>();

                    foreach (var it in modelList)
                    {
                        var find = existingRecords
                            .FirstOrDefault(p => p.FK_usr == it.FK_usr && p.FK_moalfe == it.FK_moalfe);

                        if (find != null)
                        {
                            find.value = it.value;
                        }
                        else
                        {
                            newRecords.Add(it);
                        }
                    }

                    if (newRecords.Count > 0)
                    {
                        context.dbtarifmahdodayt2.AddRange(newRecords);
                    }

                    context.SaveChanges();

                    return "true";
                }
            }
            catch (Exception ex)
            {
                // لاگ خطا (در صورت وجود logger)
                // Logger.LogError(ex, "خطا در UpdateRecordmahdodat2");
                return $"خطا: {ex.Message}";
            }
        }


        //public string UpdateRecordmahdodat2(List<dbtarifmahdodayt2> modelList)
        //{
        //    using (var context = new SaabEntities())
        //    {
        //        var existingRecords = context.dbtarifmahdodayt2

        //            .ToList();

        //        var newRecords = new List<dbtarifmahdodayt2>();

        //        foreach (var it in modelList)
        //        {
        //            var find = existingRecords.FirstOrDefault(p => p.FK_usr == it.FK_usr && p.FK_moalfe == it.FK_moalfe);

        //            if (find != null)
        //            {
        //                find.value = it.value; // به‌روزرسانی مقدار موجود
        //            }
        //            else
        //            {
        //                // بررسی اینکه FK_usr در جدول مرجع tbUsers وجود دارد
        //                bool userExists = context.tbUsers.Any(u => u.usr_ID == it.FK_usr);
        //                if (userExists)
        //                {
        //                    newRecords.Add(it);
        //                }
        //                //else
        //                //{
        //                //    return $"خطا: کاربر با FK_usr = {it.FK_usr} در جدول tbUsers وجود ندارد!";
        //                //}
        //            }
        //        }

        //        context.SaveChanges(); // ذخیره‌سازی تغییرات در رکوردهای موجود

        //        if (newRecords.Count > 0)
        //        {
        //            context.dbtarifmahdodayt2.AddRange(newRecords);
        //            context.SaveChanges(); // ذخیره‌سازی گروهی برای رکوردهای جدید
        //        }

        //        return "true";
        //    }
        //}

        public ActionResult marhal3mahdod()
        {
            return View();
        }
        public ActionResult marhal3mahdodedit(int id=0)
        {
            var find = db.dbtarifmahdodayt3.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                return View(find);

            }
            return View();
        }
        public ActionResult edittarifmahsol(int id = 0)
        {
            var find = db.dbAsnadcratname.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                return View(find);

            }
            return View();
        }
        public ActionResult marhal3mahdodeditcreat(int id = 0,string name="")
        {
            var find = db.dbtarifmahdodayt3.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                find.name = name;
                db.SaveChanges();
                return Content("True");

            }
            return Content("True");
        }
        public string UpdateRecordmarhel3(List<dbtarifmahdodayt3> modelList)
        {
            List<dbtarifmahdodayt3> tbinformation2 = new List<dbtarifmahdodayt3>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    var find = context.dbtarifmahdodayt3.Where(p => p.name == model.name).FirstOrDefault();
                    if (find != null)
                    {
                        find.name = model.name;

                        context.SaveChanges();
                    }
                    else
                    {
                        dbtarifmahdodayt3 tbinformation = new dbtarifmahdodayt3();
                        tbinformation.name = model.name;

                        tbinformation2.Add(tbinformation);
                    }




                }
                context.dbtarifmahdodayt3.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }

        //public ActionResult marhal4mahdod()
        //{
        //    return View();
        //}
        //public string UpdateRecordmahdodat3(List<dbtarifmahdodayt4> modelList)
        //{
        //    List<dbtarifmahdodayt4> tbinformation2 = new List<dbtarifmahdodayt4>();

        //    using (var context = new SaabEntities())
        //    {
        //        foreach (var it in modelList)
        //        {
        //            var find = context.dbtarifmahdodayt4.Where(p => p.FK_usr == it.FK_usr && p.FK_name == it.FK_name).FirstOrDefault();
        //            if (find != null)
        //            {
        //                find.Position = it.Position;
        //                context.SaveChanges();

        //            }
        //            else
        //            {
        //                context.dbtarifmahdodayt4.Add(it);
        //                context.SaveChanges();
        //            }
        //        }

        //        return "true";
        //    }
        //}
        //public ActionResult marhal5mahdod()
        //{
        //    return View();
        //}
        public ActionResult viewmarheh2mahdodayat()
        {
            try
            {
                using (var db = new SaabEntities())
                {
                    db.Configuration.LazyLoadingEnabled = false;

                    // 1) مؤلفه‌ها
                    var moalefeh = db.dbtarifmahdodayt1
                        .Select(x => new MoalefeVM
                        {
                            Id = x.ID,
                            Title = x.tbContractMoalefeDastmozdi.md_Title
                        }).ToList();

                    // 2) کاربران - بدون FullName
                    var usersRaw = db.tbUsers
                        .Select(u => new
                        {
                            u.usr_ID,
                            u.usr_Name,
                            u.usr_Family,
                            City = db.tbCities
                                .Where(c => c.ID == u.usr_City_Dutysystem)
                                .Select(c => c.Name)
                                .FirstOrDefault(),
                            Peyman = db.Link_User_And_Peyman
                                .Where(p => p.FK_User_ID == u.usr_ID && p.Status == true)
                                .Select(p => p.tbPeymanContracts.pec_Title)
                                .FirstOrDefault()
                        }).ToList();

                    // 2b) تبدیل به ViewModel همراه با FullName
                    var users = usersRaw
                        .Select(x => new UserRowVM
                        {
                            UserId = x.usr_ID,
                            FullName = (x.usr_Name ?? "") + " " + (x.usr_Family ?? ""),
                            City = x.City,
                            Peyman = x.Peyman
                        }).ToList();

                    // 3) مقادیر
                    var values = db.dbtarifmahdodayt2
                        .Select(v => new ValueVM
                        {
                            FK_usr = v.FK_usr ?? 0,
                            FK_moalfe = v.FK_moalfe ?? 0,
                            Value = v.value
                        }).ToList();

                    // 4) ViewModel نهایی
                    var vm = new Marhale2ViewModel
                    {
                        Users = users,
                        Moalefeh = moalefeh,
                        Values = values
                    };

                    return View(vm);
                }
            }
            catch
            {
                return View(new Marhale2ViewModel
                {
                    Users = new List<UserRowVM>(),
                    Moalefeh = new List<MoalefeVM>(),
                    Values = new List<ValueVM>()
                });
            }
        }






        public ActionResult marhal5mahdod2()
        {
            return View();
        }
        public ActionResult marhal5mahdod()
        {
            using (var db = new SaabEntities())
            {
                db.Configuration.LazyLoadingEnabled = false;

                var peymans = db.tbPeymanContracts
                                .AsNoTracking()
                                .ToList();

                var cities = db.tbpeymancities
                               .Include("tbCities")
                               .AsNoTracking()
                               .ToList();

                var marhale3 = db.dbtarifmahdodayt3
                                 .Where(p => p.name != null)
                                 .AsNoTracking()
                                 .ToList();

                var moalefs = db.dbtarifmahdodayt1
                                .Include("tbContractMoalefeDastmozdi")
                                .AsNoTracking()
                                .ToList();

                var values = db.dbtarifmahdodayt5
                               .AsNoTracking()
                               .ToList();

                var model = new Marhale5ViewModel
                {
                    Peymans = peymans,
                    Cities = cities,
                    Marhale3 = marhale3,
                    Moalefs = moalefs,
                    Values = values
                };

                return View(model);
            }
        }





        //public string UpdateRecordmahdodat5(List<dbtarifmahdodayt5> modelList)
        //{
        //    List<dbtarifmahdodayt5> tbinformation2 = new List<dbtarifmahdodayt5>();

        //    using (var context = new SaabEntities())
        //    {
        //        foreach (var it in modelList)
        //        {
        //            var find = context.dbtarifmahdodayt5.Where(p => p.FK_Pymn == it.FK_Pymn && p.FK_marhal4 == it.FK_marhal4&&p.FK_city==it.FK_city&&p.FK_moalf==it.FK_moalf).FirstOrDefault();
        //            if (find != null)
        //            {
        //                find.value = it.value;
        //                context.SaveChanges();

        //            }
        //            else
        //            {
        //                context.dbtarifmahdodayt5.Add(it);
        //                context.SaveChanges();
        //            }
        //        }

        //        return "true";
        //    }
        //}
        public string UpdateRecordmahdodat5(List<dbtarifmahdodayt5> modelList)
        {
            using (var context = new SaabEntities())
            {
                List<dbtarifmahdodayt5> dbtarifmahdodayt5 = new List<dbtarifmahdodayt5>();

                // گرفتن داده‌های موجود فقط یکبار
                var existing = context.dbtarifmahdodayt5.ToList();

                foreach (var it in modelList)
                {
                    var find = existing.FirstOrDefault(p =>
                        p.FK_Pymn == it.FK_Pymn &&
                        p.FK_marhal4 == it.FK_marhal4 &&
                        p.FK_city == it.FK_city &&
                        p.FK_moalf == it.FK_moalf
                    );

                    if (find != null)
                    {
                        find.value = it.value;
                    }
                    else
                    {
                        dbtarifmahdodayt5.Add(new dbtarifmahdodayt5
                        {
                            FK_Pymn = it.FK_Pymn,
                            FK_marhal4 = it.FK_marhal4,
                            FK_city = it.FK_city,
                            FK_moalf = it.FK_moalf,
                            value = it.value
                        });
                    }
                }
                context.dbtarifmahdodayt5.AddRange(dbtarifmahdodayt5);
                context.SaveChanges();
                return "true";
            }
        }

        public ActionResult Result_dashboard()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/Result_dashboard.cshtml");
        }
        //=================================================================================[BannerOption]Start

        [AuthorizeAAA]
        public ActionResult OptionBanner()
        {

            return View("~/Areas/Reporter/Views/Dashboard/OptionBanner.cshtml");

        }

        [AuthorizeAAA]
        public ActionResult OptionBanner_List()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/OptionBanner_List.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult OptionBanner_Create()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/OptionBanner_Create.cshtml");

        }
        public string UploadAttachmentFile(string ImgDir = "", DateTime StartDate = default, DateTime EndDate = default, string TextTop = "", string TextCenter = ""
           , string TextDown = "", string ColorTop = "", string ColorCenter = "", string ColorDown = "",bool IsActive=true,
    IEnumerable<HttpPostedFileBase> files = null)
        {


            tbBannerSlider obj = new tbBannerSlider();

            obj.StartDate = StartDate;
            obj.EndDate = EndDate;
            obj.TextTop = TextTop;
            obj.TextCenter = TextCenter;
            obj.TextDown = TextDown;
            obj.ColorTop = ColorTop;
            obj.ColorCenter = ColorCenter;
            obj.ColorDown = ColorDown;
            obj.IsActive = IsActive;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Reporter/Content/BannerImages/" + filename));
                        obj.ImgDir = filename;
                        db.tbBannerSlider.Add(obj);
                        db.SaveChanges();
                        return "True";

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {
                obj.ImgDir = ImgDir;
                db.tbBannerSlider.Add(obj);
                db.SaveChanges();
                return "True";
            }
        }

        public ActionResult GetId_Banner(int id)
        {
            var result = db.tbBannerSlider.Find(id);
            //-------------------------------------
            if (result.ID == null)
            {
                result.ID = 0;
            }
            return View("~/Areas/Reporter/Views/Dashboard/OptionBanner_Update.cshtml", result);
        }

        public string EditeUploadAttachmentFile(string ImgDir = "", DateTime StartDate = default, DateTime EndDate = default, string TextTop = "", string TextCenter = ""
          , string TextDown = "", string ColorTop = "", string ColorCenter = "", string ColorDown = "", bool IsActive = true, int ID = 0)
        {


            tbBannerSlider obj = db.tbBannerSlider.Find(ID);



            obj.StartDate = StartDate;
            obj.EndDate = EndDate;
            obj.TextTop = TextTop;
            obj.TextCenter = TextCenter;
            obj.TextDown = TextDown;
            obj.ColorTop = ColorTop;
            obj.ColorCenter = ColorCenter;
            obj.ColorDown = ColorDown;
            obj.IsActive = IsActive;

            var old = db.tbBannerSlider.Find(ID);

            //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;

            if (ColorTop == null)
            {
                db.Entry(old).CurrentValues.SetValues(obj);
                db.SaveChanges().ToString(); ;
                return "True";
            }
            else
            {
                obj.ImgDir = ImgDir;
                db.Entry(old).CurrentValues.SetValues(obj);
                db.SaveChanges().ToString();
                return "True";
            }
        }
        public ActionResult DeletBannerList(int ID)
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
                db.tbBannerSlider.Remove(db.tbBannerSlider.Find(ID));
                return System.Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }

        }
        public bool DeActiveBanner(int id)
        {
            try
            {
                var Banner = db.tbBannerSlider.Find(id);
                if ((bool)Banner.IsActive)
                {
                    Banner.IsActive = false;
                }
                else
                {
                    Banner.IsActive = true;
                }

                return System.Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }
        //---------------------------------------------------------------------------------
        public void DeactivateBannersForNewYear()
        {

                foreach (var banner in db.tbBannerSlider)
                {
                    banner.IsActive = false;
                }
                db.SaveChanges();
        }
        //=================================================================================[BannerOption]End
        //=================================================================================[]
        public ActionResult GozareshHoghogh_Dashboard()
        {
            return View();
        }
        public ActionResult GozareshHoghogh_Dashboardd()
        {
            return View();
        }
        //=================================================================================[]End
        [AuthorizeAAA]
        public ActionResult A1_Shenakht()
        {
            return View();
        }

        public ActionResult A1_Shenakht_Mod01(int id = 0)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();
            return View(find);
        }
        public ActionResult A1_Shenakht_Mod01_M01()
        {
            return View();
        }




        public ActionResult ModiriyatFaaliat_faliatsabet()
        {
            return View();
        }
        public ActionResult ModiriyatFaaliat_day()
        {
            return View();

        }
        public ActionResult viewappendone(int id)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }

        public ActionResult viewappendshghes(int id)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }

        public ActionResult A1_Shenakht_Mod02(int id = 0)
        {

            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }

        public ActionResult A1_Shenakht_Mod03(int id = 0)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }
        public ActionResult A1_Shenakht_Mod04( int id=0)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }
        public ActionResult A1_Shenakht_Mod05(int id = 0)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();

            return View(find);
        }
        //public ActionResult A1_Shenakht_Mod06()
        //{
        //    var d = db.tbresultpishkansehnaght.Where(p => p.FK_NAME == 1).FirstOrDefault();
        //    return View(d);
        //}
        public ActionResult Listname()
        {
            var y = db.tbNamshnaght.ToList();
            return View(y);        }
        public ActionResult viewsabt()
        {
            return View();
        }
        public ActionResult viewsabtname(int id=0)
        {
            var find = db.tbNamshnaght.Where(p => p.ID == id).FirstOrDefault();
            return View(find);
        }
        //public ActionResult A1_Shenakht_Mod04()
        //{
        //    return View();
        //}
        //public ActionResult A1_Shenakht_Mod05()
        //{
        //    return View();
        //}
        //public ActionResult A1_Shenakht_Mod06()
        //{
        //    return View();
        //}



        public ActionResult A1_Shenakht_Mod06(int id = 0)
        {
            //var OutPutFile = Magic_box_calculations2(id);
            //string extension = ".xlsx";

            var stream = new MemoryStream();
            //OutPutFile.Save(stream, extension);
            //var workbook = OutPutFile;
            //var mimeType = MimeTypes.ByExtension[extension];

            
            var d = db.tbresultpishkansehnaght.Where(p => p.FK_NAME == id).FirstOrDefault();
            return View(d);


        }
        //=================================================================================[ModiriyatFaaliat]Start
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_AfradeZirmajmooe()
        {
            return View(db.tbperson.ToList());
        }
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_ListeFaaliyat()
        {
            List<dbRelatinafrad> dbRelatinafrad = new List<dbRelatinafrad>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;

                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                var find = db.dbRelatinafrad
                             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                             .ToList();
                var finddd = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == personal).ToList();
                foreach(var it in finddd)
                {
                    if (it.FK_usr !=personal)
                    {
                        var ft = db.dbRelatinafrad
                          .Where(p => p.dbrelatinperson.FK_usr == it.FK_usr && p.FK_tbcreatfaal != null)
                          // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                          .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                          .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                          .ToList();
                        find.AddRange(ft);
                    }
                 
                }
                dbRelatinafrad.AddRange(find);
            }

            return View(dbRelatinafrad);
        }

        //public ActionResult pyam(int id)
        //{

        //}

        public ActionResult ModiriyatFaaliat_TarifeTatilatedit(int id=0)
        {
            var fin = db.tbcraettatel.Where(p => p.ID == id).FirstOrDefault();
            return View(fin);

        }
        public ActionResult creatjavab(int id = 0)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
            }

            // جستجوی داده‌ها و مرتب‌سازی بر اساس ستون notsee
            var find = db.tbsherkatpyam
                .Where(p => p.FK_usr1 == personal && p.FK_isr2 == id)
                .OrderByDescending(p => p.notseee == 1) // رکوردهایی که notsee == 1 در ابتدا
                .ToList();

            return View(find);
        }

        public ActionResult ModiriyatFaaliat_ListeFaaliyat_ErjaModal(int id = 0)
        {
            var fin = db.dbRelatinafrad.Where(p => p.ID == id).FirstOrDefault();
            return View(fin);
        }
        public ActionResult ModiriyatFaaliat_ListeFaaliyat_PayamModal(int id = 0)
        {
            List<tbsherkatpyam> tbsherkatpyam = new List<tbsherkatpyam>();

            if (id != 0)
            {

            }
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
            }
            var fioln = db.tbsherkatpyam.Where(p => p.notseee == 1 && p.Fk_afrad == id&&p.FK_usr1==personal).ToList();
            foreach(var it in fioln)
            {
                it.notseee = 0;
                db.SaveChanges();

            }
            var fioln2 = db.tbsherkatpyam.Where(p =>  p.Fk_afrad == id ).ToList();
            if (fioln2.Count == 0)
            {
                tbsherkatpyam tbsherkatpyam2 = new tbsherkatpyam();
                tbsherkatpyam2.Fk_afrad = id;
                tbsherkatpyam.Add(tbsherkatpyam2);
                return View(tbsherkatpyam);

            }
            else
            {
                return View(fioln2);

            }
            ////var fin = db.dbRelatinafrad.Where(p => p.ID == id).FirstOrDefault();
        }
        public ActionResult ModiriyatFaaliat_ListeFaaliyat_VaziyatModal(int id=0)
        {
            var fin = db.dbRelatinafrad.Where(p => p.ID == id).FirstOrDefault();
            return View(fin);
        }
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_GozareshFaaliyat()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            List<tbContractMoalefeDastmozdi> lstmoalafe2 = new List<tbContractMoalefeDastmozdi>();
            List<dbRelatinafrad> dbRelatinafrad = new List<dbRelatinafrad>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;
            lstMoalefeexcel2.fortariftatil = new List<MoalefeUserInfo22>();

            //// مقداردهی اولیه لیست fortariftatil
            //lstMoalefeexcel2.fortariftatil = new List<FunctionModel.MoalefeUserInfo22>();

            //lstMoalefeexcel2.fortariftatil = new List<FunctionModel.MoalefeUserInfo22>();

            // ایجاد یک نمونه از MoalefeUserInfo22 و مقداردهی FullName
           
            //fortariftatil. userInfo = new FunctionModel.MoalefeUserInfo22
            //{
            //    FullName = "نام کامل مورد نظر",
            //    PersonalCode = 123456,   // مقداردهی دیگر خواص
            //    anjam = 1,
            //    notanjam = 2,
            //    tagher = 3,
            //    darjaryan = 4,
            //    adam = 0,
            //    cant = 0,
            //    inmonth = DateTime.Now,  // مقدار دهی به تاریخ‌ها
            //    nextmonth = DateTime.Now.AddMonths(1)
            //};

            //// اضافه کردن userInfo به لیست fortariftatil
            //lstMoalefeexcel2.fortariftatil.Add(userInfo);
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
                var finddd = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == personal).ToList();
                foreach(var ittt in finddd)
                {
                    int anjam = 0;
                    int adamanjam = 0;
                    int tagher = 0;
                    int adamtanvanei = 0;
                    int adamemkan = 0;
                    int darjaryan = 0;
                    int pym = 0;

                    var fgh =db.dbvaZiat.Where(p=>p.dbRelatinafrad.FK_usr==ittt.FK_usr).ToList();
                    foreach( var it in fgh)

                    {
                        if (it.position == 1)
                        {
                            anjam += 1;

                        }
                        else if (it.position == 2)
                        {
                            adamanjam += 1;
                        }
                        else if (it.position ==3)
                        {
                            tagher += 1;
                        }
                        else if (it.position == 4)
                        {
                            adamtanvanei += 1;
                        }
                        else if (it.position == 5)
                        {
                            adamemkan += 1;
                        }
                        else if (it.position == 6)
                        {
                            darjaryan += 1;
                        }
                    }
                    var finnn = db.tbsherkatpyam.Where(p => p.FK_usr1 == personal && p.FK_isr2 == ittt.FK_usr).ToList();
                    foreach (var item in finnn)
                    {
                        if (item.notseee == 1)
                        {
                            pym += 1;
                        }
                    }
                    var obj = new MoalefeUserInfo22();
                    obj.adam = adamanjam;
                    obj.tagher = tagher;
                    obj.cant = adamemkan;
                    obj.notanjam = adamtanvanei;
                    obj.anjam = anjam;
                    obj.FullName = ittt.tbUsers.FullName;
                    obj.PersonalCode =(int) ittt.FK_usr;
                    obj.darjaryan = darjaryan;obj.pyam = pym;
                    lstMoalefeexcel2.fortariftatil.Add(obj);
                }
                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                //var find = db.dbRelatinafrad
                //             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                //             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                //             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                //             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                //             .ToList();

                //dbRelatinafrad.AddRange(find);
            }
            return View(lstMoalefeexcel2);
        }
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_TarifeTatilat()
        {
            return View(db.tbcraettatel.ToList());
        }
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_TarifeMovazafi()
        {
            return View();
        }
        public ActionResult ModiriyatFaaliat_TarifeMovazafiedit(int id=0)
        {
            var find = db.dbtarifmovazaf.Where(p => p.FK_usr == id).ToList();
            return View(find);
        }
        public ActionResult ModiriyatFaaliat_VaziyateHozooredit(int id = 0)
        {
            var find = db.dbvaziathozor.Where(p => p.ID == id).FirstOrDefault();
            return View(find);
        }
        public ActionResult ModiriyatFaaliat_TarifeMovazamodalk()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult ModiriyatFaaliat_VaziyateHozoor()
        {
            return View();
        }
        public ActionResult ModiriyatFaaliat_editertebatat(int id=0)
        {
            var find = db.dbrelatinperson.Where(p => p.FK_person == id).FirstOrDefault();
            return View(find);
        }
        
        public string ebatat_Edit(List<SaabWebProject.Models.DomainModels.dbrelatinperson> things)
        {
            try
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        List<dbrelatinperson> dbrelatinperson = new List<dbrelatinperson>();
                        int fk = 0;
                        foreach (var item in things)
                        {
                            fk = (int)item.FK_person;
                            
                            break;
                        }
                        bool cheack = true;
                        foreach (var item in things)
                        {
                            if (item.FK_usr == -1)
                            {
                                cheack = false; break;
                            }
                        }
                        // گرفتن لیست اولیه با FK_person مورد نظر




                        if (cheack==false)
                        {
                            foreach( var item in db.tbUsers.ToList())
                            {
                               dbrelatinperson dbrelatinperson2 = new dbrelatinperson();
                                dbrelatinperson2.FK_person = fk;
                                dbrelatinperson2.FK_usr = item.usr_ID;
                                dbrelatinperson.Add(dbrelatinperson2);
                            }
                            things = dbrelatinperson;
                        }

                        var find = db.dbrelatinperson.Where(p => p.FK_person == fk).ToList();

                        // پیدا کردن آیتم‌هایی که در find هستند ولی در things نیستند
                        var onlyInFind = find.Where(p => !things.Any(t => t.ID == p.ID)).ToList();

                        // پیدا کردن آیتم‌هایی که در things هستند ولی در find نیستند
                        var onlyInThings = things.Where(t => !find.Any(p => p.ID == t.ID)).ToList();
                        foreach (var item in onlyInFind)
                        {
                            item.FK_person = null;
                            db.SaveChanges();
                        }
                        db.dbrelatinperson.AddRange(onlyInThings);
                        db.SaveChanges();

                        // عملیات حذف و اضافه در جدول مورد نظر
                        //var ff = things.FirstOrDefault().FK_LevelID;
                        //var olds = db.tbReffrenceSaveLevelUser.Where(p => p.FK_LevelID == ff).ToList();
                        //if (olds != null)
                        //    db.tbReffrenceSaveLevelUser.RemoveRange(olds);

                        //db.tbReffrenceSaveLevelUser.AddRange(things);
                        //db.SaveChanges();

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
            catch (Exception ex)
            {
                return "false";
            }
        }

        //=================================================================================[ModiriyatFaaliat]End
        //=================================================================================[----------------]Start
        [AuthorizeAAA]
        public ActionResult TestSaab()
        {
            return View();
        }
        public ActionResult TableSheet()
        {
            return View();
        }
        public ActionResult TableSheetShow(int row = 50, int col = 30)
        {
            ViewBag.Row = row;
            ViewBag.Col = col;
            return View();
        }
        public ActionResult LogsView()
        {
            return View();
        }
        public ActionResult Z_peyvast()
        {
            return View();
        }
        public ActionResult TreeTab1()
        {
            return View();
        }
        public ActionResult TreeTab2()
        {
            return View();
        }
        public ActionResult TreeTab3()
        {
            return View();
        }
        //=================================================================================[----------------]End
        //--------------------------------------------------
        [AuthorizeAAA]
        public ActionResult MamoorinFaal()
        {
            return View();
        }
        public ActionResult MamoorinFaal_check()
        {
            return View();
        }
        //--------------------------------------------------
        [AuthorizeAAA]
        public ActionResult MamoorinEmkanat()
        {
            return View();
        }
        public ActionResult MamoorinEmkanat_check()
        {
            return View();
        }
        public ActionResult MamoorinEmkanat_vasile()
        {
            return View();
        }
        //--------------------------------------------------
        public string Magic_box_calculations2(int id=0)
        {
            List<tbNero> ner = new List<tbNero>();

            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/1 (2).xlsx"));
            var Moalefeexcelfile422 = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/1 (1).xlsx"));

            Row Row;
            var rc = db.tbMoadelPadashJarimeAyab.Where(p => p.Month == 4).ToList();
            var ty = db.dbTitletahod.Where(p=>p.Fk_NAME==id).ToList();
            foreach(var  it in ty)
            {
                var og = db.tbNero.Where(p => p.FKtitel == it.ID).ToList();
                ner.AddRange(og);
            }
            var informa = db.tbinformation.Where(p=>p.dbTitletahod.Fk_NAME==id).ToList();
            var nero = db.tbNero.Where(p => p.FKtitel == id).ToList();
            var nam = ner;
            var estan = db.tbestandardshenaght.Where(p=>p.ghragtdemandi==null && p.dbTitletahod.Fk_NAME == id && p.ghragtrostaei==null&&p.ghragtshahri==null).ToList();
            var to = db.tbinformation.Where(p => p.dbTitletahod.Fk_NAME == id).ToList();
            var to1 = db.tbestandardshenaght.Where(p => p.dbTitletahod.Fk_NAME == id&& p.ghragtdemandi == null && p.ghragtrostaei == null && p.ghragtshahri == null).ToList();
            var to12 = ner;
            var t = db.tbnergh.Where(p=>p.FK_name==id).ToList();
            //name and lastname and groupjob
            int counter = 4;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in to)
            {
                int val = 0;
                //var tt = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == item.usc_ID).FirstOrDefault();
                //if (tt != null)
                //{
                //    val = tt.Value ?? 0;
                //}
                Row = new Row() { Height = 20, Index = counter };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefau


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.value,
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

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.value,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 3
                //    }
                //});

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Month,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 4
                //    }
                //});
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Year,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 5
                //    }
                //});
                counter++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            int counter2 = 134;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in to1)
            {
                int val = 0;
                //var tt = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == item.usc_ID).FirstOrDefault();
                //if (tt != null)
                //{
                //    val = tt.Value ?? 0;
                //}
                Row = new Row() { Height = 20, Index = counter2 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.estandard,
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

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.value,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 3
                //    }
                //});

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Month,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 4
                //    }
                //});
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Year,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 5
                //    }
                //});
                counter2++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            int counter3 = 260;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in to12.Where(p => p.Estandard != null).ToList())
            {

                int val = 0;
                //var tt = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == item.usc_ID).FirstOrDefault();
                //if (tt != null)
                //{
                //    val = tt.Value ?? 0;
                //}
                Row = new Row() { Height = 20, Index = counter3 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.Estandard,
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

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.value,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 3
                //    }
                //});

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Month,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 4
                //    }
                //});
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Year,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 5
                //    }
                //});
                counter3++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            int counter4 = 277;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in to12.Where(p => p.Estandard != null).ToList())
            {
                int coun = counter4;

                foreach (var item1 in to12.Where(p => p.FKtitel == item.FKtitel && p.Estandard == null).ToList())

                {
                    Row = new Row() { Height = 20, Index = coun };

                    Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item1.value,
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

                    coun++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }


                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.value,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 3
                //    }
                //});

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Month,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 4
                //    }
                //});
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Year,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 5
                //    }
                //});
                counter4 += 17;
            }
            int counter5 = 409;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in t)
            {
                float val = (float)(item.ghalse + item.ezafeh + item.egareh + item.ayab + item.abzar + item.kosorat + item.sode);
                //var tt = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == item.usc_ID).FirstOrDefault();
                //if (tt != null)
                //{
                //    val = tt.Value ?? 0;
                //}
                Row = new Row() { Height = 20, Index = 409 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.ghalse,
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 410 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.ezafeh,
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 411 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.ayab,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 1
            }
        }); Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 412 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.egareh,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 1
            }
        }); Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 413 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.abzar,
                FontFamily = "B Nazanin",
                Bold = false,
                Enable = true,
                Wrap = false,
                FontSize = 12,
                Italic = false,
                Underline = false,
                Index = 1
            }
        }); Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 414 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.kosorat,
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 415 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.sode,
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

                Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 416 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = val,
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);

                Row = new Row() { Height = 20, Index = 417 };
                //var t = db..Where(p => p.fk_installment_ID == item.installment_ID && p.Month == ).FirstOrDefault();


                Row.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value = item.saatezafeh,
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
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.value,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 3
                //    }
                //});

                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Month,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 4
                //    }
                //});
                //Row.AddCells(new List<Cell>()
                //{
                //    new Cell()
                //    {
                //        Value = t.Year,
                //        FontFamily = "B Nazanin",
                //        Bold = false,
                //        Enable = true,
                //        Wrap = false,
                //        FontSize = 12,
                //        Italic = false,
                //        Underline = false,
                //        Index = 5
                //    }
                //});
                counter5++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            //int counterunit = 1;
            //var unit = UnitRepo.ListActiveUnit();
            //foreach (var item in unit)
            //{
            //    Row = new Row() { Height = 20, Index = counterunit };


            //    Row.AddCells(new List<Cell>()
            //    {
            //        new Cell()
            //        {
            //            Value = item.unt_Name,
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = 3
            //        }
            //    });
            //    counterunit++;
            //    Moalefeexcelfile.Sheets[1].AddRow(Row);
            //}














            var filePath = System.Web.HttpContext.Current.Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/");
            var fileName = "1" + ".xlsx";
            var fullPath = Path.Combine(filePath, fileName);

            // بررسی وجود فایل در مسیر مشخص شده
            //if (System.IO.File.Exists(fullPath))
            //{
            //    // اگر فایل وجود داشته باشد، حذف می‌شود
            //    System.IO.File.Delete(fullPath);
            //}

            //// ذخیره فایل جدید در مسیر مشخص شده
            //Moalefeexcelfile.Save(fullPath);

            string tempFileName = Path.Combine(filePath, "1.xlsx");
            Moalefeexcelfile.Save(tempFileName);
            var workbook = Workbook.Load(Server.MapPath("~/Areas/Salaries/Content/UserFunctionsFile/1.xlsx"));
            //var Moalefeexcelfile22 = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/auto.xlsm"));

            string filePath222 = Server.MapPath("~/Content/ExcelFiles/auto.xlsm");

            FileInfo fileInfo = new FileInfo(filePath222);

            //using (ExcelPackage package = new ExcelPackage(fileInfo))
            //{
            //    ExcelWorkbook workbook2 = package.Workbook;
            //    ExcelWorksheet worksheet = workbook2.Worksheets[0]; // انتخاب شیت اول

            //    // اعمال تغییرات در شیت
            //    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            //    {
            //        worksheet.Cells[row,1].Value = worksheet.Cells[row, 0].Value;
            //    }

            //    // ذخیره فایل
            //    package.Save();
            //}
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    var Count_Sheets3 = workbook.Sheets.Count;

                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[2];
                        var Rows = sheet.Rows;
                        var sheet2 = Moalefeexcelfile422.Sheets[2];

                        var rows = sheet2.Rows;

                        if (Rows.Count >= 1)
                        {
                            foreach (var row in rows)
                            {
                                // چک کردن اینکه آیا تعداد ستون‌ها حداقل سه عدد است
                                if (row.Cells.Count >= 3)
                                {
                                    row.Cells[2].Value = row.Cells[1].Value;
                                    //row.Cells[4].Value = row.Cells[3].Value;
                                    //row.Cells[4].Value = row.Cells[2].Value;

                                    // کپی کردن مستقیم مقدار ستون 2 به ستون 3
                                    //row.Cells[4].Value = row.Cells[2].Value;
                                }
                                else
                                {
                                    return " خطای ارزیابی در سطر " + (row.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                }
                            }
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title = workbook.Sheets[2].Rows[0].Cells;
                                var List = workbook.Sheets[2].Rows;
                                var count = workbook.Sheets[2].Rows.Count();
                                //for (int i = 5; i < title.Count; i++)
                                //{
                                //    lstMoalefe.Add(title[i].Value.ToString());
                                //}

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                tbresultpishkansehnaght resul = new tbresultpishkansehnaght();
                                for (int i = 0; i < 1; i++)
                                {
                                    var row = workbook.Sheets[2].Rows[0];
                                    //var row10 = Moalefeexcelfile22.Sheets[0].Rows[0];
                                    //var Name344 = row10.Cells[0].Value.ToString();




                                    var Name900 = row.Cells[1];
                                    if (Name900.Value == null)
                                    {
                                        Console.WriteLine("Cell value is null at row " + i);
                                        return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "دارای خطا می باشد";
                                    }

                                    string cellValue = Name900.Value.ToString();
                                    Console.WriteLine("Cell value: " + cellValue + " at row " + i);

                                    //if (cellValue.StartsWith("#"))
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "دارای خطا می باشد";
                                    //}








                                    var Name90 = row.Cells[1];
                                    //if (Name90.Value == null || Name90.Value.ToString().StartsWith("#"))
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "دارای خطا می باشد";
                                    //}

                                    int mlfvlfsh_Year11133;
                                    bool isNumeric3453 = int.TryParse(Name90.Value.ToString(), out mlfvlfsh_Year11133);

                                    //if (!isNumeric3453)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}



                                    var Name = row.Cells[1];
                                    int mlfvlfsh_Year111;
                                    bool isNumeric345 = int.TryParse(Name.Value.ToString(), out mlfvlfsh_Year111);

                                    //if (!isNumeric345)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.count_person = mlfvlfsh_Year111;
                                    row = workbook.Sheets[2].Rows[1];
                                    var Name2 = row.Cells[1];
                                    int mlfvlfsh_Year;
                                    bool isNumeric = int.TryParse(Name2.Value.ToString(), out mlfvlfsh_Year);

                                    //if (!isNumeric)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.count_tagsoso_littel = mlfvlfsh_Year;
                                    row = workbook.Sheets[2].Rows[2];

                                    var Name3 = row.Cells[1];
                                    float mlfvlfsh_Year22;
                                    bool isNumeric2 = float.TryParse(Name3.Value.ToString(), out mlfvlfsh_Year22);

                                    //if (!isNumeric2)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.persent_tagsoso_littel = mlfvlfsh_Year22;
                                    row = workbook.Sheets[2].Rows[4];

                                    var Name4 = row.Cells[1];
                                    float mlfvlfsh_Year224;
                                    bool isNumeric24 = float.TryParse(Name4.Value.ToString(), out mlfvlfsh_Year224);

                                    //if (!isNumeric24)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.persent_person_tavarom = mlfvlfsh_Year224;
                                    row = workbook.Sheets[2].Rows[3];

                                    var Name5 = row.Cells[1];
                                    int mlfvlfsh_Year225;
                                    bool isNumeric25 = int.TryParse(Name5.Value.ToString(), out mlfvlfsh_Year225);

                                    //if (!isNumeric25)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.count_person_tavarom = mlfvlfsh_Year225;
                                    row = workbook.Sheets[2].Rows[6];
                                    var Name6 = row.Cells[1];
                                    float mlfvlfsh_Year2256;
                                    bool isNumeric256 = float.TryParse(Name6.Value.ToString(), out mlfvlfsh_Year2256);

                                    //if (!isNumeric256)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.persent_person_jobran = mlfvlfsh_Year2256;
                                    row = workbook.Sheets[2].Rows[5];
                                    var Name7 = row.Cells[1];
                                    int mlfvlfsh_Year2257;
                                    bool isNumeric257 = int.TryParse(Name7.Value.ToString(), out mlfvlfsh_Year2257);

                                    //if (!isNumeric257)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.count_person_jobran = mlfvlfsh_Year2257;
                                    row = workbook.Sheets[2].Rows[7];
                                    var Name8 = row.Cells[1];
                                    int mlfvlfsh_Year22578;
                                    bool isNumeric2578 = int.TryParse(Name8.Value.ToString(), out mlfvlfsh_Year22578);

                                    //if (!isNumeric2578)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}


                                    resul.count_need_person = mlfvlfsh_Year22578;
                                    row = workbook.Sheets[2].Rows[8];
                                    var Name81 = row.Cells[1];
                                    float mlfvlfsh_Year225781;
                                    bool isNumeric25781 = float.TryParse(Name81.Value.ToString(), out mlfvlfsh_Year225781);

                                    //if (!isNumeric25781)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}


                                    resul.Condition1_mojod = mlfvlfsh_Year225781;
                                    row = workbook.Sheets[2].Rows[9];
                                    var Name812 = row.Cells[1];
                                    float mlfvlfsh_Year2257812;
                                    bool isNumeric257812 = float.TryParse(Name812.Value.ToString(), out mlfvlfsh_Year2257812);

                                    //if (!isNumeric257812)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}

                                    resul.Condition1_jobran = mlfvlfsh_Year2257812;
                                    row = workbook.Sheets[2].Rows[10];
                                    var Name8123 = row.Cells[1];
                                    float mlfvlfsh_Year22578123;
                                    bool isNumeric2578123 = float.TryParse(Name8123.Value.ToString(), out mlfvlfsh_Year22578123);

                                    //if (!isNumeric2578123)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.Condition1_beheneh = mlfvlfsh_Year22578123;
                                    var row2 = workbook.Sheets[2].Rows[11];
                                    //var Name81234 = row.Cells[1];
                                    var Name81234 = row2.Cells[1];
                                    float mlfvlfsh_Year225781234;
                                    bool isNumeric25781234 = float.TryParse(Name81234.Value.ToString(), out mlfvlfsh_Year225781234);

                                    //if (!isNumeric25781234)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}


                                    resul.Condition1_edehal = mlfvlfsh_Year225781234;
                                    row = workbook.Sheets[2].Rows[12];
                                    var Name812345 = row.Cells[1];
                                    float mlfvlfsh_Year2257812345;
                                    bool isNumeric257812345 = float.TryParse(Name812345.Value.ToString(), out mlfvlfsh_Year2257812345);

                                    //if (!isNumeric257812345)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}


                                    resul.Condition2_mojod = mlfvlfsh_Year2257812345;
                                    row = workbook.Sheets[2].Rows[13];
                                    var Name8123456 = row.Cells[1];
                                    float mlfvlfsh_Year22578123456;
                                    bool isNumeric2578123456 = float.TryParse(Name8123456.Value.ToString(), out mlfvlfsh_Year22578123456);

                                    //if (!isNumeric2578123456)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}




                                    resul.Condition2_jobran = mlfvlfsh_Year22578123456;
                                    var row4 = workbook.Sheets[2].Rows[14];
                                    //var Name812345678993 = row4.Cells[2].Value.ToString();
                                    //var cellValue = row4.Cells[2].Value;
                                    //if (cellValue != null)
                                    //{
                                    //    var Name81234567899322 = cellValue.ToString();

                                    //}
                                    //var cellValue2 = row4.Cells[3].Value;
                                    //if (cellValue != null)
                                    //{
                                    //    var Name812345678993222 = cellValue.ToString();

                                    //}
                                    var Name81234567 = row4.Cells[1];
                                    float mlfvlfsh_Year225781234567;
                                    bool isNumeric25781234567 = float.TryParse(Name81234567.Value.ToString(), out mlfvlfsh_Year225781234567);

                                    //if (!isNumeric25781234567)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    resul.Condition2_jobran = mlfvlfsh_Year225781234567;
                                    var row22 = workbook.Sheets[2].Rows[14];
                                    var Name812345678 = row22.Cells[1];
                                    //var Name81234567899 = row22.Cells[2].Value.ToString();

                                    float mlfvlfsh_Year2257812345678;
                                    bool isNumeric257812345678 = float.TryParse(Name812345678.Value.ToString(), out mlfvlfsh_Year2257812345678);

                                    //if (!isNumeric257812345678)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}

                                    resul.Condition2_beheneh = mlfvlfsh_Year2257812345678;
                                    row = workbook.Sheets[2].Rows[15];
                                    var Name8123456789 = row.Cells[1];
                                    float mlfvlfsh_Year22578123456789;
                                    bool isNumeric2578123456789 = float.TryParse(Name8123456789.Value.ToString(), out mlfvlfsh_Year22578123456789);

                                    //if (!isNumeric2578123456789)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}



                                    resul.Condition2_edehal = mlfvlfsh_Year22578123456789;
                                    //row = workbook.Sheets[2].Rows[16];
                                    //var Name81234567890 = row.Cells[1];
                                    //float mlfvlfsh_Year225781234567890;
                                    //bool isNumeric25781234567890 = float.TryParse(Name81234567890.Value.ToString(), out mlfvlfsh_Year225781234567890);

                                    //if (!isNumeric25781234567890)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}

                                    //var Name5 = row.Cells[4];
                                    //int mlfvlfsh_Year225;
                                    //bool isNumeric25 = int.TryParse(Name5.Value.ToString(), out mlfvlfsh_Year225);

                                    //if (!isNumeric25)
                                    //{
                                    //    return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                    //}
                                    //tbSoratvaziat obj5 = new tbSoratvaziat();
                                    //var id = py.Where(p => p.pec_Title.Contains(name)).FirstOrDefault();
                                    //tbSoratvaziat obj5 = new tbSoratvaziat();
                                    //obj5.number_sorat = mlfvlfsh_Year224;
                                    //obj5.nergh = mlfvlfsh_Year2256;
                                    //obj5.gham = (long)mlfvlfsh_Year2257;
                                    //obj5.Enerjy = mlfvlfsh_Year225;
                                    //obj5.Fk_pymn = id.pec_ID;
                                    //obj5.Finalvaluesorat = mlfvlfsh_Year22578;
                                    //obj5.Year = mlfvlfsh_Year;
                                    //obj5.Final_accept = true;
                                    //obj5.month = mlfvlfsh_Year22;
                                    //objectsToAdd.Add(obj5);
                                    var ex = db.tbresultpishkansehnaght.Where(p => p.FK_NAME == id).FirstOrDefault();
                                    if (ex != null)
                                    {
                                        ex.count_tagsoso_littel = resul.count_tagsoso_littel;
                                        ex.persent_tagsoso_littel = resul.persent_tagsoso_littel;
                                        ex.persent_tagsoso_littel = resul.persent_tagsoso_littel;
                                        ex.count_person = resul.count_person;
                                        ex.persent_person_tavarom = resul.persent_person_tavarom;
                                        ex.count_person_tavarom = resul.count_person_tavarom;
                                        ex.persent_person_jobran = resul.persent_person_jobran;
                                        ex.count_person_jobran = resul.count_person_jobran;
                                        ex.count_need_person = resul.count_need_person;
                                        ex.Condition1_mojod = resul.Condition1_mojod;
                                        ex.Condition1_jobran = resul.Condition1_jobran;
                                        ex.Condition2_mojod = resul.Condition2_mojod;
                                        ex.Condition2_jobran = resul.Condition2_jobran;
                                        ex.Condition2_beheneh = resul.Condition2_beheneh;
                                        ex.Condition2_edehal = resul.Condition2_edehal;
                                        ex.Condition1_edehal = resul.Condition1_edehal;
                                        //db.SaveChanges();
                                    }
                                    else
                                    {
                                        resul.FK_NAME = id;
                                        db.tbresultpishkansehnaght.Add(resul);
                                        //db.SaveChanges();
                                    }
                                    //foreach (var it in lstMoalefe)
                                    //{
                                    //    tbSoratvaziat obj5 = new tbSoratvaziat();
                                    //    var id = py.Where(p => p.pec_Title.Contains(name)).FirstOrDefault();
                                    //    var id2 = sabad.Where(p => p.Title.Contains(it)).FirstOrDefault();

                                    //    var Value = row.Cells[shomarande];
                                    //    var value222 = Value.Value ?? null;
                                    //    double? countDays2;
                                    //    if (double.TryParse(value222.ToString(), out double parsedValue3))
                                    //    {
                                    //        countDays2 = parsedValue3;
                                    //    }
                                    //    else
                                    //    {
                                    //        countDays2 = 0; // Or handle the case where the value cannot be parsed
                                    //    }
                                    //    obj5.number_sorat = mlfvlfsh_Year224;
                                    //    obj5.FK_Sabad = id2.ID;
                                    //    obj5.Fk_pymn = id.pec_ID;
                                    //    obj5.role = mlfvlfsh_Year225;
                                    //    obj5.value = (long)countDays2;
                                    //    obj5.Year = mlfvlfsh_Year;
                                    //    obj5.Final_accept = true;
                                    //    obj5.month = mlfvlfsh_Year22;
                                    //    objectsToAdd.Add(obj5);
                                    //    shomarande++;

                                    //}


                                }

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








            //System.IO.File.Replace(tempFileName, fullPath, null);


            //return Moalefeexcelfile;
        }















        [HttpPost]
        public string UpdateRecord( List<shenaght> modelList)
        {
             List< tbinformation> tbinformation2 =  new List<tbinformation>();
            List<dbcountshenaght> tbinformation24 = new List<dbcountshenaght>();

            using (var context = new SaabEntities())
            {
                var fi = false;
                var gr = false;
                foreach (var model in modelList)
                {
                    
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    if (model.ghragtdemandi != 0&& model.ghragtrostaei != 0&& model.ghragtshahri!=0)
                    {
                        var fi33 = context.dbcountshenaght.Where(p => p.FK_Name == model.id).FirstOrDefault();
                        if (fi33 != null)
                        {
                            if (model.ghragtdemandi != null)
                            {
                                fi33.demandi = model.ghragtdemandi;
                            }
                            if (model.ghragtrostaei != null)
                            {
                                fi33.rostaei = model.ghragtrostaei;
                            }
                            if (model.ghragtshahri != null)
                            {
                                fi33.shahri = model.ghragtshahri;
                            }
                            context.SaveChanges();
                        }

                        else
                        {
                            dbcountshenaght tbinformation223 = new dbcountshenaght();
                            tbinformation223.FK_Name = model.id;
                            if (model.ghragtdemandi != null)
                            {
                                tbinformation223.demandi = model.ghragtdemandi;
                            }
                            if (model.ghragtrostaei != null)
                            {
                                tbinformation223.rostaei = model.ghragtrostaei;
                            }
                            if (model.ghragtshahri != null)
                            {
                                tbinformation223.shahri = model.ghragtshahri;
                            }
                            tbinformation24.Add(tbinformation223);

                        }
                    }
             


                    if (model.type != null)
                    {
                        if (model.type == "fardi")
                        {
                            var find = context.tbinformation.Where(p => p.fktitel == model.id && p.fardi == true).ToList();
                            context.tbinformation.RemoveRange(find);
                            context.SaveChanges();
                        }
                        else if (model.type == "grohi")
                        {
                            var find = context.tbinformation.Where(p => p.fktitel == model.id && p.grohi == true).ToList();
                            context.tbinformation.RemoveRange(find);
                            context.SaveChanges();
                        }
                    }
                  
                    if(model.type != null)
                    {
                        tbinformation tbinformation = new tbinformation();
                        if (model.type == "fardi")
                            tbinformation.fardi = true;
                        if (model.type == "grohi")
                            tbinformation.grohi = true;
                        if (model.phase == "takfaz")
                            tbinformation.takfaz = true;
                        if (model.phase == "sefaz")
                            tbinformation.sefaz = true;
                        if (model.area == "shahri")
                            tbinformation.Shahri = true;
                        if (model.area == "roostayi")
                            tbinformation.rostaeei = true;
                        if (model.phase == "demandi")
                            tbinformation.demandi = true;
                        tbinformation.value = model.value;
                        tbinformation.fktitel = model.id;
                        tbinformation2.Add(tbinformation);
                    }
                   
                    
                  

                }
                context.tbinformation.AddRange(tbinformation2);
                context.SaveChanges();
                context.dbcountshenaght.AddRange(tbinformation24);
                context.SaveChanges();
                return "true";
            }
        }

        [HttpPost]
        public string UpdateRecord2(List<shenaght> modelList)
        {
            List<tbestandardshenaght> tbinformation2 = new List<tbestandardshenaght>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    if (model.type == "fardi")
                    {
                        var find = context.tbestandardshenaght.Where(p => p.fktitel == model.id && p.fardi == true).ToList();
                        context.tbestandardshenaght.RemoveRange(find);
                        context.SaveChanges();
                    }
                    else if (model.type == "grohi")
                    {
                        var find = context.tbestandardshenaght.Where(p => p.fktitel == model.id && p.grohi == true).ToList();
                        context.tbestandardshenaght.RemoveRange(find);
                        context.SaveChanges();
                    }
                    var find2 = context.tbestandardshenaght.Where(p => p.fktitel == model.id && p.ghragtrostaei!=0).ToList();
                    context.tbestandardshenaght.RemoveRange(find2);
                    context.SaveChanges();
                    var find3 = context.tbestandardshenaght.Where(p => p.fktitel == model.id && p.ghragtshahri != 0).ToList();
                    context.tbestandardshenaght.RemoveRange(find3);
                    context.SaveChanges();
                    var find4 = context.tbestandardshenaght.Where(p => p.fktitel == model.id && p.ghragtdemandi != 0).ToList();
                    context.tbestandardshenaght.RemoveRange(find4);
                    context.SaveChanges();

                    tbestandardshenaght tbinformation = new tbestandardshenaght();
                    if (model.type == "fardi")
                        tbinformation.fardi = true;
                    if (model.type == "grohi")
                        tbinformation.grohi = true;
                    if (model.phase == "takfaz")
                        tbinformation.takfaz = true;
                    if (model.phase == "sefaz")
                        tbinformation.sefaz = true;
                    if (model.area == "shahri")
                        tbinformation.Shahri = true;
                    if (model.area == "roostayi")
                        tbinformation.rostaeei = true;
                    if (model.phase == "demandi")
                        tbinformation.demandi = true;
                    if (model.ghragtrostaei != 0)
                    {
                        tbinformation.ghragtrostaei = model.ghragtrostaei;

                    }
                    if (model.ghragtshahri != 0)
                    {
                        tbinformation.ghragtshahri = model.ghragtshahri;

                    }
                    tbinformation.estandard = model.value;
                    tbinformation.fktitel = model.id;
                    tbinformation2.Add(tbinformation);



                }
                context.tbestandardshenaght.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }


        public string UpdateRecaddone(List<dbTitletahod> modelList)
        {
            List<dbTitletahod> tbinformation2 = new List<dbTitletahod>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    var find = context.dbTitletahod.Where(p => p.Fk_NAME == model.Fk_NAME && p.Title == model.Title).FirstOrDefault();  
                    if (find!=null)
                    {
                        if (model.grohi != null)
                        {
                            find.grohi = model.grohi;

                        }
                        if (model.Fardi != null)
                        {
                            find.Fardi = model.Fardi;

                        }
                        find.count=model.count;
                        context.SaveChanges();

                 
                    }
                    else
                    {
                        dbTitletahod dbTitletahod = new dbTitletahod();
                        dbTitletahod.count = model.count;
                        dbTitletahod.Fardi= model.Fardi;
                        dbTitletahod.grohi= model.grohi;
                        dbTitletahod.Title=model.Title;
                        dbTitletahod.Fk_NAME= model.Fk_NAME;
                        tbinformation2.Add(dbTitletahod);
                    }




                }
                context.dbTitletahod.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }



        public string UpdateRecord4(List<shenaght2> modelList)
        {
            List<tbNero> tbinformation2 = new List<tbNero>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    var find = context.tbNero.Where(p => p.FKtitel == model.id).ToList();
                    if (find.Count != 0)
                    {
                        context.tbNero.RemoveRange(find);
                        context.SaveChanges();

                    }
                }
                    foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    //var find = context.tbNero.Where(p => p.FKtitel == model.id).ToList();
                    //if (find.Count!=0)
                    //{
                    //    context.tbNero.RemoveRange(find);
                    //    context.SaveChanges();

                    //}

                    tbNero tbinformation = new tbNero();
                    if (model.Estandard != 0)
                    {
                        tbinformation.FKtitel = model.id;
                        tbinformation.Estandard = model.Estandard;
                        tbinformation.FkTitlerelat = null;
                        tbinformation.value = 0;
                    }
                    else
                    {
                        tbinformation.FKtitel = model.id;
                        tbinformation.Estandard = null;
                        tbinformation.FkTitlerelat = model.FkTitlerelat;
                        tbinformation.value = model.value;
                    }
                    tbinformation2.Add(tbinformation);



                }
                context.tbNero.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }

        public string UpdateRecord9(List<tbNamshnaght> modelList)
        {
            List<tbNamshnaght> tbinformation2 = new List<tbNamshnaght>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    var find = context.tbNamshnaght.Where(p => p.Name == model.Name).FirstOrDefault();
                    if (find != null)
                    {
                        find.Name = model.Name;
                      
                        context.SaveChanges();
                    }
                    else
                    {
                        tbNamshnaght tbinformation = new tbNamshnaght();
                        tbinformation.Name = model.Name;
                      
                        tbinformation2.Add(tbinformation);
                    }




                }
                context.tbNamshnaght.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }

        public string UpdateRecor100(string Name="",int type2=0,int fkname=0)
        {
            List<dbTitletahod> dbTitletahod = new List<dbTitletahod>();
            using (var context = new SaabEntities())
            {
                dbTitletahod dbTitletahod2 = new dbTitletahod();
                dbTitletahod2.Fk_NAME = fkname;
                dbTitletahod2.Title = Name;
                if (type2 == 1)
                {
                    dbTitletahod2.Fardi = true;
                    dbTitletahod2.grohi = false;
                }
                else if (type2 == 2)
                {
                    dbTitletahod2.grohi = true;
                    dbTitletahod2.Fardi = false;
                }
                else if(type2 == 3)
                {
                    dbTitletahod2.Fardi = true;
                    dbTitletahod2.grohi = true;
                }

                context.dbTitletahod.Add(dbTitletahod2);
                context.SaveChanges();
                return "true";
            }
        }
        public string UpdateRecord6(List<tbnergh> modelList)
        {
            List<tbnergh> tbinformation2 = new List<tbnergh>();
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    var find = context.tbnergh.Where(p => p.FK_name == model.FK_name).FirstOrDefault();
                    if (find!=null)
                    {
                        find.abzar = model.abzar;
                        find.ghalse = model.ghalse;
                        find.ayab = model.ayab;
                        find.kosorat = model.kosorat;
                        find.sode = model.sode;
                        find.saatezafeh = model.saatezafeh;
                        find.egareh = model.egareh;
                        find.ezafeh = model.ezafeh;
                        find.FK_name = model.FK_name;
                        context.SaveChanges();
                    }
                    else
                    {
                        tbnergh tbinformation = new tbnergh();
                        tbinformation.abzar = model.abzar;
                        tbinformation.ghalse = model.ghalse;
                        tbinformation.ayab = model.ayab;
                        tbinformation.kosorat = model.kosorat;
                        tbinformation.sode = model.sode;
                        tbinformation.saatezafeh = model.saatezafeh;
                        tbinformation.egareh = model.egareh;
                        tbinformation.ezafeh = model.ezafeh;
                        tbinformation.FK_name = model.FK_name;
                        tbinformation2.Add(tbinformation);
                    }
                    



                }
                context.tbnergh.AddRange(tbinformation2);
                context.SaveChanges();
                return "true";
            }
        }

        public string UpdateRecord7(List<tbshaghes> modelList)
        {
            using (var context = new SaabEntities())
            {
                foreach (var model in modelList)
                {
                    //var record = context.tbinformation.Find(model.id&&);
                    //if (record != null)
                    //{
                    //    if (model.type == "fardi")
                    //        record.fardi = true;
                    //    if (model.type == "grohi")
                    //        record.grohi = true;
                    //    if (model.phase == "takfaz")
                    //        record.takfaz = true;
                    //    if (model.phase == "sefaz")
                    //        record.sefaz = true;
                    //    if (model.area == "shahri")
                    //        record.Shahri = true;
                    //    if (model.area == "rostaeei")
                    //        record.rostaeei = true;
                    //    if (model.phase == "demandi")
                    //        record.demandi = true;
                    //    record.value=model.value;
                    //    context.SaveChanges();
                    //}
                    var fin= context.tbshaghes.Where(p=>p.ID==model.ID).FirstOrDefault();
                    if(fin!=null)
                    {
                        fin.value=model.value;
                        context.SaveChanges();

                    }




                }
        
                return "true";
            }
        }

        //====================================
        public ActionResult News()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult Dashboard2()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/Dashboard2.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult slidshoe()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult Document_Dashboard()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/Document_Dashboard.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult Main_KPI_Dashboard()
        {
            using (var db = new SaabWebProject.Models.DomainModels.SaabEntities())
            {
                return PartialView("~/Areas/Reporter/Views/Dashboard/Main_KPI_Dashboard.cshtml", db.tbVosoolMotalebat.ToList());

            }
        }
        [AuthorizeAAA]

        public ActionResult Arzyabi_KPI_Dashboard()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/Arzyabi_KPI_Dashboard.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult MaMoorin_Dashboard2(int id = 0)
        {

            var Model = new List<tbMoalefeValuePishkhan>();
            Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == id ).OrderByDescending(t=>t.mlfval_ID).ToList();
            return PartialView("~/Areas/Reporter/Views/Dashboard/MaMoorin_Dashboard.cshtml", Model);

        }
        [AuthorizeAAA]
        public ActionResult MaMoorin_Dashboard()
        {
            var Model = new List<tbMoalefeValuePishkhan>();
            int? personal = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
            }
            if (UserStuf.PosID == "2175")
            {
                var USSER = db.tbMoalefeValuePishkhan.FirstOrDefault().mlfval_FKUser;
                Model = db.tbMoalefeValuePishkhan.OrderByDescending(t => t.mlfval_ID).Where(p => p.mlfval_FKUser == USSER).ToList();
            }
            else
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == personal).OrderByDescending(t => t.mlfval_ID).ToList();
            }
            return PartialView("~/Areas/Reporter/Views/Dashboard/MaMoorin_Dashboard.cshtml", Model);
        }
        public ActionResult MaMoorin_EXAMPEL()
        {
            var Model = new List<tbMoalefeValuePishkhan>();
            int? personal = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
            }
            if (UserStuf.PosID == "2175")
            {
                var USSER = db.tbMoalefeValuePishkhan.FirstOrDefault().mlfval_FKUser;
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == USSER).ToList();
            }
            else
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == personal).ToList();
            }
            return PartialView("~/Areas/Reporter/Views/Dashboard/MaMoorin_EXAMPEL.cshtml", Model);
        }
        [AuthorizeAAA]
        public ActionResult MaMoorin_EXAMPEL2(int id = 0)
        {
            var Model = new List<tbMoalefeValuePishkhan>();
            Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == id).ToList();
            return PartialView("~/Areas/Reporter/Views/Dashboard/MaMoorin_EXAMPEL.cshtml", Model);

        }
        [AuthorizeAAA]
        public ActionResult Peyman_Dashboard(int? PeymanID = null)
        {
            //if (cookie_user != null)
            //{
            //    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //    var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
            //    personal = user.usr_ID;
            //}
            if (UserStuf.PosID == "9" && PeymanID == null)
            {
                PeymanID = db.tbPeymanContracts.Where(p=>p.Inactive!=true).FirstOrDefault()?.pec_ID;
            }

            if (PeymanID == null)
            {
                int? personal = 0;
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
                var tarafdovom = db.tbCompanies.Where(p => p.FK_ManagerID == personal).FirstOrDefault();
                if (tarafdovom != null)
                {

                    PeymanID = db.tbPeymanContracts.FirstOrDefault(p => p.FK_UserTarafDovvom == tarafdovom.ID && p.Inactive != true).pec_ID;
                }
                else
                {
                    PeymanID = db.tbPeymanContracts.FirstOrDefault().pec_ID;
                }

            }

            List<tbMoalefeValuePishkhan> Model = new List<tbMoalefeValuePishkhan>();
            List<tbMoalefeValuePishkhan> Result = new List<tbMoalefeValuePishkhan>();
            if (PeymanID != null)
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == PeymanID).ToList();
                foreach (var item in Model)
                {

                    var list = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKMoalafeDastmozdi == item.mlfval_FKMoalafeDastmozdi && p.mlfval_FKPeyman == PeymanID).ToList();
                    var maxyear = list.Max(p => p.mlfval_Year);
                    list = list.Where(p => p.mlfval_Year == maxyear).ToList();
                    var maxmonth = list.Max(p => p.mlfval_Month);
                    var moalefe = list.FirstOrDefault(p => p.mlfval_Month == maxmonth);
                    Result.Add(moalefe);
                }
                // var maxyear = Model.Max(p => p.mlfval_Year);

            }
            else//in bayad pak shavad
            {
                // Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == db.tbPeymanContracts.FirstOrDefault().pec_ID).ToList();
                Result = new List<tbMoalefeValuePishkhan>();
            }
            return PartialView("~/Areas/Reporter/Views/Dashboard/Peyman_Dashboard.cshtml", Result);
        }
        public ActionResult viewmoadtestrtt() {


            return PartialView("~/Areas/Reporter/Views/Dashboard/viewmoadtestrtt.cshtml");

        }

        public ActionResult Peyman_Dashboard2(int? PeymanID = null)
        {
            //if (cookie_user != null)
            //{
            //    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //    var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
            //    personal = user.usr_ID;
            //}
            if (UserStuf.PosID == "9" && PeymanID == null)
            {
                PeymanID = db.tbPeymanContracts.FirstOrDefault().pec_ID;
            }
            if (PeymanID == null)
            {
                int? personal = 0;
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
                var tarafdovom = db.tbCompanies.Where(p => p.FK_ManagerID == personal).FirstOrDefault();
                if (tarafdovom != null)
                {

                    PeymanID = db.tbPeymanContracts.FirstOrDefault(p => p.FK_UserTarafDovvom == tarafdovom.ID && p.Inactive != true).pec_ID;
                }
                else
                {
                    PeymanID = db.tbPeymanContracts.FirstOrDefault().pec_ID;
                }

            }

            List<tbMoalefeValuePishkhan> Model = new List<tbMoalefeValuePishkhan>();
            List<tbMoalefeValuePishkhan> Result = new List<tbMoalefeValuePishkhan>();
            if (PeymanID != null)
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == PeymanID).ToList();
                foreach (var item in Model)
                {

                    var list = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKMoalafeDastmozdi == item.mlfval_FKMoalafeDastmozdi && p.mlfval_FKPeyman == PeymanID).ToList();
                    var maxyear = list.Max(p => p.mlfval_Year);
                    list = list.Where(p => p.mlfval_Year == maxyear).ToList();
                    var maxmonth = list.Max(p => p.mlfval_Month);
                    var moalefe = list.FirstOrDefault(p => p.mlfval_Month == maxmonth);
                    Result.Add(moalefe);
                }
                // var maxyear = Model.Max(p => p.mlfval_Year);

            }
            else//in bayad pak shavad
            {
                // Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == db.tbPeymanContracts.FirstOrDefault().pec_ID).ToList();
                Result = new List<tbMoalefeValuePishkhan>();
            }
            return PartialView("~/Areas/Contracts/Views/JobGroups/persianCalendar.cshtml", Result);
        }


        [AuthorizeAAA]
        public ActionResult SharhVazayef()
        {
            return PartialView("~/Areas/Reporter/Views/Dashboard/SharhVazayef.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult Download(int id)
        {
            if (id == 1)
            {
                byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Content/PDFs/مانده پیمان.pdf"));

                return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "mamorian" + ".pdf");
            }
            else
            {
                byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Content/PDFs/مانده مامورین.pdf"));
                return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "peymanha" + ".pdf");
            }
        }

        [AuthorizeAAA]
        public ActionResult DNqaradad()//قراردادها با شماره پرسنلی بایستی ذخیره شوند
        {
            int? personal = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_Personal_ID;
            }

            byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Content/PDFs/" + personal + ".pdf"));

            return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, personal + ".pdf");
        }
        [AuthorizeAAA]
        public ActionResult _KarkonanDarDashboardPeyman()
        {


            return PartialView("_KarkonanDarDashboardPeyman");


        }


        public ActionResult _KarkonanDetail(int? UsrID = null)
        {
            List<tbMoalefeValuePishkhan> Model = new List<tbMoalefeValuePishkhan>();
            if (UsrID == null)
            {
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                    var tarafdovom = db.tbCompanies.Where(p => p.FK_ManagerID == user.usr_ID).FirstOrDefault();
                    int PeymanID = 0;
                    if (UserStuf.PosID == "9" || tarafdovom == null)
                    {
                        PeymanID = db.tbPeymanContracts.FirstOrDefault().pec_ID;
                    }
                    else
                    {
                        PeymanID = db.tbPeymanContracts.FirstOrDefault(p => p.FK_UserTarafDovvom == tarafdovom.ID && p.Inactive != true).pec_ID;

                    }


                    var usser = db.Link_User_And_Peyman.OrderByDescending(s=>s.Link_User_And_Peyman_ID).FirstOrDefault(p => p.FK_Peyman_ID == PeymanID&&p.tbUsers.tbMoalefeValuePishkhan.Any(w=>w.mlfval_FKUser==p.FK_User_ID)).FK_User_ID;
                    Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == usser).ToList();

                }
                

            }


            else
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == UsrID).ToList();
            }
            return PartialView("_KarkonanDetail", Model);

        }
        public ActionResult _KarkonanDetail2(int? UsrID = null)
        {
            List<tbMoalefeValuePishkhan> Model = new List<tbMoalefeValuePishkhan>();
            if (UsrID == null)
            {
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                    var tarafdovom = db.tbCompanies.Where(p => p.FK_ManagerID == user.usr_ID).FirstOrDefault();
                    int PeymanID = 0;
                    if (UserStuf.PosID == "9" || tarafdovom == null)
                    {
                        PeymanID = db.tbPeymanContracts.FirstOrDefault().pec_ID;
                    }
                    else
                    {
                        PeymanID = db.tbPeymanContracts.FirstOrDefault(p => p.FK_UserTarafDovvom == tarafdovom.ID && p.Inactive != true).pec_ID;

                    }


                    var usser = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_Peyman_ID == PeymanID && p.Status == true).FK_User_ID;
                    Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == usser).ToList();
                    

                }


            }


            else
            {
                Model = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKUser == UsrID).ToList();
            }
            return PartialView("_KarkonanDetail2", Model);

        }

        [AuthorizeAAA]
        public ActionResult _UserPeyman(int idPeyman)
        {
            var linkuserpeyman = linkuserPeymanRepo.Update().Where(p => p.FK_Peyman_ID == idPeyman&&p.Status==true).Select(p => p.FK_User_ID).ToList();
            var Model = UserRepo.Update().Where(p => linkuserpeyman.Contains(p.usr_ID)).ToList();

            return PartialView("_UserPeyman", Model);
        }

        public ActionResult _banner()
        {
            //DeactivateBannersForNewYear();
            return PartialView();
        }
     
    }
}