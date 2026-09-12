using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Models.Repositories.Settings;
using SaabWebProject.Utility;
using Telerik.Web.Spreadsheet;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
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
using Microsoft.Ajax.Utilities;
using System.Globalization;
using System.Threading.Tasks;
using Org.BouncyCastle.Bcpg.OpenPgp;

namespace SaabWebProject.Areas.Setting.Controllers
{
    public class AdminController : Controller
    {
        SaabEntities db = new SaabEntities();
        tbVisibleFishRepository VsibileFish;
        public tbCaransSettingRepository tbCaransSettingRepository;
        public AdminController()
        {
            SaabEntities db = new SaabEntities();
            tbCaransSettingRepository = new tbCaransSettingRepository(db);
            VsibileFish = new tbVisibleFishRepository(db);
        }
        // GET: Setting/Admin
        [AuthorizeAAA]
        public ActionResult AdminSetting()
        {
            return View("~/Areas/Setting/Views/Admin/AdminSetting.cshtml");
        }



        /// <summary>
        /// تابع ذخیره سازی تنظیمات کران ها
        /// </summary>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <param name="Criterion"></param>
        /// <param name="standard"></param>
        /// <param name="AcceptLimit"></param>
        /// <param name="select_moalefe"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult SaveCaranSetting(double? min, double? max, double? Criterion, double? standard, double? AcceptLimit, int select_moalefe, int? CaranId, double? AyabOZahab, double? AstaneShoroPadash, double? SaranePadash, double? AstaneShoroJarime, double? SaraneJarime, double? mablagh, bool? agree, bool? zarib_sharestan)
        {
            var caranSettings2 = db.tbCaranSettings.Where(p => p.FK_Moalefe_ID == select_moalefe).FirstOrDefault();
            if (caranSettings2 == null)
            {
                if (CaranId == null)
                {
                    tbCaranSettings caranSettings = new tbCaranSettings();
                    if (min != null)
                    {
                        caranSettings.CaranMin = min;
                    }
                    if (max != null)
                    {
                        caranSettings.CaranMax = max;
                    }
                    if (agree == true)
                    {
                        caranSettings.sarmayeh = true;
                    }
                    else
                    {
                        caranSettings.sarmayeh = false;

                    }
                    if (zarib_sharestan != null && zarib_sharestan == true)
                    {
                        caranSettings.zarib_sharestan = true;
                    }
                    else
                    {
                        caranSettings.zarib_sharestan = false;

                    }
                    if (mablagh != null)
                    {
                        caranSettings.caranmablagh = mablagh;
                    }
                    if (Criterion != null)
                    {
                        caranSettings.CaranCriterion = Criterion;
                    }
                    if (standard != null)
                    {
                        caranSettings.CaranStandard = standard;
                    }
                    if (AcceptLimit != null)
                    {
                        caranSettings.CaranAcceptLimit = AcceptLimit;
                    }
                    if (select_moalefe != 0)
                    {
                        caranSettings.FK_Moalefe_ID = select_moalefe;
                    }
                    if (AyabOZahab != null)
                    {
                        caranSettings.CaranAyabOZahab = AyabOZahab;
                    }
                    if (AstaneShoroPadash != null)
                    {
                        caranSettings.CaranAstaneShoroPadash = AstaneShoroPadash;
                    }
                    if (SaranePadash != null)
                    {
                        caranSettings.CaranSaranePadash = SaranePadash;
                    }
                    if (AstaneShoroJarime != null)
                    {
                        caranSettings.CaranAstaneShoroJarime = AstaneShoroJarime;
                    }
                    if (AstaneShoroPadash != null)
                    {
                        caranSettings.CaranAstaneShoroPadash = AstaneShoroPadash;
                    }
                    if (SaraneJarime != null)
                    {
                        caranSettings.CaranSaraneJarime = SaraneJarime;
                    }
                    tbCaransSettingRepository.Create(caranSettings);
                }
                else
                {
                    tbCaranSettings caranSettings = tbCaransSettingRepository.Find((int)CaranId);
                    if (min != null)
                    {
                        caranSettings.CaranMin = min;
                    }
                    if (max != null)
                    {
                        caranSettings.CaranMax = max;
                    }
                    if (Criterion != null)
                    {
                        caranSettings.CaranCriterion = Criterion;
                    }
                    if (agree == true)
                    {
                        caranSettings.sarmayeh = true;
                    }
                    else
                    {
                        caranSettings.sarmayeh = false;

                    }
                    if (zarib_sharestan != null && zarib_sharestan == true)
                    {
                        caranSettings.zarib_sharestan = true;
                    }
                    else
                    {
                        caranSettings.zarib_sharestan = false;

                    }
                    if (mablagh != null)
                    {
                        caranSettings.caranmablagh = mablagh;
                    }
                    if (standard != null)
                    {
                        caranSettings.CaranStandard = standard;
                    }
                    if (AcceptLimit != null)
                    {
                        caranSettings.CaranAcceptLimit = AcceptLimit;
                    }
                    if (select_moalefe != 0)
                    {
                        caranSettings.FK_Moalefe_ID = select_moalefe;
                    }
                    if (AyabOZahab != null)
                    {
                        caranSettings.CaranAyabOZahab = AyabOZahab;
                    }
                    if (AstaneShoroPadash != null)
                    {
                        caranSettings.CaranAstaneShoroPadash = AstaneShoroPadash;
                    }
                    if (SaranePadash != null)
                    {
                        caranSettings.CaranSaranePadash = SaranePadash;
                    }
                    if (AstaneShoroJarime != null)
                    {
                        caranSettings.CaranAstaneShoroJarime = AstaneShoroJarime;
                    }
                    if (AstaneShoroPadash != null)
                    {
                        caranSettings.CaranAstaneShoroPadash = AstaneShoroPadash;
                    }
                    if (SaraneJarime != null)
                    {
                        caranSettings.CaranSaraneJarime = SaraneJarime;
                    }
                    tbCaransSettingRepository.SaveChanges();
                }
            }
            else
            {
                //tbCaranSettings caranSettings2 = tbCaransSettingRepository.Find((int)CaranId);
                if (min != null)
                {
                    caranSettings2.CaranMin = min;
                }
                if (max != null)
                {
                    caranSettings2.CaranMax = max;
                }
                if (Criterion != null)
                {
                    caranSettings2.CaranCriterion = Criterion;
                }
                if (agree == true)
                {
                    caranSettings2.sarmayeh = true;
                }
                else
                {
                    caranSettings2.sarmayeh = false;

                }
                if (zarib_sharestan != null && zarib_sharestan == true)
                {
                    caranSettings2.zarib_sharestan = true;
                }
                else
                {
                    caranSettings2.zarib_sharestan = false;

                }
                if (mablagh != null)
                {
                    caranSettings2.caranmablagh = mablagh;
                }
                if (standard != null)
                {
                    caranSettings2.CaranStandard = standard;
                }
                if (AcceptLimit != null)
                {
                    caranSettings2.CaranAcceptLimit = AcceptLimit;
                }
                if (select_moalefe != 0)
                {
                    caranSettings2.FK_Moalefe_ID = select_moalefe;
                }
                if (AyabOZahab != null)
                {
                    caranSettings2.CaranAyabOZahab = AyabOZahab;
                }
                if (AstaneShoroPadash != null)
                {
                    caranSettings2.CaranAstaneShoroPadash = AstaneShoroPadash;
                }
                if (SaranePadash != null)
                {
                    caranSettings2.CaranSaranePadash = SaranePadash;
                }
                if (AstaneShoroJarime != null)
                {
                    caranSettings2.CaranAstaneShoroJarime = AstaneShoroJarime;
                }
                if (AstaneShoroPadash != null)
                {
                    caranSettings2.CaranAstaneShoroPadash = AstaneShoroPadash;
                }
                if (SaraneJarime != null)
                {
                    caranSettings2.CaranSaraneJarime = SaraneJarime;
                }
                db.SaveChanges();
            }
            return RedirectToAction("AdminSetting");

        }


        [AuthorizeAAA]
        public ActionResult _CaranList()
        {
            var Model = tbCaransSettingRepository.Update();
            return PartialView(Model);
        }
        public ActionResult modell()
        {
            return View(db.tbCaranSettings.ToList());
        }





        public ActionResult creatcheack(int year_pY = 0, int dore = 0, bool isavailble = true)
        {
            try
            {
                var yy = db.tbCheaksorat.Where(p => p.Year == year_pY && p.number_sorat == dore).FirstOrDefault();
                if (yy != null)
                {
                    yy.Isavailable = isavailble;
                    db.SaveChanges();
                }
                else
                {
                    tbCheaksorat n = new tbCheaksorat();
                    n.Year = year_pY;
                    n.Month = null;
                    n.Fk_pymn = null;
                    n.number_sorat = dore;

                    n.Isavailable = isavailble;

                    // Assuming tbCheaksorat is a DbSet in your DbContext
                    db.tbCheaksorat.Add(n);

                    db.SaveChanges();
                }
                var yy2 = db.tbSoratvaziat.Where(p => p.Year == year_pY && p.number_sorat == dore).ToList();

               foreach(var it in yy2)
                {
                    it.Final_accept = isavailble;
                    db.SaveChanges();
                }
                return Content("True");
            }
            catch (Exception ex)
            {
                // Log the exception or handle it appropriately
                return Content($"{{'success': false, 'error': '{ex.Message}'}}");
            }
        }



        public ActionResult ImportExcel_AddFish(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)

            {

                string Message = GetDataFromExcel_Fish(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public ActionResult ImportExcel_step(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)

            {

                string Message = GetDataFromExcel_step(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public ActionResult ImportExcel_stepcity(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)

            {

                string Message = GetDataFromExcel_citystep(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public ActionResult ImportExcel_Overtimeceiling(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)

            {

                string Message = GetDataFromExcel_Overtimeceiling(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public ActionResult viewforshogl()
        {
            return View();

        }


        public ActionResult viewmax()//maxحقوق کارمندان
        {
            return View();
        }
        public ActionResult Viewmlayatpale()//پله های مالیاتی
        {
            return View();
        }
        public ActionResult ViewmlaSteptaxList()//پله های مالیاتی
        {
            return View(db.tb_Step_tax.ToList());
        }
        public ActionResult City_tax4()//مالیات شهرستان
        {
            return View("~/Areas/Setting/Views/Admin/City_tax.cshtml");
        }


        public ActionResult Overtimeceiling()//مالیات شهرستان
        {
            return View("~/Areas/Setting/Views/Admin/Overtimeceiling.cshtml");
        }

        public ActionResult ExportMoalefeExcelAdamShahr()
        {
            var OutPutFile = SetDataExcel_MoalefeAdamShahr();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "link_persontocity" + extension);
        }
        public Workbook SetDataExcel_MoalefeAdamShahr()
        {
            Workbook Moalefeexcelfile;
            try
            {
                Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/AdamShahr.xlsx"));
            }
            catch (Exception ex)
            {
                throw new Exception("خطا در بارگذاری فایل Excel: " + ex.Message);
            }
            var c = db.tbCities.ToList();
            int rowIndex = 2;
            foreach (var item in c)
            {
                Row row1 = new Row() { Height = 20, Index = rowIndex };

                row1.AddCells(new List<Cell>()
        {
            new Cell
            {
                Value = item.Name,
                FontFamily = "B Nazanin",
                Bold= false ,
                Enable = true,
                 Wrap = false,
                 FontSize = 12,
                 Italic = false,
                 Underline = false,
                 Index = 0
            },
            new Cell
            {
                Value = item.ID,
                FontFamily = "B Nazanin",
                Bold= false ,
                Enable = true,
                 Wrap = false,
                 FontSize = 12,
                 Italic = false,
                 Underline = false,
                 Index = 1
            }


        });
                Moalefeexcelfile.Sheets[0].AddRow(row1);
                rowIndex++;
            }
            int columnIndex1 = 2;
            var users = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            Row row = new Row() { Height = 20, Index = 0 };
            foreach (var user in users)
            {
                row.AddCells(new List<Cell>()
        {
            new Cell
            {
                 Value =user.usr_Personal_ID,
                FontFamily = "B Nazanin",
                Bold= false ,
                Enable = true,
                 Wrap = false,
                 FontSize = 12,
                 Italic = false,
                 Underline = false,
                 Index = columnIndex1

            }
        });
                Moalefeexcelfile.Sheets[0].AddRow(row);
                columnIndex1++;





            }
            var columnIndex2 = 2;
            var row3 = new Row { Height = 20, Index = 1 };
            foreach (var item in users)
            {
                row3.AddCells(new List<Cell>()
        {
            new Cell
            {
                 Value =item.FullName,
                FontFamily = "B Nazanin",
                Bold= false ,
                Enable = true,
                 Wrap = false,
                 FontSize = 12,
                 Italic = false,
                 Underline = false,
                 Index = columnIndex2
            }
        });
                Moalefeexcelfile.Sheets[0].AddRow(row3);
                columnIndex2++;
            }
            return Moalefeexcelfile;
        }




        public async Task< ActionResult> ImportExcel_AdamShahr(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message =await GetDataFromExcel_AdamShahr(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }
        public async Task< ActionResult> ExportMoalefeExcelAdamAdam()
        {
            var OutPutFile =await SetDataExcel_MoalefeAdamAdam();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "link_persontoperson" + extension);
        }
        public async Task<Workbook>  SetDataExcel_MoalefeAdamAdam()
        {
            Workbook Moalefeexcelfile;
            try
            {
                Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/AdamAdam.xlsx"));
            }
            catch (Exception ex)
            {
                throw new Exception("خطا در بارگذاری فایل Excel: " + ex.Message);
            }

            var users =await db.tbUsers.Where(p => p.usr_Personal_ID != null).ToListAsync();

            int rowIndex = 2; // شروع از ردیف 1

            foreach (var user in users)
            {
                Row row = new Row() { Height = 20, Index = rowIndex };

                // اضافه کردن نام کامل کاربر
                row.AddCells(new List<Cell>
{
    new Cell
    {
        Value = user.FullName,
        FontFamily = "B Nazanin",
        Bold= false ,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },
    new Cell
    {
        Value = user.usr_Personal_ID,
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

                // اضافه کردن ردیف به شیت
                Moalefeexcelfile.Sheets[0].AddRow(row);
                rowIndex++;

            }

            int columnIndex1 = 2;
            var row1 = new Row() { Height = 20, Index = 1 };
            foreach (var item in users)
            {
                row1.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value=item.FullName,
                FontFamily="B Nazanin",
                Bold=true,
                Enable=true,
                Wrap=false,
                FontSize=12,
                Italic=false,
                Underline=false,
                Index=columnIndex1
            }

        });
                columnIndex1++;
                Moalefeexcelfile.Sheets[0].AddRow(row1);
            }
            int columnIndex12 = 2;

            var row121 = new Row() { Height = 20, Index = 0 };
            foreach (var item in users)
            {
                row121.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value=item.usr_Personal_ID,
                FontFamily="B Nazanin",
                Bold=true,
                Enable=true,
                Wrap=false,
                FontSize=12,
                Italic=false,
                Underline=false,
                Index=columnIndex12
            }

        });
                columnIndex12++;
                Moalefeexcelfile.Sheets[0].AddRow(row121);
            }
            return Moalefeexcelfile;
        }

        public async Task< ActionResult> ImportExcel_AdamAdam(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message =await GetDataFromExcel_AdamAdam(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }












        public async Task< string> GetDataFromExcel_AdamAdam(HttpPostedFileBase MyExcelStream)
        {
            int name = 0;
            List<tbAdamAdam> AdamAdam = new List<tbAdamAdam>();
            var usr =await db.tbUsers.ToListAsync();
            var tbAdamAdam = await db.tbAdamAdam.ToListAsync();

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
                                var title2 = workbook.Sheets[0].Rows[0].Cells;

                                //var title = workbook.Sheets[0].Rows[0].Cells;
                                //var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                List<int> ID = new List<int>();
                                for (int i = 2; i < title2.Count; i++)
                                {
                                    ID.Add(int.Parse(title2[i].Value.ToString()));
                                }

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                for (int i = 2; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[1];
                                    if (Name.Value != null)
                                    {
                                        name = int.Parse(Name.Value.ToString());
                                        //name = Name.Value; // Convert Name.Value to string
                                        //if (!int.TryParse(Name.Value?.ToString(), out name))
                                        //{
                                        //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        //}

                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 2;

                                    

                                        var find = usr.Where(s => s.usr_Personal_ID == name).FirstOrDefault();
                                    if (name == 4044)
                                    {

                                    }
                                        foreach (var it in ID)

                                        {
                                            var find2 = usr.Where(s => s.usr_Personal_ID == it).FirstOrDefault();
                                        // ابتدا بررسی می‌کنیم که سلول مورد نظر (shomarande) وجود دارد یا خیر.
                                        if (shomarande >= row.Cells.Count || row.Cells[shomarande]?.Value == null)
                                        {
                                            shomarande++;

                                            continue;
                                            // اگر سلول وجود نداشت یا مقدار آن null بود، به آیتم بعدی بروید.
                                        }

                                        // اگر سلول معتبر بود، مقدار آن را استخراج کرده و ادامه پردازش را انجام می‌دهیم.
                                        //int parsedValue = int.Parse(Value.Value.ToString());

                                        // بررسی و اضافه کردن یا به‌روزرسانی مقدار در دیتابیس
                                       

                                        var Value = row.Cells[shomarande];
                                            if (Value.Value != null)
                                            {
                                                var finddd = tbAdamAdam.Where(p => p.Name == find.usr_ID && p.nam == find2.usr_ID).FirstOrDefault();
                                                if (finddd == null)
                                                {
                                                    if (int.Parse(Value.Value.ToString()) != 0)
                                                    {
                                                        tbAdamAdam AdamAdam2 = new tbAdamAdam();
                                                        AdamAdam2.nam = find2.usr_ID;
                                                        AdamAdam2.Name = find.usr_ID;

                                                        AdamAdam2.Value = int.Parse(Value.Value.ToString());
                                                        AdamAdam.Add(AdamAdam2);

                                                    }
                                                }
                                                else
                                                {
                                                    finddd.Value = int.Parse(Value.Value.ToString());
                                                    await db.SaveChangesAsync();
                                                }
                                               
                                            
                                            }
                                            //else
                                            //{
                                            //    tbAdamAdam AdamAdam2 = new tbAdamAdam();
                                            //    AdamAdam2.nam = find2.usr_ID;
                                            //    AdamAdam2.Name = find.usr_ID;

                                            //    AdamAdam2.Value = 0;
                                            //    AdamAdam.Add(AdamAdam2);


                                            //}
                                            shomarande++;
                                        }
                                    
                                    //else
                                    //{
                                    //    return "ایندکس خارج از محدوده در سطر " + (i + 1);
                                    //}


                                }
                                db.tbAdamAdam.AddRange(AdamAdam);
                              await  db.SaveChangesAsync();
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

        public async Task<ActionResult> viewHosdar()
        {
            return View();
        }

        public async Task< string> GetDataFromExcel_AdamShahr(HttpPostedFileBase MyExcelStream)
        {
            var tbUsers=await db.tbUsers.ToListAsync();
            var tbCities = await db.tbCities.ToListAsync();
            var tbAdamShahr = await db.tbAdamShahr.ToListAsync();

            int shahr = 0;
            List<tbAdamShahr> AdamShahr = new List<tbAdamShahr>();

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
                                var title2 = workbook.Sheets[0].Rows[0].Cells;

                                //var title = workbook.Sheets[0].Rows[0].Cells;
                                //var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                List<int> ID = new List<int>();
                                for (int i = 2; i < title2.Count; i++)
                                {
                                    ID.Add(int.Parse(title2[i].Value.ToString()));
                                }

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                for (int i = 2; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Shahr = row.Cells[1];
                                    if (Shahr.Value != null)
                                    {
                                        shahr = int.Parse(Shahr.Value.ToString());
                                        //name = Name.Value; // Convert Name.Value to string
                                        //if (!int.TryParse(Name.Value?.ToString(), out name))
                                        //{
                                        //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        //}

                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 2;

                                    if (shomarande < row.Cells.Count)
                                    {

                                        var find = tbCities.Where(s => s.ID == shahr).FirstOrDefault();
                                        foreach (var it in ID)

                                        {
                                            var find2 = tbUsers.Where(s => s.usr_Personal_ID == it).FirstOrDefault();

                                            var Value = row.Cells[shomarande];
                                            if (Value.Value != null)
                                            {
                                                var finddd = tbAdamShahr.Where(p => p.idshahr == find.ID && p.idkarbar == find2.usr_ID).FirstOrDefault();
                                                if (finddd == null)
                                                {
                                                    if (int.Parse(Value.Value.ToString()) != 0)
                                                    {

                                                        tbAdamShahr tbAdamShahr1 = new tbAdamShahr();
                                                        tbAdamShahr1.idshahr = find.ID;
                                                        tbAdamShahr1.idkarbar = find2.usr_ID;

                                                        tbAdamShahr1.value = int.Parse(Value.Value.ToString());
                                                        AdamShahr.Add(tbAdamShahr1);
                                                    }
                                                }
                                                else
                                                {
                                                    finddd.value= int.Parse(Value.Value.ToString());
                                                    await db.SaveChangesAsync();
                                                }

                                                

                                            }
                                            //else
                                            //{
                                            //    tbAdamShahr tbAdamShahr1 = new tbAdamShahr();
                                            //    tbAdamShahr1.idshahr = find.ID;
                                            //    tbAdamShahr1.idkarbar = find2.usr_ID;

                                            //    tbAdamShahr1.value = 0;
                                            //    AdamShahr.Add(tbAdamShahr1);
                                            //}
                                            shomarande++;
                                        }
                                    }
                                    else
                                    {
                                        return "ایندکس خارج از محدوده در سطر " + (i + 1);
                                    }


                                }
                                db.tbAdamShahr.AddRange(AdamShahr);
                               await db.SaveChangesAsync();
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




        public ActionResult City_taxList(){//مالیات شهرستان
            return View(db.tbCities.Where(p=>p.Zaribmalyat!=null).ToList());
        }
        public ActionResult City_overtimeist()
        {//مالیات شهرستان
            return View(db.tbCities.Where(p => p.Max_ezafeh != null&&p.Ayab_zohab!=null).ToList());
        }
        public ActionResult MAX_KarKard()//maxکارکرد افراد
        {
            return View();
        }
        public ActionResult MAX_KarKarxList()
        {//مالیات شهرستان
            return View(db.tbMaxkarkardMonth.ToList());
        }
        public ActionResult viewforcheackfish()
        {
            return View();
        }
        public ActionResult shahrviewforcheackfish()
        {
            return View();
        }
        public ActionResult personviewforcheackfish()
        {
            var find = db.tbAdamAdam.Where(s => s.Value == 1&&s.nam!=null&&s.Name!=null).ToList();
            return View(find);
        }
        public async Task<ActionResult> ImportExcel_Cheakfishusr(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = await GetDataFromExcel_Fishcheak(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public ActionResult Listcheackfish()
        {
            return View(db.tbcheackfishforusers.Where(s=>s.HasFish==true).ToList());
        }
        public async Task<string> GetDataFromExcel_Fishcheak(HttpPostedFileBase MyExcelStream)
        {
            int CaranMin, countDays, sanavatt, olad, mask, shoghl, ghar, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID;
            string name = ""; string timestart = ""; string timeend = ""; string CaranAyabOZahab = ""; string month4 = ""; string month2 = ""; string year2 = ""; string year22 = ""; string day22 = "";
            string oladd, maskan, gharbar, mozdd, sanavattt;
            var usr = await db.tbUsers.ToListAsync();
            var moalfe = await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<string> lstMoalefe = new List<string>();
            List<tbcheackfishforusers> ISMAX = new List<tbcheackfishforusers>();
            var MAXHOghogh = db.tbcheackfishforusers.ToList();
            List<tbCaranSettings> IsFish = new List<tbCaranSettings>();
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
                                //for (int i = 16; i < title.Count; i++)
                                //{
                                //    lstMoalefe.Add(title[i].Value.ToString());
                                //}

                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[1];

                                    if (int.TryParse(Name.Value.ToString(), out int parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }

                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[2];//نوع قرارداد
                                    if (int.TryParse(year.Value.ToString(), out int result))
                                    {
                                        CaranMax = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[3];//کارفرما
                                    if (int.TryParse(month.Value.ToString(), out int result2))
                                    {
                                        CaranMin = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    //var CaranAcceptLimit2 = row.Cells[4];//کارفرما اصلی
                                    //if (int.TryParse(CaranAcceptLimit2.Value.ToString(), out int result3))
                                    //{
                                    //    CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    var CaranAcceptLimit23 = row.Cells[4];
                                    bool cheak = false;


                                   //کارفرما اصلی
                                    if (int.TryParse(CaranAcceptLimit23.Value.ToString(), out int result33))
                                    {
                                        CaranCriterion = result33; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        CaranCriterion = 0;
                                    }
                                    if (CaranCriterion != 0)
                                    {
                                        cheak = true;
                                    }
                                    var tt = usr.FirstOrDefault(p => p.usr_Personal_ID == countDays);
                                    if (MAXHOghogh.Count != 0)
                                    {
                                        var trt = db.tbcheackfishforusers.Where(p => p.FK_User == tt.usr_ID && p.Month == CaranMin && p.Year == CaranMax).FirstOrDefault();
                                        if (trt != null)
                                        {
                                            trt.Year = CaranMax;
                                            trt.HasFish = cheak;
                                            trt.Month = CaranMin;
                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tbcheackfishforusers obj4 = new tbcheackfishforusers();

                                            // اختصاص تاریخ به usc_StartTime

                                            obj4.Year = CaranMax;
                                            obj4.HasFish = cheak;
                                            obj4.FK_User = tt.usr_ID;
                                            obj4.Month = CaranMin;
                                            ISMAX.Add(obj4);

                                        }
                                    }
                                    else
                                    {
                                        tbcheackfishforusers obj4 = new tbcheackfishforusers();

                                        // اختصاص تاریخ به usc_StartTime

                                        obj4.Year = CaranMax;
                                        obj4.HasFish = cheak;
                                        obj4.FK_User = tt.usr_ID;
                                        obj4.Month = CaranMin;
                                        ISMAX.Add(obj4);
                                    }










                                    //if (DateTime.TryParseExact(inputDate, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    //{
                                    //    usc_StartTime = parsedDate;
                                    //}
                                    //else
                                    //{
                                    //    // در صورتی که تبدیل موفقیت‌آمیز نباشد، می‌توانید یک پیام خطا نمایش دهید.
                                    //    return"تاریخ معتبر نیست";
                                    //}

                                    // حالا متغیر تاریخ را به مقدار usc_StartTime اختصاص دهید.
                                    //obj4.usc_EndTime = p.ToDateTime(CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, 0, 0, 0, 0).Date;
                                    //  tbUserContracts.obj_tbuserContracts.usc_EndTime = p.ToDateTime(obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_year, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_month, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_day, 0, 0, 0, 0).Date;


                                    //var c = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(name)).FirstOrDefault();
                                    //var tttt = db.tbCaranSettings.Where(p => p.FK_Moalefe_ID == c.md_ID).FirstOrDefault();
                                    //if (tttt == null)
                                    //{
                                    //    if (sarmah == 1)
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = true,
                                    //            caranmablagh = mablgh,

                                    //        }; IsFish.Add(obj);
                                    //    }
                                    //    else
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = false,
                                    //            caranmablagh = mablgh,


                                    //        }; IsFish.Add(obj);
                                    //    }

                                    //}
                                    //else
                                    //{
                                    //    if (sarmah == 1)
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = true;
                                    //        tttt.caranmablagh = mablgh;

                                    //        db.SaveChanges();
                                    //    }
                                    //    else
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = false;
                                    //        tttt.caranmablagh = mablgh;
                                    //        db.SaveChanges();

                                    //    }
                                    //}






                                }
                                db.tbcheackfishforusers.AddRange(ISMAX);
                                await db.SaveChangesAsync();
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
                    if (IsFish.Count != 0)
                    {
                        //if (!tbCaransSettingRepository.AddRange(IsFish))
                        //{
                        //    transaction.Rollback();
                        //    return "ثبت کردن مولفه با خطا مواجه شد";
                        //}
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

        public ActionResult Excel_fishcheak()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = Fishch();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Fishcheackhas " + extension);

        }

        private Workbook Fishch()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/fishcheakusr.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            //var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            //var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            var city = db.tbUsers.Where(p=>p.usr_Personal_ID!=null).ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in city)
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            return Moalefeexcelfile;








        }





        public ActionResult viewmaxhoghogh()
        {
            return View(db.tbMaxPayForMonth.ToList());
        }
        public async Task<ActionResult> ImportExcel_MAXHOGHOGH(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = await GetDataFromExcel_MAXHOGHOGH(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public async Task<ActionResult> ImportExcel_MAXKARKARD(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = await GetDataFromExcel_MAXKARKARD(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        public static DateTime PersianDateToDateTime(string persianDate)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            string[] parts = persianDate.Split('/');
            int year = int.Parse(parts[2]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[0]);
            return persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        }

        public async Task<string> GetDataFromExcel_MAXHOGHOGH(HttpPostedFileBase MyExcelStream)
        {
            int CaranMin, countDays, sanavatt, olad, mask, shoghl, ghar, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID;
            string name = ""; string timestart = ""; string timeend = ""; string CaranAyabOZahab = ""; string month4 = ""; string month2 = ""; string year2 = ""; string year22 = ""; string day22 = "";
            string oladd, maskan, gharbar, mozdd, sanavattt;
            var usr = await db.tbUsers.ToListAsync();
            var moalfe = await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<string> lstMoalefe = new List<string>();
            List<tbMaxPayForMonth> ISMAX = new List<tbMaxPayForMonth>();
            var MAXHOghogh=db.tbMaxPayForMonth.ToList();
            List<tbCaranSettings> IsFish = new List<tbCaranSettings>();
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
                                //for (int i = 16; i < title.Count; i++)
                                //{
                                //    lstMoalefe.Add(title[i].Value.ToString());
                                //}

                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[1];

                                    if (int.TryParse(Name.Value.ToString(), out int parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }

                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[2];//نوع قرارداد
                                    if (int.TryParse(year.Value.ToString(), out int result))
                                    {
                                        CaranMax = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[3];//کارفرما
                                    if (int.TryParse(month.Value.ToString(), out int result2))
                                    {
                                        CaranMin = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    var CaranAcceptLimit2 = row.Cells[4];//کارفرما اصلی
                                    if (int.TryParse(CaranAcceptLimit2.Value.ToString(), out int result3))
                                    {
                                        CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var CaranAcceptLimit23 = row.Cells[5];//کارفرما اصلی
                                    if (int.TryParse(CaranAcceptLimit23.Value.ToString(), out int result33))
                                    {
                                        CaranCriterion = result33; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        CaranCriterion = 0;
                                    }
                                    var tt = usr.FirstOrDefault(p => p.usr_Personal_ID == countDays);
                                    if (MAXHOghogh.Count != 0)
                                    {
                                        var trt = db.tbMaxPayForMonth.Where(p => p.FKUser == tt.usr_ID && p.Month == CaranMin && p.Year == CaranAcceptLimit).FirstOrDefault();
                                        if (trt != null)
                                        {
                                            trt.Year = CaranAcceptLimit;
                                            trt.RemainValue = CaranCriterion;
                                            trt.Value = CaranMax;
                                            trt.Month = CaranMin;
                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tbMaxPayForMonth obj4 = new tbMaxPayForMonth();

                                            // اختصاص تاریخ به usc_StartTime

                                            obj4.Year = CaranAcceptLimit;
                                            obj4.RemainValue = CaranCriterion;
                                            obj4.FKUser = tt.usr_ID;
                                            obj4.Value = CaranMax;
                                            obj4.Month = CaranMin;
                                            ISMAX.Add(obj4);

                                        }
                                    }
                                    else
                                    {
                                        tbMaxPayForMonth obj4 = new tbMaxPayForMonth();

                                        // اختصاص تاریخ به usc_StartTime

                                        obj4.Year = CaranAcceptLimit;
                                        obj4.RemainValue = CaranCriterion;
                                        obj4.FKUser = tt.usr_ID;
                                        obj4.Value = CaranMax;
                                        obj4.Month = CaranMin;
                                        ISMAX.Add(obj4);
                                    }







                                   
                                 

                                    //if (DateTime.TryParseExact(inputDate, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    //{
                                    //    usc_StartTime = parsedDate;
                                    //}
                                    //else
                                    //{
                                    //    // در صورتی که تبدیل موفقیت‌آمیز نباشد، می‌توانید یک پیام خطا نمایش دهید.
                                    //    return"تاریخ معتبر نیست";
                                    //}

                                    // حالا متغیر تاریخ را به مقدار usc_StartTime اختصاص دهید.
                                    //obj4.usc_EndTime = p.ToDateTime(CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, 0, 0, 0, 0).Date;
                                    //  tbUserContracts.obj_tbuserContracts.usc_EndTime = p.ToDateTime(obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_year, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_month, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_day, 0, 0, 0, 0).Date;


                                    //var c = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(name)).FirstOrDefault();
                                    //var tttt = db.tbCaranSettings.Where(p => p.FK_Moalefe_ID == c.md_ID).FirstOrDefault();
                                    //if (tttt == null)
                                    //{
                                    //    if (sarmah == 1)
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = true,
                                    //            caranmablagh = mablgh,

                                    //        }; IsFish.Add(obj);
                                    //    }
                                    //    else
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = false,
                                    //            caranmablagh = mablgh,


                                    //        }; IsFish.Add(obj);
                                    //    }

                                    //}
                                    //else
                                    //{
                                    //    if (sarmah == 1)
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = true;
                                    //        tttt.caranmablagh = mablgh;

                                    //        db.SaveChanges();
                                    //    }
                                    //    else
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = false;
                                    //        tttt.caranmablagh = mablgh;
                                    //        db.SaveChanges();

                                    //    }
                                    //}






                                }
                                db.tbMaxPayForMonth.AddRange(ISMAX);
                                await db.SaveChangesAsync();
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
                    if (IsFish.Count != 0)
                    {
                        //if (!tbCaransSettingRepository.AddRange(IsFish))
                        //{
                        //    transaction.Rollback();
                        //    return "ثبت کردن مولفه با خطا مواجه شد";
                        //}
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

        public async Task<string> GetDataFromExcel_MAXKARKARD(HttpPostedFileBase MyExcelStream)
        {
            float CaranCriterion , ayabzoh;
            int CaranMin, countDays, sanavatt, olad, mask, shoghl, ghar, CaranMax, CaranAcceptLimit, CaranStandard, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID;
            string name = ""; string timestart = ""; string timeend = ""; string CaranAyabOZahab = ""; string month4 = ""; string month2 = ""; string year2 = ""; string year22 = ""; string day22 = "";
            string oladd, maskan, gharbar, mozdd, sanavattt;
            var usr = await db.tbUsers.ToListAsync();
            var moalfe = await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<string> lstMoalefe = new List<string>();
            List<tbMaxkarkardMonth> ISMAX = new List<tbMaxkarkardMonth>();
            var MAXHOghogh = db.tbMaxkarkardMonth.ToList();
            List<tbCaranSettings> IsFish = new List<tbCaranSettings>();
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
                                //for (int i = 16; i < title.Count; i++)
                                //{
                                //    lstMoalefe.Add(title[i].Value.ToString());
                                //}

                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[1];

                                    if (int.TryParse(Name.Value.ToString(), out int parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }

                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[2];//نوع قرارداد
                                    if (int.TryParse(year.Value.ToString(), out int result))
                                    {
                                        CaranMax = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[3];//کارفرما
                                    if (int.TryParse(month.Value.ToString(), out int result2))
                                    {
                                        CaranMin = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    var CaranAcceptLimit2 = row.Cells[4];//کارفرما اصلی
                                    if (int.TryParse(CaranAcceptLimit2.Value.ToString(), out int result3))
                                    {
                                        CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var CaranAcceptLimit23 = row.Cells[5];//کارفرما اصلی
                                    if (float.TryParse(CaranAcceptLimit23.Value.ToString(), out float result33))
                                    {
                                        CaranCriterion = result33; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        CaranCriterion = 0;
                                    }
                                    var ayab = row.Cells[6];//کارفرما اصلی
                                    if (float.TryParse(ayab.Value.ToString(), out float ayab2))
                                    {
                                        ayabzoh = ayab2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        ayabzoh = 0;
                                    }
                                    var tt = usr.FirstOrDefault(p => p.usr_Personal_ID == countDays);
                                    if (MAXHOghogh.Count != 0)
                                    {
                                        var trt = db.tbMaxkarkardMonth.Where(p => p.FKUser == tt.usr_ID && p.Month == CaranMin && p.Year == CaranAcceptLimit).FirstOrDefault();
                                        if (trt != null)
                                        {
                                            trt.Year = CaranAcceptLimit;
                                            trt.RemainValue = CaranCriterion;
                                            trt.Value = CaranMax;
                                            trt.Month = CaranMin;
                                            trt.valuemaxayab = ayabzoh;

                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tbMaxkarkardMonth obj4 = new tbMaxkarkardMonth();

                                            // اختصاص تاریخ به usc_StartTime

                                            obj4.Year = CaranAcceptLimit;
                                            obj4.RemainValue = CaranCriterion;
                                            obj4.FKUser = tt.usr_ID;
                                            obj4.Value = CaranMax;
                                            obj4.Month = CaranMin;
                                            obj4.valuemaxayab = ayabzoh;

                                            ISMAX.Add(obj4);

                                        }
                                    }
                                    else
                                    {
                                        tbMaxkarkardMonth obj4 = new tbMaxkarkardMonth();

                                        // اختصاص تاریخ به usc_StartTime

                                        obj4.Year = CaranAcceptLimit;
                                        obj4.RemainValue = CaranCriterion;
                                        obj4.FKUser = tt.usr_ID;
                                        obj4.Value = CaranMax;
                                        obj4.Month = CaranMin;
                                        obj4.valuemaxayab = ayabzoh;

                                        ISMAX.Add(obj4);
                                    }










                                    //if (DateTime.TryParseExact(inputDate, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    //{
                                    //    usc_StartTime = parsedDate;
                                    //}
                                    //else
                                    //{
                                    //    // در صورتی که تبدیل موفقیت‌آمیز نباشد، می‌توانید یک پیام خطا نمایش دهید.
                                    //    return"تاریخ معتبر نیست";
                                    //}

                                    // حالا متغیر تاریخ را به مقدار usc_StartTime اختصاص دهید.
                                    //obj4.usc_EndTime = p.ToDateTime(CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, 0, 0, 0, 0).Date;
                                    //  tbUserContracts.obj_tbuserContracts.usc_EndTime = p.ToDateTime(obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_year, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_month, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_day, 0, 0, 0, 0).Date;


                                    //var c = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(name)).FirstOrDefault();
                                    //var tttt = db.tbCaranSettings.Where(p => p.FK_Moalefe_ID == c.md_ID).FirstOrDefault();
                                    //if (tttt == null)
                                    //{
                                    //    if (sarmah == 1)
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = true,
                                    //            caranmablagh = mablgh,

                                    //        }; IsFish.Add(obj);
                                    //    }
                                    //    else
                                    //    {
                                    //        tbCaranSettings obj = new tbCaranSettings
                                    //        {
                                    //            CaranMax = (float)CaranMax,
                                    //            CaranSaranePadash = (float)CaranSaranePadash,
                                    //            CaranMin = (float)CaranMin,
                                    //            CaranAcceptLimit = (float)CaranAcceptLimit,
                                    //            CaranCriterion = (float)CaranCriterion,
                                    //            CaranStandard = (float)CaranStandard,
                                    //            CaranAyabOZahab = (float)CaranAyabOZahab,
                                    //            CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                    //            CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                    //            CaranSaraneJarime = (float)CaranSaraneJarime,
                                    //            FK_Moalefe_ID = c.md_ID,
                                    //            sarmayeh = false,
                                    //            caranmablagh = mablgh,


                                    //        }; IsFish.Add(obj);
                                    //    }

                                    //}
                                    //else
                                    //{
                                    //    if (sarmah == 1)
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = true;
                                    //        tttt.caranmablagh = mablgh;

                                    //        db.SaveChanges();
                                    //    }
                                    //    else
                                    //    {

                                    //        tttt.CaranMax = (float)CaranMax;
                                    //        tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                    //        tttt.CaranMin = (float)CaranMin;
                                    //        tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                    //        tttt.CaranCriterion = (float)CaranCriterion;
                                    //        tttt.CaranStandard = (float)CaranStandard;
                                    //        tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                    //        tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                    //        tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                    //        tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                    //        tttt.FK_Moalefe_ID = c.md_ID;
                                    //        tttt.sarmayeh = false;
                                    //        tttt.caranmablagh = mablgh;
                                    //        db.SaveChanges();

                                    //    }
                                    //}






                                }
                                db.tbMaxkarkardMonth.AddRange(ISMAX);
                                await db.SaveChangesAsync();
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
                    if (IsFish.Count != 0)
                    {
                        //if (!tbCaransSettingRepository.AddRange(IsFish))
                        //{
                        //    transaction.Rollback();
                        //    return "ثبت کردن مولفه با خطا مواجه شد";
                        //}
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


        public ActionResult Step_tax()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = Step_taxExcel();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Step_tax " + extension);

        }

        private Workbook Step_taxExcel()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EXcel_PALEMALYAT.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID &&p.Status==true&& p.tbUsers.usr_Personal_ID != null).ToList();
            //var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            //var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            //var city = db.tbCities.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            //foreach (var item in city)
            //{
            //    Row = new Row() { Height = 20, Index = count1 };
            //    Row.AddCells(new List<Cell>()
            //    {
            //        new Cell()
            //        {
            //            Value = item.Name,
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = 0
            //        },
                        
            //    });
            //    count1++;
            //    Moalefeexcelfile.Sheets[0].AddRow(Row);
            //}

            return Moalefeexcelfile;








        }
        public ActionResult city_tax()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = city_taxExcel();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Step_tax " + extension);

        }


        public ActionResult city_overtime()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = city_overtimexcel();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "overtimecelling " + extension);

        }


        private Workbook city_taxExcel()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/city_tax.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            //var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            //var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            var city = db.tbCities.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in city)
            {
                Row = new Row() { Height = 20, Index = count1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Name,
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
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }

            return Moalefeexcelfile;








        }

        private Workbook city_overtimexcel()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/overtime.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            //var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            //var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            var city = db.tbCities.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in city)
            {
                Row = new Row() { Height = 20, Index = count1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Name,
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
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }

            return Moalefeexcelfile;








        }
        public string GetDataFromExcel_citystep(HttpPostedFileBase MyExcelStream)
        {
            float CaranMin, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAyabOZahab, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID, number;
            string name = "";
            List<tb_Step_tax> step = new List<tb_Step_tax>();
            List<tbCities> city = new List<tbCities>();
            var citu = db.tbCities.ToList();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            //var steptax = db.tb_Step_tax.ToList();
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
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null)
                                    {
                                        name = (string)Name.Value; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[1];
                                    if (float.TryParse(year.Value.ToString(), out float result))
                                    {
                                        CaranMin = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                               

                                    //var CaranStandard2 = row.Cells[5];
                                    //if (int.TryParse(CaranStandard2.Value.ToString(), out int CaranStandard22))
                                    //{
                                    //    CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    if (citu.Count != 0)
                                    {
                                        var t = citu.Where(p => p.Name == name ).FirstOrDefault();
                                        if (t != null)
                                        {
                                            t.Zaribmalyat = CaranMin;
                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tbCities cit = new tbCities();
                                          
                                            cit.Zaribmalyat= CaranMin;
                                            cit.Name= name;
                                            city.Add(cit);

                                        }
                                    }
                                    else
                                    {
                                        tbCities cit = new tbCities();

                                        cit.Zaribmalyat = CaranMin;
                                        cit.Name = name;
                                        city.Add(cit);


                                    }



                                }
                                db.tbCities.AddRange(city);
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
                    //if (IsFish.Count != 0)
                    //{
                    //    if (!tbCaransSettingRepository.AddRange(IsFish))
                    //    {
                    //        transaction.Rollback();
                    //        return "ثبت کردن مولفه با خطا مواجه شد";
                    //    }
                    //}

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
        public string GetDataFromExcel_Overtimeceiling(HttpPostedFileBase MyExcelStream)
        {
            float CaranMin, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAyabOZahab, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID, number;
            string name = "";
            List<tb_Step_tax> step = new List<tb_Step_tax>();
            List<tbCities> city = new List<tbCities>();
            var citu = db.tbCities.ToList();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            //var steptax = db.tb_Step_tax.ToList();
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
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null)
                                    {
                                        name = (string)Name.Value; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[1];
                                    if (float.TryParse(year.Value.ToString(), out float result))
                                    {
                                        CaranMin = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year333 = row.Cells[2];
                                    if (float.TryParse(year333.Value.ToString(), out float result22))
                                    {
                                        CaranStandard = result22;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    //var CaranStandard2 = row.Cells[5];
                                    //if (int.TryParse(CaranStandard2.Value.ToString(), out int CaranStandard22))
                                    //{
                                    //    CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    if (citu.Count != 0)
                                    {
                                        var t = citu.Where(p => p.Name == name).FirstOrDefault();
                                        if (t != null)
                                        {
                                            t.Max_ezafeh = CaranMin;
                                            t.Ayab_zohab = CaranStandard;
                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tbCities cit = new tbCities();

                                            cit.Max_ezafeh = CaranMin;
                                            cit.Name = name;
                                            cit.Ayab_zohab = CaranStandard;

                                            city.Add(cit);

                                        }
                                    }
                                    else
                                    {
                                        tbCities cit = new tbCities();

                                        cit.Max_ezafeh = CaranMin;
                                        cit.Ayab_zohab = CaranStandard;

                                        cit.Name = name;
                                        city.Add(cit);


                                    }



                                }
                                db.tbCities.AddRange(city);
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
                    //if (IsFish.Count != 0)
                    //{
                    //    if (!tbCaransSettingRepository.AddRange(IsFish))
                    //    {
                    //        transaction.Rollback();
                    //        return "ثبت کردن مولفه با خطا مواجه شد";
                    //    }
                    //}

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


        public ActionResult EXCEL_MAXKARKARD()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMAXKARKARD();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "MAXKARKARD" + extension);

        }
        public ActionResult EXCEL_MAXHOGHOGH()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel4();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "MAXHOGHOGH" + extension);

        }
        private Workbook SetDataToMachinsOrToolsExcel4()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Maxhoghog.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            var company = db.tbCompanies.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in usr)
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }
            
            return Moalefeexcelfile;








        }
        private Workbook SetDataToMAXKARKARD()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/MAXKARKARD.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var usr = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            var company = db.tbCompanies.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in usr)
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
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            return Moalefeexcelfile;








        }

        public string GetDataFromExcel_Fish(HttpPostedFileBase MyExcelStream)
        {
            float CaranMin, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAyabOZahab, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime,mablgh,sarmah, zarib_sharestan;
            int FK_Moalefe_ID;
            string name = "";
            int id = 0;
            int zarib_sharestan1 = 0;

            List<tbCaranSettings> IsFish = new List<tbCaranSettings>();
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
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (int.TryParse(Name.Value.ToString(), out int result1233))
                                    {
                                        id = result1233; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[3];
                                    if (float.TryParse(year.Value.ToString(), out float result))
                                    {
                                        CaranMin = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[4];
                                    if (float.TryParse(month.Value.ToString(), out float result2))
                                    {
                                        CaranMax = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    var CaranAcceptLimit2 = row.Cells[5];
                                    if (float.TryParse(CaranAcceptLimit2.Value.ToString(), out float result3))
                                    {
                                        CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranCriterion2 = row.Cells[6];
                                    if (float.TryParse(CaranCriterion2.Value.ToString(), out float CaranCriterion22))
                                    {
                                        CaranCriterion = CaranCriterion22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var CaranStandard2 = row.Cells[7];
                                    if (float.TryParse(CaranStandard2.Value.ToString(), out float CaranStandard22))
                                    {
                                        CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranAyabOZahab2 = row.Cells[8];
                                    if (float.TryParse(CaranAyabOZahab2.Value.ToString(), out float CaranAyabOZahab22))
                                    {
                                        CaranAyabOZahab = CaranAyabOZahab22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranAstaneShoroPadash2 = row.Cells[9];
                                    if (float.TryParse(CaranAstaneShoroPadash2.Value.ToString(), out float CaranAstaneShoroPadash22))
                                    {
                                        CaranAstaneShoroPadash = CaranAstaneShoroPadash22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranSaranePadash2 = row.Cells[10];
                                    if (float.TryParse(CaranSaranePadash2.Value.ToString(), out float result6))
                                    {
                                        CaranSaranePadash = result6; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranAstaneShoroJarime2 = row.Cells[11];
                                    if (float.TryParse(CaranAstaneShoroJarime2.Value.ToString(), out float CaranAstaneShoroJarime222))
                                    {
                                        CaranAstaneShoroJarime = CaranAstaneShoroJarime222; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var CaranSaraneJarime2 = row.Cells[12];
                                    if (float.TryParse(CaranSaraneJarime2.Value.ToString(), out float CaranSaraneJarime22))
                                    {
                                        CaranSaraneJarime = CaranSaraneJarime22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var Caranmablagh = row.Cells[13];
                                    if (float.TryParse(Caranmablagh.Value.ToString(), out float CaranSaraneJarime222))
                                    {
                                        mablgh = CaranSaraneJarime222; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var sarma = row.Cells[14];
                                    if (float.TryParse(sarma.Value.ToString(), out float CaranSaraneJarime2222))
                                    {
                                        sarmah = CaranSaraneJarime2222; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var zarib_sharestan11 = row.Cells[15];
                                    if (float.TryParse(zarib_sharestan11.Value.ToString(), out float zarib_sharestan12))
                                    {
                                        zarib_sharestan = zarib_sharestan12; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var c = db.tbContractMoalefeDastmozdi.Where(p => p.md_ID== id).FirstOrDefault();
                                    var tttt = db.tbCaranSettings.Where(p => p.FK_Moalefe_ID == c.md_ID).FirstOrDefault();
                                    if (tttt == null)
                                    {
                                        if (sarmah == 1)
                                        {
                                            if (zarib_sharestan == 1)
                                            {
                                                tbCaranSettings obj1 = new tbCaranSettings
                                                {
                                                    CaranMax = (float)CaranMax,
                                                    CaranSaranePadash = (float)CaranSaranePadash,
                                                    CaranMin = (float)CaranMin,
                                                    CaranAcceptLimit = (float)CaranAcceptLimit,
                                                    CaranCriterion = (float)CaranCriterion,
                                                    CaranStandard = (float)CaranStandard,
                                                    CaranAyabOZahab = (float)CaranAyabOZahab,
                                                    CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                                    CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                                    CaranSaraneJarime = (float)CaranSaraneJarime,
                                                    FK_Moalefe_ID = c.md_ID,
                                                    sarmayeh = true,
                                                    caranmablagh = mablgh,
                                                    zarib_sharestan = true

                                                }; IsFish.Add(obj1);
                                            }
                                            else
                                            {
                                                tbCaranSettings obj = new tbCaranSettings
                                                {
                                                    CaranMax = (float)CaranMax,
                                                    CaranSaranePadash = (float)CaranSaranePadash,
                                                    CaranMin = (float)CaranMin,
                                                    CaranAcceptLimit = (float)CaranAcceptLimit,
                                                    CaranCriterion = (float)CaranCriterion,
                                                    CaranStandard = (float)CaranStandard,
                                                    CaranAyabOZahab = (float)CaranAyabOZahab,
                                                    CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                                    CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                                    CaranSaraneJarime = (float)CaranSaraneJarime,
                                                    FK_Moalefe_ID = c.md_ID,
                                                    sarmayeh = true,
                                                    caranmablagh = mablgh,
                                                    zarib_sharestan = false


                                                }; IsFish.Add(obj);
                                            }
                                            
                                        }
                                        else
                                        {
                                            if (zarib_sharestan == 1)
                                            {
                                                tbCaranSettings obj1 = new tbCaranSettings
                                                {
                                                    CaranMax = (float)CaranMax,
                                                    CaranSaranePadash = (float)CaranSaranePadash,
                                                    CaranMin = (float)CaranMin,
                                                    CaranAcceptLimit = (float)CaranAcceptLimit,
                                                    CaranCriterion = (float)CaranCriterion,
                                                    CaranStandard = (float)CaranStandard,
                                                    CaranAyabOZahab = (float)CaranAyabOZahab,
                                                    CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                                    CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                                    CaranSaraneJarime = (float)CaranSaraneJarime,
                                                    FK_Moalefe_ID = c.md_ID,
                                                    sarmayeh = false,
                                                    caranmablagh = mablgh,
                                                    zarib_sharestan = true

                                                }; IsFish.Add(obj1);
                                            }
                                            else
                                            {
                                                tbCaranSettings obj = new tbCaranSettings
                                                {
                                                    CaranMax = (float)CaranMax,
                                                    CaranSaranePadash = (float)CaranSaranePadash,
                                                    CaranMin = (float)CaranMin,
                                                    CaranAcceptLimit = (float)CaranAcceptLimit,
                                                    CaranCriterion = (float)CaranCriterion,
                                                    CaranStandard = (float)CaranStandard,
                                                    CaranAyabOZahab = (float)CaranAyabOZahab,
                                                    CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash,
                                                    CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime,
                                                    CaranSaraneJarime = (float)CaranSaraneJarime,
                                                    FK_Moalefe_ID = c.md_ID,
                                                    sarmayeh = false,
                                                    caranmablagh = mablgh,
                                                    zarib_sharestan = false

                                                }; IsFish.Add(obj);
                                            }
                                         
                                        }

                                    }
                                    else
                                    {
                                        if (sarmah == 1)
                                        {

                                            tttt.CaranMax = (float)CaranMax;
                                            tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                            tttt.CaranMin = (float)CaranMin;
                                            tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                            tttt.CaranCriterion = (float)CaranCriterion;
                                            tttt.CaranStandard = (float)CaranStandard;
                                            tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                            tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                            tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                            tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                            tttt.FK_Moalefe_ID = c.md_ID;
                                            tttt.sarmayeh = true;
                                            tttt.caranmablagh = mablgh;
                                            if (zarib_sharestan == 1)
                                            {
                                                tttt.zarib_sharestan = true;
                                            }
                                            else
                                            {
                                                tttt.zarib_sharestan = false;

                                            }
                                            db.SaveChanges();
                                        }
                                        else
                                        {
                                         
                                            tttt.CaranMax = (float)CaranMax;
                                            tttt.CaranSaranePadash = (float)CaranSaranePadash;
                                            tttt.CaranMin = (float)CaranMin;
                                            tttt.CaranAcceptLimit = (float)CaranAcceptLimit;
                                            tttt.CaranCriterion = (float)CaranCriterion;
                                            tttt.CaranStandard = (float)CaranStandard;
                                            tttt.CaranAyabOZahab = (float)CaranAyabOZahab;
                                            tttt.CaranAstaneShoroPadash = (float)CaranAstaneShoroPadash;
                                            tttt.CaranAstaneShoroJarime = (float)CaranAstaneShoroJarime;
                                            tttt.CaranSaraneJarime = (float)CaranSaraneJarime;
                                            tttt.FK_Moalefe_ID = c.md_ID;
                                            tttt.sarmayeh = false;
                                            tttt.caranmablagh = mablgh;
                                            if (zarib_sharestan == 1)
                                            {
                                                tttt.zarib_sharestan = true;
                                            }
                                            else
                                            {
                                                tttt.zarib_sharestan = false;

                                            }
                                            db.SaveChanges();

                                        }
                                    }






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
                    if (IsFish.Count!=0)
                    {
                        if (!tbCaransSettingRepository.AddRange(IsFish))
                        {
                            transaction.Rollback();
                            return "ثبت کردن مولفه با خطا مواجه شد";
                        }
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



        public string GetDataFromExcel_step(HttpPostedFileBase MyExcelStream)
        {
            float CaranMin, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAyabOZahab, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID,number;
            string name = "";
            List<tb_Step_tax> step = new List<tb_Step_tax>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            var steptax = db.tb_Step_tax.ToList();
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
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null)
                                    {
                                        number = (int)Name.Value; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[1];
                                    if (float.TryParse(year.Value.ToString(), out float result))
                                    {
                                        CaranMin = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[2];
                                    if (float.TryParse(month.Value.ToString(), out float result2))
                                    {
                                        CaranMax = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    var CaranAcceptLimit2 = row.Cells[3];
                                    if (float.TryParse(CaranAcceptLimit2.Value.ToString(), out float result3))
                                    {
                                        CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranCriterion2 = row.Cells[4];
                                    if (int.TryParse(CaranCriterion2.Value.ToString(), out int CaranCriterion22))
                                    {
                                        CaranCriterion = CaranCriterion22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    //var CaranStandard2 = row.Cells[5];
                                    //if (int.TryParse(CaranStandard2.Value.ToString(), out int CaranStandard22))
                                    //{
                                    //    CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    if (steptax.Count != 0)
                                    {
                                        var t = steptax.Where(p => p.year == CaranCriterion&&p.number==number).FirstOrDefault();
                                        if (t != null)
                                        {
                                            t.number = number;
                                            t.Fromamount = CaranMin;
                                            t.Toamount = CaranMax;

                                            t.Percent_amount = CaranAcceptLimit;
                                            t.year = (int)CaranCriterion;
                                            db.SaveChanges();

                                        }
                                        else
                                        {
                                            tb_Step_tax stepp = new tb_Step_tax();
                                            stepp.number = number;
                                            stepp.Fromamount = CaranMin;
                                            stepp.Toamount = CaranMax;

                                            stepp.Percent_amount = CaranAcceptLimit;
                                            stepp.year = (int)CaranCriterion;

                                            step.Add(stepp);

                                        }
                                    }
                                    else
                                    {
                                        tb_Step_tax stepp = new tb_Step_tax();
                                        stepp.number = number;
                                        stepp.Fromamount = CaranMin;
                                        stepp.Toamount = CaranMax;

                                        stepp.Percent_amount = CaranAcceptLimit;
                                        stepp.year = (int)CaranCriterion;

                                        step.Add(stepp);

                                    }



                                }
                                db.tb_Step_tax.AddRange(step);
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
                    //if (IsFish.Count != 0)
                    //{
                    //    if (!tbCaransSettingRepository.AddRange(IsFish))
                    //    {
                    //        transaction.Rollback();
                    //        return "ثبت کردن مولفه با خطا مواجه شد";
                    //    }
                    //}

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















        #region دلایل پیش فرض
        #region variables

        tbDefault_reasonsRepository rep_def_rsn = new tbDefault_reasonsRepository();
        #endregion
        #region pages
        [AuthorizeAAA]
        public ActionResult Defaultreasons_manage()
        {
            return View("~/Areas/Setting/Views/Admin/Defaultreasons_manage.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult Defaultreasons_List()
        {
            return View("~/Areas/Setting/Views/Admin/Defaultreasons_List.cshtml",rep_def_rsn.Update());
        }

        [AuthorizeAAA]
        public ActionResult Defaultreasons_edit(int id)
        {
            var result = rep_def_rsn.Find(id);
            return View("~/Areas/Setting/Views/Admin/Defaultreasons_edit.cshtml",result);
        }

        [AuthorizeAAA]
        public ActionResult Defaultreasons_crate()
        {
            return View("~/Areas/Setting/Views/Admin/Defaultreasons_crate.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult DefultReson_Disabe(int id)
        {
            var result = rep_def_rsn.Find(id);
            return View("~/Areas/Setting/Views/Admin/DefultReson_Disabe.cshtml", result);

        }

        [AuthorizeAAA]
        public ActionResult _AddMoreResons()
        {
            return PartialView("~/Areas/Setting/Views/Admin/_AddMoreResons.cshtml");
        }

        #endregion
        #region events
        /// <summary>
        /// ذخیره لیست دلایل پیش فرض برای هر دسته بندی
        /// خودرو : 1
        /// عملکرد : 2
        /// محدودیت : 3
        /// ابزار : 4
        /// </summary>
        /// <param name="Filters"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public string DefaultReasons_Add(List<tbDefault_reasons> Filters)
        {
             return rep_def_rsn.Create(Filters);
        }

        [AuthorizeAAA]
        public string Defaultedit_Add(List<tbDefault_reasons> Filters)
        {
            return rep_def_rsn.Update(Filters);
            
        }

        [AuthorizeAAA]
        public string Defaultedit2_ADD(tbDefault_reasons Filters)
        {
            return rep_def_rsn.Update(Filters);
        }

        [AuthorizeAAA]
        public bool defaltReason_Delete(int ID)
        {
            return rep_def_rsn.Disable(ID);//deleted
        }





        public ActionResult SetVisibleForFish(int Month,int Year,bool IsVisible)
        {

            try
            {

                var Exist = db.tbVisibleFish.FirstOrDefault(p => p.Month == Month && p.Year == Year);


                if (Exist != null)//update
                {
                    Exist.IsVisibile = IsVisible;
                    db.SaveChanges();
                }
                else//insert
                {
                    tbVisibleFish Entity = new tbVisibleFish
                    {
                        Month = Month,
                        Year = Year,
                        IsVisibile = IsVisible
                    };
                    if (VsibileFish.Create(Entity) != "True")
                    {
                        return Content("0");
                    }
                }
                return Content("1");
            }
            catch (Exception)
            {

                return Content("0");
            }


           
           
        }

        public async Task<ActionResult> SetVisibleForFishSUBMITFISH(int Month, int Year, bool IsVisible)
        {
            try
            {
                var records = await db.tbMoalefeValueFish
                    .Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year)
                    .ToListAsync();  // ابتدا کل داده‌ها را دریافت می‌کنیم

                if (records.Any()) // چک می‌کنیم که داده‌ای برای آپدیت وجود دارد
                {
                    foreach (var item in records)
                    {
                        item.mlfvlfsh_Submit = IsVisible;
                    }

                    await db.SaveChangesAsync(); // فقط یکبار SaveChangesAsync اجرا شود

                    return Content("1");
                }
                else
                {
                    return Content("21"); // اگر رکوردی یافت نشد
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Content("0");
            }
        }


        public ActionResult SetVisibleForFishexch(int year_pY = 0, int dore = 0, bool isavailble = true)
        {
            try
            {
                var yy = db.tbSoratvaziat.Where(p => p.Year == year_pY && p.number_sorat == dore).ToList();
foreach(var it in yy)                {
                    it.Final_accept = isavailble;
                    db.SaveChanges();
                }
                //else
                ////{
                ////    tbCheaksorat n = new tbCheaksorat();
                ////    n.Year = year_pY;
                ////    n.Month = null;
                ////    n.Fk_pymn = null;
                ////    n.number_sorat = dore;

                ////    n.Isavailable = isavailble;

                ////    // Assuming tbCheaksorat is a DbSet in your DbContext
                ////    db.tbCheaksorat.Add(n);

                ////    db.SaveChanges();
                ////}


                return Content("True");
            }
            catch (Exception ex)
            {
                // Log the exception or handle it appropriately
                return Content($"{{'success': false, 'error': '{ex.Message}'}}");
            }
        }

            public ActionResult _VisibleOrNot(int Month,int Year)
        {
            var x = db.tbVisibleFish.FirstOrDefault(p => p.Month == Month && p.Year == Year);
            bool Model = false;
            if(x!=null)
            {
                Model = x.IsVisibile;
            }
           
            return PartialView("_VisibleOrNot", Model);
        }

        #endregion
        #endregion


    }
}