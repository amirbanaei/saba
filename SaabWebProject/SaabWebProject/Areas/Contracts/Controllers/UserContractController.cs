using OfficeOpenXml;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.Contracts.UserContracts;
using SaabWebProject.Utility;
using Syncfusion.XlsIO;
using System;
using GemBox.Spreadsheet;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Web.Spreadsheet;
using Microsoft.Office.Interop;
using static SaabWebProject.Areas.Users.Controllers.MessageBoxController;
using System.Threading.Tasks;
using System.Data.Entity;
using SaabWebProject.Models.Repositories.Contracts;
using Stimulsoft.Blockly.Model;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using ExcelLibrary.BinaryFileFormat;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.Repositories.Salaries.Formula;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using PdfSharp.Pdf.Content.Objects;
using Syncfusion.ExcelToPdfConverter;
using iTextSharp.text.pdf;
using Syncfusion.XlsIO;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf; // از این namespace استفاده می‌کنیم تا کلاس PdfDocument از Syncfusion.Pdf استفاده شود
using System.IO;
using iTextSharp.text;
using Syncfusion.Pdf.Graphics;
using DocumentFormat.OpenXml.Office2010.Excel;
using static Aspose.Pdf.Operator;
using System.IO.Compression;
using ICSharpCode.SharpZipLib.Zip;
using static SaabWebProject.Areas.Contracts.Controllers.UserContractController;
using System.Net.Http;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using SaabWebProject.Models.ViewModels.Contracts.Function;

namespace SaabWebProject.Areas.Contracts.Controllers

{
    public class UserContractController : Controller
    {
        #region تعریف متغیر ها
        SaabEntities context;
        tbUserContractsRepository rep_userContracts;
        tbUserRentContractsRepository rep_userrentContracts = new tbUserRentContractsRepository();
        PersianCalendar p = new PersianCalendar();
        tbUsersRepository rep_users;
        SaabEntities db=new SaabEntities();
        tbContractMoalefeDastMozdiRepository contractMoalefeDastMozdiRepository;
        tbUserSalaryDetailRepositories usersalaryRepo;
        tbMoalefeValueFishRepositories moalefeValueFishRepo;
        FishUtilities fishuti;
        SoratUtilities Soratuti;
        tb_Step_taxRepositories rf_tb_Step_tax = new tb_Step_taxRepositories();
        ZaribForFishRepository ZaribforFish = new ZaribForFishRepository();
        tbmoalfeExcelRepository ExelFish = new tbmoalfeExcelRepository();
        tbAddReduceMoadelkarkardRepository addorreduceRepo;
        tbSaleryIsCalculatedRepository saleryiscalcRepo;
        tbHighLowMoalefeValRepository highloeRepo;
        HighLowUtilities highlowUti;
        #endregion

        #region سازنده ها
        public UserContractController()
        {
            context = new SaabEntities();
            rep_userContracts = new tbUserContractsRepository(context);
            rep_users = new tbUsersRepository(context);
      
            contractMoalefeDastMozdiRepository = new tbContractMoalefeDastMozdiRepository();
            usersalaryRepo = new tbUserSalaryDetailRepositories(db);
            moalefeValueFishRepo = new tbMoalefeValueFishRepositories(db);
            fishuti = new FishUtilities(db);
            Soratuti = new SoratUtilities(db);
            addorreduceRepo = new tbAddReduceMoadelkarkardRepository(db);
            saleryiscalcRepo = new tbSaleryIsCalculatedRepository(db);
            highloeRepo = new tbHighLowMoalefeValRepository(db);
            highlowUti = new HighLowUtilities(db);
        }

        #endregion

        #region صفحات
        [AuthorizeAAA]

        /// <summary>
        /// صفحه لیست قرارداد های پرسنلی و اجاره ای
        /// </summary>
        /// <param name="user_ID"></param>
        /// <returns></returns>
        public ActionResult UserContract_List(int user_ID)
        {
            //ContractsListViewModel obj = new ContractsListViewModel();
            //obj.list_usercontracts = rep_userContracts.Listt(user_ID);
            //obj.list_rentContracts = rep_userrentContracts.Listt(user_ID);
            var contracts = rep_userContracts.Listt(user_ID);
            return View("~/Areas/Contracts/Views/UserContract/UserContract_List.cshtml", contracts);

        }
        public ActionResult View2()
        {
           
            return View("~/Areas/Contracts/Views/UserContract/View.cshtml", rep_userContracts.Listt());

        }
        public ActionResult Viewlistmamor()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;
            }
            var listt= db.tbUserContracts.Where(s=>s.FK_UserID== personal).ToList();

            return View("~/Areas/Contracts/Views/UserContract/Viewlistmamor.cshtml", listt);

        }

        [AuthorizeAAA]
        /// <summary>
        /// ثبت قرارداد جدید
        /// </summary>
        /// <returns></returns>
        public ActionResult UserContract_Create() // ok
        {
            return View("~/Areas/Contracts/Views/UserContract/UserContract_Create.cshtml");
        }
        [AuthorizeAAA]
        /// <summary>
        /// مدیریت قراردادها
        /// </summary>
        /// <returns></returns>
        public ActionResult ManageContracts() // ok
        {
            return View("~/Areas/Contracts/Views/UserContract/ManageContracts.cshtml");
        }
        public ActionResult UserContract_Details_Gozaresh(int contractId)
            
        {
            var find = db.tbUserContracts.Where(p => p.FK_UserID == contractId).OrderBy(s => s.usc_ID).FirstOrDefault();

            if(find != null)
            {
                contractId=find.usc_ID;
            }
            var Result = rep_userContracts.Find(contractId);
            //var userMoalefeGhararDadi = Result.tbUserContractsAndMoalefeGhararDadi.FirstOrDefault(u => u.FKContractID == Result.usc_ID);

            //// Check if userMoalefeGhararDadi is not null before assigning
            //if (userMoalefeGhararDadi != null)
            //{
            //    // Assuming tbUserContractsAndMoalefeGhararDadi is a collection, add userMoalefeGhararDadi to it
            //    Result.tbUserContractsAndMoalefeGhararDadi = new List<tbUserContractsAndMoalefeGhararDadi>(); // Initialize the collection if necessary
            //    Result.tbUserContractsAndMoalefeGhararDadi.Add(userMoalefeGhararDadi);
            //}

            Result.tbUsers = context.tbUsers.Find(Result.FK_UserID);
            Result.jobgroup_ValueMozdGroup = Result.jobgroup_ValueMozdGroup.Replace(",", "");
            Result.jobgroup_ValueSanavat = Result.jobgroup_ValueSanavat.Replace(",", "");
            Result.jobgroup_HagheOlad = Result.jobgroup_HagheOlad.Replace(",", "");
            Result.jobgroup_HagheMaskan = Result.jobgroup_HagheMaskan.Replace(",", "");
            Result.jobgroup_KharoBar = Result.jobgroup_KharoBar.Replace(",", "");

            Fillexceltst(Result);
            return View("~/Areas/Contracts/Views/UserContract/UserContract_Details.cshtml", Result);
        }
        [AuthorizeAAA]
        public ActionResult UserContract_Details(int contractId)
        {
            var Result = rep_userContracts.Find(contractId);
            //var userMoalefeGhararDadi = Result.tbUserContractsAndMoalefeGhararDadi.FirstOrDefault(u => u.FKContractID == Result.usc_ID);

            //// Check if userMoalefeGhararDadi is not null before assigning
            //if (userMoalefeGhararDadi != null)
            //{
            //    // Assuming tbUserContractsAndMoalefeGhararDadi is a collection, add userMoalefeGhararDadi to it
            //    Result.tbUserContractsAndMoalefeGhararDadi = new List<tbUserContractsAndMoalefeGhararDadi>(); // Initialize the collection if necessary
            //    Result.tbUserContractsAndMoalefeGhararDadi.Add(userMoalefeGhararDadi);
            //}

            Result.tbUsers = context.tbUsers.Find(Result.FK_UserID);
            Result.jobgroup_ValueMozdGroup = Result.jobgroup_ValueMozdGroup.Replace(",", "");
            Result.jobgroup_ValueSanavat = Result.jobgroup_ValueSanavat.Replace(",", "");
            Result.jobgroup_HagheOlad = Result.jobgroup_HagheOlad.Replace(",", "");
            Result.jobgroup_HagheMaskan = Result.jobgroup_HagheMaskan.Replace(",", "");
            Result.jobgroup_KharoBar = Result.jobgroup_KharoBar.Replace(",", "");

            Fillexceltst(Result);
            return View("~/Areas/Contracts/Views/UserContract/UserContract_Details.cshtml", Result);
        }
        #endregion
        public ActionResult Fillexceltst(tbUserContracts Model)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.xlsx"));
            var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx");

            var obj = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
            var birthplace = obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfBrith).Select(p => p.cityName).FirstOrDefault();
            var PlaceOfRegister = obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfIssue).Select(p => p.cityName).FirstOrDefault();
            var typeofcontract = "";
            string maragestatus = "";
            string DutySystem = "";
            string MadrakTahsili = "";

            if (Model.usc_TypeOfContract == false)
            {
                typeofcontract = "وقت مزدی";
            }
            else
            {
                typeofcontract = "کارمزدی";
            }

            if (Model.tbUsers.usr_MaritaIStatus == 1)
            {
                maragestatus = "متاهل";
            }
            else if (Model.tbUsers.usr_MaritaIStatus == 2)
            {
                maragestatus = "مجرد";
            }

            switch (Model.tbUsers.usr_Dutysystem)
            {
                case 1:
                    DutySystem = "مشمول";
                    break;
                case 2:
                    DutySystem = "پایان خدمت";
                    break;
                case 3:
                    DutySystem = "معافیت دائم";
                    break;
                case 4:
                    DutySystem = "معافیت موقت";
                    break;
                case 5:
                    DutySystem = "غایب";
                    break;
                case 6:
                    DutySystem = "خانم";
                    break;
                default:
                    DutySystem = "نامعلوم";
                    break;
            }

            if (Model.tbUsers.usr_Degree != null)
            {
                switch (Model.tbUsers.usr_Degree)
                {
                    case 1:
                        MadrakTahsili = "بیسواد";
                        break;
                    case 2:
                        MadrakTahsili = "سیکل";
                        break;
                    case 3:
                        MadrakTahsili = "دیپلم";
                        break;
                    case 4:
                        MadrakTahsili = "کاردانی";
                        break;
                    case 5:
                        MadrakTahsili = "کارشناسی";
                        break;
                    case 6:
                        MadrakTahsili = "کارشناسی ارشد";
                        break;
                    case 8:
                        MadrakTahsili = "دکتری";
                        break;
                    default:
                        MadrakTahsili = "نامعلوم";
                        break;
                }
            }
            else
            {
                MadrakTahsili = "نامعلوم";
            }

            try
            {
                using (var package = new ExcelPackage())
                {
                    using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
                    {
                        package.Load(stream);
                    }

                    var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);

                    if (worksheet2 != null)
                    {
                        var stringList = new List<string>();

                        #region fill string list
                        stringList.Add(Model.usc_ContractNumber);
                        stringList.Add(Model.tbUsers.usr_Family);
                        stringList.Add(Model.tbUsers.usr_Name);
                        stringList.Add(Model.tbUsers.usr_SHCode.ToString());
                        stringList.Add(Model.tbUsers.usr_FatherName);
                        stringList.Add(Model.tbUsers.shamsiDateOfBirth);
                        stringList.Add(birthplace);
                        stringList.Add(PlaceOfRegister);
                        stringList.Add(maragestatus);
                        stringList.Add(Model.tbUsers.usr_Child_Allowance.ToString());
                        stringList.Add(DutySystem);
                        stringList.Add(MadrakTahsili);
                        stringList.Add("---");
                        stringList.Add(Model.tbUsers.usr_NationalCode);
                        stringList.Add(Model.tbUsers.usr_Personal_ID.ToString());
                        stringList.Add(Model.tbCompanies.CompanyName);
                        stringList.Add(Model.tbCompanies.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies.Company_Address);
                        stringList.Add(Model.tbCompanies.RegistrationNumber.ToString());
                        stringList.Add(Model.tbCompanies.PlaceOfRegister.ToString());
                        stringList.Add(Model.usc_Jobtitle);
                        stringList.Add(Model.usc_JobCode.ToString());
                        stringList.Add(Model.FK_JobGroup.ToString());
                        stringList.Add(Model.tbCompanies1.CompanyName);
                        stringList.Add(Model.tbCompanies1.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies1.Company_Address);
                        stringList.Add(Model.usc_ShmasiStartTime);
                        stringList.Add(Model.usc_ShmasiENDTime);
                        stringList.Add((Model.usc_EndTime - Model.usc_StartTime).ToString());
                        stringList.Add(typeofcontract);
                        stringList.Add(Model.jobgroup_ValueMozdGroup);
                        stringList.Add(Model.jobgroup_ValueSanavat);
                        stringList.Add("0");
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) + System.Convert.ToInt32(Model.jobgroup_ValueSanavat)).ToString());
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30).ToString());
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30).ToString());
                        stringList.Add("0");
                        stringList.Add(((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30) + (System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30)).ToString());
                        stringList.Add(Model.jobgroup_HagheMaskan);
                        stringList.Add(Model.jobgroup_KharoBar);
                        stringList.Add(Model.jobgroup_HagheOlad);
                        var gharar = Model.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == Model.usc_ID).FirstOrDefault();
                        if (gharar != null)
                        {
                            stringList.Add(gharar.Value.ToString());
                            stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + gharar.Value + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());
                        }
                        else
                        {
                            stringList.Add("0");
                            stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());
                        }
                        #endregion


                        for (int i = 1; i <= 43; i++)
                        {
                            var cell = worksheet2.Cells[$"C{i}"];
                            cell.Value = stringList[i - 1];
                        }
                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }

                    var worksheet3 = package.Workbook.Worksheets[0];
                    if (worksheet3 != null)
                    {

                        worksheet3.Calculate();


                        var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address]
                            .Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();

                        if (cellsWithFormulas.Any())
                        {
                            foreach (var cell in cellsWithFormulas)
                            {
                                var value = cell.Value;
                                cell.Value = value;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
                        }

                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }
                }

                // بارگذاری فایل با Aspose
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(newFilePath);

                for (int i = 1; i < workbook.Worksheets.Count; i++)
                {
                    workbook.Worksheets[i].IsVisible = false;
                }


                workbook.Save(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), Aspose.Cells.SaveFormat.Pdf);

                System.Diagnostics.Debug.WriteLine("تبدیل به PDF با موفقیت انجام شد.");
                return Content("فایل با موفقیت پردازش و ذخیره شد.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("خطا در تبدیل فایل: " + ex.Message);
                System.Diagnostics.Debug.WriteLine("جزئیات خطا: " + ex.StackTrace);
                return Content("یک خطا رخ داده است: " + ex.Message);
            }
        }
        //E)------------------------------------------------------------|Mk|
        #region توابع

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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/" + filename));
                        var contract = rep_userContracts.Find(ID);
                        contract.usc_FileSystemName = filename;
                        contract.usc_FileName = file.FileName;
                        return rep_userContracts.Update(contract).ToString();
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

        public async Task< ActionResult> ExportMoalefeExcel4()
        {
            var OutPutFile = await SetDataExcel_Moalefe5();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Moalefestandard" + extension);
        }
        public async Task<Workbook> SetDataExcel_Moalefe5()
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
                            Value = item.md_Title,
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
                            Value = x2,
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
            return Moalefeexcelfile;
        }










        public async Task<string> GetDataFromExcel_Fish3(HttpPostedFileBase myExcelStream)
        {
            if (myExcelStream == null || myExcelStream.ContentLength == 0)
                return "فایل انتخاب نشده است";

            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(
                myExcelStream.InputStream,
                Path.GetExtension(myExcelStream.FileName));

            if (workbook.Sheets.Count < 1)
                return "هیچ شیتی در این اکسل وجود ندارد";

            var sheet = workbook.Sheets[0];
            var rows = sheet.Rows;

            if (rows == null || rows.Count == 0)
                return "این شیت فاقد سطر می باشد";

            if (rows.Count == 1)
                return "فایل اکسل فاقد اطلاعات می باشد";

            using (var transaction = db.Database.BeginTransaction())
            {
                var oldAutoDetect = db.Configuration.AutoDetectChangesEnabled;
                db.Configuration.AutoDetectChangesEnabled = false;

                try
                {
                    // فقط یک بار بخوان
                    var users = await db.tbUsers
                        .AsNoTracking()
                        .ToListAsync();

                    var userMap = users
                        .GroupBy(x => x.usr_Personal_ID)
                        .ToDictionary(g => g.Key, g => g.First());

                    var moalefeList = await db.tbContractMoalefeDastmozdi
                        .AsNoTracking()
                        .ToListAsync();

                    // اگر عنوان‌ها دقیق هستند بهتر است Equals بزنید نه Contains
                    var moalefeMap = moalefeList
                        .GroupBy(x => x.md_Title.Trim())
                        .ToDictionary(g => g.Key, g => g.First());

                    var titleCells = rows[0].Cells;

                    // چک تعداد ستون‌ها
                    int columnCount = titleCells.Count;
                    var dataRows = rows
        .Skip(1)
        .Where(r =>
            r.Cells != null &&
            r.Cells.Count > 1 &&
            r.Cells[1] != null &&
            r.Cells[1].Value != null &&
            !string.IsNullOrWhiteSpace(r.Cells[1].Value.ToString()))
        .ToList();

                    if (dataRows.Count == 0)
                        return "فایل اکسل فاقد اطلاعات معتبر می باشد";
                    //for (int r = 1; r < rows.Count; r++)
                    //{
                    //    if (rows[r].Cells.Count != columnCount)
                    //        return $"خطای ارزیابی در سطر {r + 1} رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد";
                    //}

                    // هدر مولفه‌ها فقط یک بار
                    var lstMoalefe = new List<string>();
                    for (int i = 17; i < titleCells.Count; i++)
                    {
                        var val = titleCells[i]?.Value?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(val))
                            lstMoalefe.Add(val);
                    }

                    // آخرین شماره قرارداد فقط یک بار
                    var lastContractNumber = await db.tbUserContracts
                        .AsNoTracking()
                        .OrderByDescending(x => x.usc_ID)
                        .Select(x => x.usc_ContractNumber)
                        .FirstOrDefaultAsync();

                    int counterContractNumber = 0;
                    if (!string.IsNullOrWhiteSpace(lastContractNumber))
                    {
                        var lastPart = lastContractNumber.Split('_').LastOrDefault();
                        int.TryParse(lastPart, out counterContractNumber);
                    }

                    var contractsToAdd = new List<tbUserContracts>();
                    var tempRows = new List<(tbUserContracts Contract, Telerik.Web.Spreadsheet.Row Row)>();

                    for (int i = 1; i < rows.Count; i++)
                    {
                        var row = rows[i];

                        if (!int.TryParse(row.Cells[1]?.Value?.ToString(), out int personnelId))
                            return $"لطفا ستون 1 سطر ({i}) را پر کنید";

                        if (!int.TryParse(row.Cells[2]?.Value?.ToString(), out int contractType))
                            return $"لطفا ستون 2 سطر ({i}) را پر کنید";

                        if (!int.TryParse(row.Cells[3]?.Value?.ToString(), out int karfarmaId))
                            return $"لطفا ستون 3 سطر ({i}) را پر کنید";

                        if (!int.TryParse(row.Cells[4]?.Value?.ToString(), out int karfarmaAsli))
                            return $"لطفا ستون 4 سطر ({i}) را پر کنید";

                        if (!int.TryParse(row.Cells[5]?.Value?.ToString(), out int jobGroup))
                            return $"لطفا ستون 5 سطر ({i}) را پر کنید";

                        if (!int.TryParse(row.Cells[6]?.Value?.ToString(), out int sanavatYears))
                            return $"لطفا ستون 6 سطر ({i}) را پر کنید";

                        var timestart = row.Cells[7]?.Value?.ToString();
                        if (string.IsNullOrWhiteSpace(timestart))
                            return $"لطفا ستون 8 سطر ({i}) را پر کنید";

                        var timeend = row.Cells[8]?.Value?.ToString();
                        if (string.IsNullOrWhiteSpace(timeend))
                            return $"لطفا ستون 9 سطر ({i}) را پر کنید";

                        var mozdd = row.Cells[9]?.Value?.ToString() ?? "";
                        var sanavattt = row.Cells[10]?.Value?.ToString() ?? "";
                        var oladd = row.Cells[11]?.Value?.ToString() ?? "";
                        var maskan = row.Cells[12]?.Value?.ToString() ?? "";
                        var gharbar = row.Cells[13]?.Value?.ToString() ?? "";
                        var sharh = row.Cells[14]?.Value?.ToString() ?? "";
                        var shraytfasgh = row.Cells[15]?.Value?.ToString() ?? "";
                        var sharhvazayf = row.Cells[16]?.Value?.ToString() ?? "";

                        if (!userMap.TryGetValue(personnelId, out var user))
                            return $"کاربری با کد پرسنلی {personnelId} در سطر ({i}) یافت نشد";

                        DateTime persianDate;
                        DateTime persianDate2;

                        try
                        {
                            persianDate = PersianDateToDateTime(timestart);
                            persianDate2 = PersianDateToDateTime(timeend);
                        }
                        catch
                        {
                            return $"فرمت تاریخ در سطر ({i}) معتبر نیست";
                        }

                        counterContractNumber++;

                        var contract = new tbUserContracts
                        {
                            usc_StartTime = persianDate,
                            usc_EndTime = persianDate2,
                            FK_KarfarmaID = karfarmaId,
                            sharh = sharh,
                            Sharhvazayf = sharhvazayf,
                            shraytfasgh = shraytfasgh,
                            FK_KarfarmaAsli = karfarmaAsli,
                            FK_UserID = user.usr_ID,
                            usc_TypeOfContract = contractType == 1,
                            FK_JobGroup = jobGroup,
                            usc_Jobtitle = null,
                            jobgroup_HagheMaskan = maskan,
                            jobgroup_HagheOlad = oladd,
                            jobgroup_KharoBar = gharbar,
                            jobgroup_ValueSanavat = sanavattt,
                            jobgroup_ValueMozdGroup = mozdd,
                            usc_TedadSalSanavat = sanavatYears,
                            usc_JobCode = 0,
                            usc_ContractNumber = $"{SaabWebProject.Models.Utilitis.ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now)}_{user.usr_Personal_ID}_{counterContractNumber}"
                        };

                        contractsToAdd.Add(contract);
                        tempRows.Add((contract, row));
                    }

                    db.tbUserContracts.AddRange(contractsToAdd);
                    await db.SaveChangesAsync();

                    var contractMoalefeToAdd = new List<tbUserContractsAndMoalefeGhararDadi>();

                    foreach (var item in tempRows)
                    {
                        int shomarande = 17;

                        foreach (var moalefeTitle in lstMoalefe)
                        {
                            if (!moalefeMap.TryGetValue(moalefeTitle, out var moalefe))
                            {
                                shomarande++;
                                continue;
                            }

                            var cellValue = item.Row.Cells.Count > shomarande
                                ? item.Row.Cells[shomarande]?.Value?.ToString()
                                : null;

                            int parsedValue = 0;
                            if (!string.IsNullOrWhiteSpace(cellValue))
                                int.TryParse(cellValue, out parsedValue);

                            contractMoalefeToAdd.Add(new tbUserContractsAndMoalefeGhararDadi
                            {
                                FKContractID = item.Contract.usc_ID,
                                FKMoalefeGhararDadi = moalefe.md_ID,
                                Value = parsedValue
                            });

                            shomarande++;
                        }
                    }

                    if (contractMoalefeToAdd.Count > 0)
                    {
                        db.tbUserContractsAndMoalefeGhararDadi.AddRange(contractMoalefeToAdd);
                        await db.SaveChangesAsync();
                    }

                    transaction.Commit();
                    return "با موفقیت انجام شد";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;
                }
                finally
                {
                    db.Configuration.AutoDetectChangesEnabled = true;
                }
            }
        }
        public async Task<string> GetDataFromExcel_Fish(HttpPostedFileBase myExcelStream)
        {
            if (myExcelStream == null || myExcelStream.ContentLength == 0)
                return "فایل انتخاب نشده است";

            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(
                myExcelStream.InputStream,
                Path.GetExtension(myExcelStream.FileName));

            if (workbook.Sheets.Count < 1)
                return "هیچ شیتی در این اکسل وجود ندارد";

            var sheet = workbook.Sheets[0];
            var rows = sheet.Rows;

            if (rows == null || rows.Count == 0)
                return "این شیت فاقد سطر می باشد";

            var titleCells = rows[0].Cells;
            if (titleCells == null || titleCells.Count == 0)
                return "سطر عنوان فایل نامعتبر است";

            // فقط ردیف‌هایی که ستون 1 آنها مقدار دارد
            var dataRows = rows
                .Skip(1)
                .Where(r =>
                    r.Cells != null &&
                    r.Cells.Count > 1 &&
                    r.Cells[1] != null &&
                    r.Cells[1].Value != null &&
                    !string.IsNullOrWhiteSpace(r.Cells[1].Value.ToString()))
                .ToList();

            if (dataRows.Count == 0)
                return "فایل اکسل فاقد اطلاعات معتبر می باشد";

            using (var transaction = db.Database.BeginTransaction())
            {
                var oldAutoDetect = db.Configuration.AutoDetectChangesEnabled;
                db.Configuration.AutoDetectChangesEnabled = false;

                try
                {
                    var users = await db.tbUsers
                        .AsNoTracking()
                        .ToListAsync();

                    var userMap = users
                        .GroupBy(x => x.usr_Personal_ID)
                        .ToDictionary(g => g.Key, g => g.First());

                    var moalefeList = await db.tbContractMoalefeDastmozdi
                        .AsNoTracking()
                        .ToListAsync();

                    var moalefeMap = moalefeList
                        .Where(x => x.md_Title != null)
                        .GroupBy(x => x.md_Title.Trim())
                        .ToDictionary(g => g.Key, g => g.First());

                    // چک تعداد ستون‌ها فقط روی ردیف‌های معتبر
                    int columnCount = titleCells.Count;
                    foreach (var row in dataRows)
                    {
                        if (row.Cells.Count != columnCount)
                            return $"خطای ارزیابی در سطر {row.Index + 1} رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد";
                    }

                    // هدر مولفه‌ها فقط یک بار
                    var lstMoalefe = new List<string>();
                    for (int i = 17; i < titleCells.Count; i++)
                    {
                        var val = titleCells[i]?.Value?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(val))
                            lstMoalefe.Add(val);
                    }

                    var lastContractNumber = await db.tbUserContracts
                        .AsNoTracking()
                        .OrderByDescending(x => x.usc_ID)
                        .Select(x => x.usc_ContractNumber)
                        .FirstOrDefaultAsync();

                    int counterContractNumber = 0;
                    if (!string.IsNullOrWhiteSpace(lastContractNumber))
                    {
                        var lastPart = lastContractNumber.Split('_').LastOrDefault();
                        int.TryParse(lastPart, out counterContractNumber);
                    }

                    var contractsToAdd = new List<tbUserContracts>();
                    var tempRows = new List<(tbUserContracts Contract, Telerik.Web.Spreadsheet.Row Row)>();

                    foreach (var row in dataRows)
                    {
                        int excelRowNumber = (int)(row.Index + 1);

                        if (!int.TryParse(row.Cells[1]?.Value?.ToString(), out int personnelId))
                            return $"لطفا ستون 1 سطر ({excelRowNumber}) را پر کنید";

                        if (!int.TryParse(row.Cells[2]?.Value?.ToString(), out int contractType))
                            return $"لطفا ستون 2 سطر ({excelRowNumber}) را پر کنید";

                        if (!int.TryParse(row.Cells[3]?.Value?.ToString(), out int karfarmaId))
                            return $"لطفا ستون 3 سطر ({excelRowNumber}) را پر کنید";

                        if (!int.TryParse(row.Cells[4]?.Value?.ToString(), out int karfarmaAsli))
                            return $"لطفا ستون 4 سطر ({excelRowNumber}) را پر کنید";

                        if (!int.TryParse(row.Cells[5]?.Value?.ToString(), out int jobGroup))
                            return $"لطفا ستون 5 سطر ({excelRowNumber}) را پر کنید";

                        if (!int.TryParse(row.Cells[6]?.Value?.ToString(), out int sanavatYears))
                            return $"لطفا ستون 6 سطر ({excelRowNumber}) را پر کنید";

                        var timestart = row.Cells[7]?.Value?.ToString();
                        if (string.IsNullOrWhiteSpace(timestart))
                            return $"لطفا ستون 8 سطر ({excelRowNumber}) را پر کنید";

                        var timeend = row.Cells[8]?.Value?.ToString();
                        if (string.IsNullOrWhiteSpace(timeend))
                            return $"لطفا ستون 9 سطر ({excelRowNumber}) را پر کنید";

                        var mozdd = row.Cells[9]?.Value?.ToString() ?? "";
                        var sanavattt = row.Cells[10]?.Value?.ToString() ?? "";
                        var oladd = row.Cells[11]?.Value?.ToString() ?? "";
                        var maskan = row.Cells[12]?.Value?.ToString() ?? "";
                        var gharbar = row.Cells[13]?.Value?.ToString() ?? "";
                        var sharh = row.Cells[14]?.Value?.ToString() ?? "";
                        var shraytfasgh = row.Cells[15]?.Value?.ToString() ?? "";
                        var sharhvazayf = row.Cells[16]?.Value?.ToString() ?? "";

                        if (!userMap.TryGetValue(personnelId, out var user))
                            return $"کاربری با کد پرسنلی {personnelId} در سطر ({excelRowNumber}) یافت نشد";

                        DateTime persianDate;
                        DateTime persianDate2;

                        try
                        {
                            persianDate = PersianDateToDateTime(timestart);
                            persianDate2 = PersianDateToDateTime(timeend);
                        }
                        catch
                        {
                            return $"فرمت تاریخ در سطر ({excelRowNumber}) معتبر نیست";
                        }

                        counterContractNumber++;

                        var contract = new tbUserContracts
                        {
                            usc_StartTime = persianDate,
                            usc_EndTime = persianDate2,
                            FK_KarfarmaID = karfarmaId,
                            sharh = sharh,
                            Sharhvazayf = sharhvazayf,
                            shraytfasgh = shraytfasgh,
                            FK_KarfarmaAsli = karfarmaAsli,
                            FK_UserID = user.usr_ID,
                            usc_TypeOfContract = contractType == 1,
                            FK_JobGroup = jobGroup,
                            usc_Jobtitle = null,
                            jobgroup_HagheMaskan = maskan,
                            jobgroup_HagheOlad = oladd,
                            jobgroup_KharoBar = gharbar,
                            jobgroup_ValueSanavat = sanavattt,
                            jobgroup_ValueMozdGroup = mozdd,
                            usc_TedadSalSanavat = sanavatYears,
                            usc_JobCode = 0,
                            usc_ContractNumber = $"{SaabWebProject.Models.Utilitis.ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now)}_{user.usr_Personal_ID}_{counterContractNumber}"
                        };

                        contractsToAdd.Add(contract);
                        tempRows.Add((contract, row));
                    }

                    db.tbUserContracts.AddRange(contractsToAdd);
                    await db.SaveChangesAsync();

                    var contractMoalefeToAdd = new List<tbUserContractsAndMoalefeGhararDadi>();

                    foreach (var item in tempRows)
                    {
                        int shomarande = 17;

                        foreach (var moalefeTitle in lstMoalefe)
                        {
                            if (!moalefeMap.TryGetValue(moalefeTitle, out var moalefe))
                            {
                                shomarande++;
                                continue;
                            }

                            var cellValue = item.Row.Cells.Count > shomarande
                                ? item.Row.Cells[shomarande]?.Value?.ToString()
                                : null;

                            int parsedValue = 0;
                            if (!string.IsNullOrWhiteSpace(cellValue))
                                int.TryParse(cellValue, out parsedValue);

                            contractMoalefeToAdd.Add(new tbUserContractsAndMoalefeGhararDadi
                            {
                                FKContractID = item.Contract.usc_ID,
                                FKMoalefeGhararDadi = moalefe.md_ID,
                                Value = parsedValue
                            });

                            shomarande++;
                        }
                    }

                    if (contractMoalefeToAdd.Count > 0)
                    {
                        db.tbUserContractsAndMoalefeGhararDadi.AddRange(contractMoalefeToAdd);
                        await db.SaveChangesAsync();
                    }

                    transaction.Commit();
                    return "با موفقیت انجام شد";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;
                }
                finally
                {
                    db.Configuration.AutoDetectChangesEnabled = true;
                }
            }
        }

        public async Task<ActionResult> ImportExcel_AddFish(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message =await GetDataFromExcel_Fish(MyExcelStream);
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

        public async  Task< string> GetDataFromExcel_Fish2(HttpPostedFileBase MyExcelStream)
        {
            int CaranMin, countDays, sanavatt,olad,mask,shoghl,ghar,CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID;
            string name = ""; string timestart = ""; string timeend = ""; string CaranAyabOZahab = "";string month4 = ""; string month2 = "";string year2 = ""; string year22 = ""; string day22 = "";
            string oladd, maskan, gharbar, mozdd, sanavattt, Sharhvazayf, shraytfasgh, sharh;
            var usr=await db.tbUsers.ToListAsync();
            var moalfe = await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<string> lstMoalefe = new List<string>();
            List<tbUserContracts> Gharardad = new List<tbUserContracts>();

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
                                for (int i = 17; i < title.Count; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString());
                                }

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
                                        return " لطفا ستون 1  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var year = row.Cells[2];//نوع قرارداد
                                    if (int.TryParse(year.Value.ToString(), out int result))
                                    {
                                        CaranMin = result;
                                    }
                                    else
                                    {
                                        return " لطفا ستون2  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var month = row.Cells[3];//کارفرما
                                    if (int.TryParse(month.Value.ToString(), out int result2))
                                    {
                                        CaranMax = result2; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 3  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }



                                    var CaranAcceptLimit2 = row.Cells[4];//کارفرما اصلی
                                    if (int.TryParse(CaranAcceptLimit2.Value.ToString(), out int result3))
                                    {
                                        CaranAcceptLimit = result3; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 4  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranCriterion2 = row.Cells[5];//گروه شغلی
                                    if (int.TryParse(CaranCriterion2.Value.ToString(), out int CaranCriterion22))
                                    {
                                        CaranCriterion = CaranCriterion22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 5  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var CaranStandard2 = row.Cells[6];//sanavat
                                    if (int.TryParse(CaranStandard2.Value.ToString(), out int CaranStandard22))
                                    {
                                        CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 6  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    //var CaranStandard322 = row.Cells[7];//sanavat
                                    //if (int.TryParse(CaranStandard322.Value.ToString(), out int CaranStandard222))
                                    //{
                                    //    ghar = CaranStandard222; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    ghar=0;
                                    //}

                                    //var CaranAyabOZahab2 = row.Cells[8];//عنوان شعل
                                    //if (CaranAyabOZahab2!=null)
                                    //{
                                    //    CaranAyabOZahab = CaranAyabOZahab2.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 7  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}


                                    var CaranAstaneShoroPadash2 = row.Cells[7];//روز
                                    if (CaranAstaneShoroPadash2.Value!=null)
                                    {
                                        timestart = CaranAstaneShoroPadash2.Value != null ? CaranAstaneShoroPadash2.Value.ToString() : "";
                                    }
                                    else
                                    {
                                        return " لطفا ستون 8  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranAstaneShoroPadash22 = row.Cells[8];//روز
                                    if (CaranAstaneShoroPadash22.Value != null)
                                    {
                                        timeend = CaranAstaneShoroPadash22.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 9  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var mozd = row.Cells[9];//مزد شغل
                                    if (mozd.Value.ToString()!=null)
                                    {
                                        mozdd = mozd.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 10  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var sanavat = row.Cells[10];//سنوات
                                    if (sanavat.Value.ToString()!=null)
                                    {
                                        sanavattt = sanavat.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 11  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var oladdd = row.Cells[11];//اولاد
                                    if (oladdd.Value.ToString()!=null)
                                    {
                                        oladd = oladdd.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 12  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var maskans = row.Cells[12];//مسکن
                                    if (maskans.Value.ToString()!=null)
                                    {
                                        maskan = maskans.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 12  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var gharr = row.Cells[13];//خواروبار
                                    if (gharr.Value.ToString()!=null)
                                    {
                                        gharbar = gharr.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 14  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var shraytfasgh1 = row.Cells[15];//خواروبار
                                    if (shraytfasgh1.Value.ToString() != null)
                                    {
                                        shraytfasgh = shraytfasgh1.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 14  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var sharh1 = row.Cells[14];//خواروبار
                                    if (sharh1.Value.ToString() != null)
                                    {
                                        sharh = sharh1.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 14  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var Sharhvazayf1 = row.Cells[16];//خواروبار
                                    if (Sharhvazayf1.Value.ToString() != null)
                                    {
                                        Sharhvazayf = Sharhvazayf1.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون 14  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }








                                    tbUserContracts obj4 = new tbUserContracts();
                                    DateTime? usc_StartTime = null;

                                    string inputDate = timestart;
                                    string inputDate2 =timeend;

                                    // تبدیل تاریخ به تاریخ میلادی
                                    DateTime persianDate = PersianDateToDateTime(inputDate);
                                    DateTime persianDate2 = PersianDateToDateTime(inputDate2);

                                    // تبدیل تاریخ به فرمت مناسب برای SQL Server
                                    string sqlFormattedDate = persianDate.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                               var tt=     usr.FirstOrDefault(p => p.usr_Personal_ID == countDays);
                                    // اختصاص تاریخ به usc_StartTime
                                    obj4.usc_StartTime = persianDate;
                                    obj4.usc_EndTime = persianDate2;
                                    obj4.FK_KarfarmaID = CaranMax;
                                    obj4.sharh = sharh;
                                    obj4.Sharhvazayf = Sharhvazayf;
                                    obj4.shraytfasgh = shraytfasgh;

                                    obj4.FK_KarfarmaAsli = CaranAcceptLimit;
                                    obj4.FK_UserID = tt.usr_ID;
                                    if (CaranMin == 1)
                                    {
                                        obj4.usc_TypeOfContract =true ;

                                    }
                                    else
                                    {
                                        obj4.usc_TypeOfContract = false;

                                    }



                                    obj4.FK_JobGroup = CaranCriterion;
                                    obj4.usc_Jobtitle = null;
                                    obj4.jobgroup_HagheMaskan = maskan;
                                    obj4.jobgroup_HagheOlad = oladd;
                                    obj4.jobgroup_KharoBar = gharbar;
                                    obj4.jobgroup_ValueSanavat = sanavattt;
                                    obj4.jobgroup_ValueMozdGroup = mozdd;
                                    obj4.usc_TedadSalSanavat = CaranStandard;
                                    obj4.usc_JobCode = 0;

                                    int counter_ContractNumber = 0;

                                    var list_contracts = db.tbUserContracts.ToList();
                                    if (list_contracts.Count() != 0)
                                    {
                                        var lastnumber = list_contracts.LastOrDefault().usc_ContractNumber;

                                        if (lastnumber != null)
                                        {
                                            counter_ContractNumber = System.Convert.ToInt32(lastnumber.Split('_').Last());
                                        }
                                    }
                                    counter_ContractNumber++;

                                    var userpersonnel =tt.usr_Personal_ID;
                                    obj4.usc_ContractNumber = SaabWebProject.Models.Utilitis.ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now) + "_" + userpersonnel + "_" + counter_ContractNumber;
                                    db.tbUserContracts.Add(obj4);
                                    await db.SaveChangesAsync();
                                    int shomarande = 17;
                                    List<tbUserContractsAndMoalefeGhararDadi> objectsToAdd = new List<tbUserContractsAndMoalefeGhararDadi>();

                                    foreach (var it in lstMoalefe)
                                    {
                                        tbUserContractsAndMoalefeGhararDadi obj5=new tbUserContractsAndMoalefeGhararDadi();
                                        var id =  moalfe.Where(p => p.md_Title.Contains(it)).FirstOrDefault();
                                        var Value = row.Cells[shomarande];
                                        var value222 = Value.Value ?? null;
                                        int? countDays2;
                                        if (int.TryParse(value222.ToString(), out int parsedValue3))
                                        {
                                            countDays2 = parsedValue3;
                                        }
                                        else
                                        {
                                            countDays2 = 0; // Or handle the case where the value cannot be parsed
                                        }
                                        obj5.FKContractID = obj4.usc_ID;
                                        obj5.FKMoalefeGhararDadi = id.md_ID;
                                        obj5.Value = countDays2;
                                        objectsToAdd.Add(obj5);
                                        shomarande++;

                                    }

                                    db.tbUserContractsAndMoalefeGhararDadi.AddRange(objectsToAdd);
                                    await db.SaveChangesAsync();

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



        #region قرارداد پرسنل

        [AuthorizeAAA] 
        /// <summary>
        /// ذخیره قرارداد جدید
        /// </summary>
        /// <param name="obj_ContractsListViewModel"></param>
        /// <returns></returns>
        public string UserContract_Save(CreateListOfUsersContract obj_ContractsListViewModel)
        {
            obj_ContractsListViewModel.obj_tbuserContracts.usc_StartTime = p.ToDateTime(obj_ContractsListViewModel.obj_tbuserContracts.ShamsiStartTime_year, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiStartTime_month, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiStartTime_day, 0, 0, 0, 0).Date;
            obj_ContractsListViewModel.obj_tbuserContracts.usc_EndTime = p.ToDateTime(obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_year, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_month, obj_ContractsListViewModel.obj_tbuserContracts.ShamsiENDTime_day, 0, 0, 0, 0).Date;
            var flag = false;
            foreach (var item in obj_ContractsListViewModel.list_usersIDs) // چک کردن عدم همپوشانی
            {
                if (!CheckForOverLapConracts(obj_ContractsListViewModel.obj_tbuserContracts.usc_StartTime, obj_ContractsListViewModel.obj_tbuserContracts.usc_EndTime, item)) // عدمهم پوشانی زمانی با بقیه قرار داد ها{
                {
                    flag = true;


                }
            }
            if (flag == true)
            {
                return ("Falsee_1");
            }
            else
            {
                List<tbUserContracts> list_userContract = new List<tbUserContracts>();
                tbUserContracts obj1 = new tbUserContracts();
                #region بدست آوردن شماره قرارداد
                int counter_ContractNumber = 0;
              
                    var list_contracts = db.tbUserContracts.ToList();
                    if (list_contracts.Count() != 0)
                    {
                        var lastnumber = list_contracts.LastOrDefault().usc_ContractNumber;

                        if (lastnumber != null)
                        {
                            counter_ContractNumber =System.Convert.ToInt32(lastnumber.Split('_').Last());
                        }
                    }

                
                #endregion
                foreach (var item in obj_ContractsListViewModel.list_usersIDs) // ادد کردن لیست
                {
                    //string contractNumber = "";
                    obj1 = new tbUserContracts();

                   
                        counter_ContractNumber++;

                        var userpersonnel = db.tbUsers.Find(item).usr_Personal_ID;
                        obj1.usc_ContractNumber = SaabWebProject.Models.Utilitis.ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now) + "_" + userpersonnel + "_" + counter_ContractNumber;
                    

                    obj1.FK_JobGroup = obj_ContractsListViewModel.obj_tbuserContracts.FK_JobGroup;
                    obj1.FK_KarfarmaID = obj_ContractsListViewModel.obj_tbuserContracts.FK_KarfarmaID;
                    obj1.FK_KarfarmaAsli = obj_ContractsListViewModel.obj_tbuserContracts.FK_KarfarmaAsli;
                    obj1.FK_UserID = item;
                    obj1.jobgroup_HagheMaskan = obj_ContractsListViewModel.obj_tbuserContracts.jobgroup_HagheMaskan;
                    obj1.jobgroup_HagheOlad = obj_ContractsListViewModel.obj_tbuserContracts.jobgroup_HagheOlad;
                    obj1.jobgroup_KharoBar = obj_ContractsListViewModel.obj_tbuserContracts.jobgroup_KharoBar;
                    obj1.jobgroup_ValueMozdGroup = obj_ContractsListViewModel.obj_tbuserContracts.jobgroup_ValueMozdGroup;
                    obj1.jobgroup_ValueSanavat = obj_ContractsListViewModel.obj_tbuserContracts.jobgroup_ValueSanavat;
                    //obj1.FK_job = obj_ContractsListViewModel.obj_tbuserContracts.FK_job;
                    
                    // obj1.usc_ContractNumber = obj_ContractsListViewModel.obj_tbuserContracts.usc_ContractNumber;
                    obj1.usc_EndTime = obj_ContractsListViewModel.obj_tbuserContracts.usc_EndTime;
                    obj1.usc_JobCode = obj_ContractsListViewModel.obj_tbuserContracts.usc_JobCode;
                    obj1.usc_Jobtitle = obj_ContractsListViewModel.obj_tbuserContracts.usc_Jobtitle;
                    obj1.usc_StartTime = obj_ContractsListViewModel.obj_tbuserContracts.usc_StartTime;
                    obj1.usc_TedadSalSanavat = obj_ContractsListViewModel.obj_tbuserContracts.usc_TedadSalSanavat;
                    obj1.usc_TypeOfContract = obj_ContractsListViewModel.obj_tbuserContracts.usc_TypeOfContract;
                    obj1.sharh = obj_ContractsListViewModel.obj_tbuserContracts.sharh;
                    obj1.shraytfasgh = obj_ContractsListViewModel.obj_tbuserContracts.shraytfasgh;
                    obj1.Sharhvazayf = obj_ContractsListViewModel.obj_tbuserContracts.Sharhvazayf;

                    obj1.tbUserContractsAndMoalefeGhararDadi = obj_ContractsListViewModel.obj_tbuserContracts.tbUserContractsAndMoalefeGhararDadi;

                    list_userContract.Add(obj1);
                }

                if (list_userContract != null)
                {
                    //sabt();
                    //rep_userContracts.sabt();
                    return rep_userContracts.Create(list_userContract).ToString();
                }

                return "Falsee_WithoutUser";
            }



        }
        //public int sabt()
        //{
        //    tbUserContractsAndMoalefeGhararDadi obj1 = new tbUserContractsAndMoalefeGhararDadi();
        //    obj1.FKMoalefeGhararDadi = 1800;
        //    obj1.Value = 45;
        //    obj1.FKContractID = 440;
        //    db.tbUserContractsAndMoalefeGhararDadi.Add(obj1);
        //    db.SaveChanges();
        //    return 1;

        //}
        /// <summary>
        /// آپدیت  اطلاعات قرارداد
        /// </summary>
        /// <param name="objUsercontract"></param>
        /// <returns></returns>
        [AuthorizeAAA]
        public string UserContract_Update(tbUserContracts objUsercontract)
        {
            if (CheckForOverLapConracts(objUsercontract.usc_StartTime, objUsercontract.usc_EndTime, objUsercontract.FK_UserID)) // عدمهم پوشانی زمانی با بقیه قرار داد ها
                return rep_userContracts.Update(objUsercontract).ToString();
            else
                return "false_overlap";

        }
        /// <summary>
        /// تایید قرارداد
        /// </summary>
        /// <param name="userContractID"></param>
        /// <returns></returns>
       
        //public bool UserContract_Accept(int userContractID)
        //{
        //    return false ;
        //    //return rep_userContracts.AcceptUserContract(userContractID);
        //}
        /// <summary>
        /// حذف قرارداد
        /// </summary>
        /// <param name="userContractID"></param>
        /// <returns></returns>
        [AuthorizeAAA]
        public string UserContract_Delete(int userContractID)
        {
            return rep_userContracts.Disable(userContractID); // delete
        }
        /// <summary>
        /// آپلود فایل های قرارداد
        /// </summary>
        /// <param name="ContractID"></param>
        /// <param name="files"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        [AuthorizeAAA]
        public bool UserContract_UploadFile(int ContractID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Content/ContractFiles/" + filename));
                        var contract = rep_userContracts.Find(ContractID);
                        contract.usc_FileSystemName = filename;
                        contract.usc_FileName = file.FileName;
                        return rep_userContracts.Update(contract);
                    }
                    else
                    {
                        return true;
                    }
                }
                return true;
            }
            else
            {
                return true;
            }


        }
        #endregion

       public ActionResult viewpdffishorexcel()
        {
            return View();
        }

        public ActionResult UserContract_Details10(List<string> Fkpym, DateTime az, DateTime ta)
        {
            List<tbUserContracts> contractID = new List<tbUserContracts>();

            PersianCalendar pc = new PersianCalendar();
            DateTime dateTimeaz = (DateTime)az;
            DateTime dateTimeta = (DateTime)ta;

            string shamsiDateaz = $"{pc.GetYear(dateTimeaz)}/{pc.GetMonth(dateTimeaz):00}/{pc.GetDayOfMonth(dateTimeaz):00}";
            string shamsiDateta = $"{pc.GetYear(dateTimeta)}/{pc.GetMonth(dateTimeta):00}/{pc.GetDayOfMonth(dateTimeta):00}";

            var userIds = Fkpym.Where(x => int.TryParse(x, out _)) // فقط مقادیر عددی معتبر
                                    .Select(int.Parse).ToList();

            //var contractIDfkusr = db.tbUserContracts.GroupBy(s => s.FK_UserID).ToList();
            if (userIds.Count == 1)
            {

                foreach (var it1 in userIds)
                {
                    if (it1 == -1)
                    {
                        foreach (var it in userIds)
                        {
                            var contractIDfkusr = db.tbUserContracts.Where(s => s.usc_StartTime <= az && s.usc_EndTime >= ta).GroupBy(s => s.FK_UserID).ToList();
                            foreach(var it2 in contractIDfkusr)
                            {
                                var find = db.tbUserContracts.Where(s => s.FK_UserID == it2.Key && s.usc_StartTime <= az && s.usc_EndTime >= ta).OrderByDescending(s => s.usc_ID).FirstOrDefault();
                                if (find != null)
                                {
                                    contractID.Add(find);
                                }
                            }
                          
                        }
                    }
                    else
                    {
                        foreach (var it in userIds)
                        {
                            var find = db.tbUserContracts.Where(s => s.FK_UserID == it && s.usc_StartTime <= az && s.usc_EndTime >= ta).OrderByDescending(s => s.usc_ID).FirstOrDefault();
                            if (find != null)
                            {
                                contractID.Add(find);
                            }
                        }
                    }
                }

            }
            else
            {
                foreach (var it in userIds)
                {
                    var find = db.tbUserContracts.Where(s => s.FK_UserID == it && s.usc_StartTime <= az && s.usc_EndTime >= ta).OrderByDescending(s => s.usc_ID).FirstOrDefault();
                    if (find != null)
                    {
                        contractID.Add(find);
                    }
                }

            }






            foreach (var item in contractID)


            {
                var Result = rep_userContracts.Find(item.usc_ID);

                Result.tbUsers = context.tbUsers.Find(Result.FK_UserID);
                Result.jobgroup_ValueMozdGroup = Result.jobgroup_ValueMozdGroup.Replace(",", "");
                Result.jobgroup_ValueSanavat = Result.jobgroup_ValueSanavat.Replace(",", "");
                Result.jobgroup_HagheOlad = Result.jobgroup_HagheOlad.Replace(",", "");
                Result.jobgroup_HagheMaskan = Result.jobgroup_HagheMaskan.Replace(",", "");
                Result.jobgroup_KharoBar = Result.jobgroup_KharoBar.Replace(",", "");

                Fillexceltst9(Result, shamsiDateaz, shamsiDateta);
            }
            return Content("True");

        }
        public ActionResult Fillexceltst9(tbUserContracts Model, string shamsiDateaz, string shamsiDateta)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Contract form.xlsx"));
            var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/New Contract form.xlsx");

            var obj = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
            var birthplace = obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfBrith).Select(p => p.cityName).FirstOrDefault();
            var PlaceOfRegister = obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfIssue).Select(p => p.cityName).FirstOrDefault();
            var typeofcontract = "";
            string maragestatus = "";
            string DutySystem = "";
            string MadrakTahsili = "";

            if (Model.usc_TypeOfContract == false)
            {
                typeofcontract = "وقت مزدی";
            }
            else
            {
                typeofcontract = "کارمزدی";
            }

            if (Model.tbUsers.usr_MaritaIStatus == 1)
            {
                maragestatus = "متاهل";
            }
            else if (Model.tbUsers.usr_MaritaIStatus == 2)
            {
                maragestatus = "مجرد";
            }

            switch (Model.tbUsers.usr_Dutysystem)
            {
                case 1:
                    DutySystem = "مشمول";
                    break;
                case 2:
                    DutySystem = "پایان خدمت";
                    break;
                case 3:
                    DutySystem = "معافیت دائم";
                    break;
                case 4:
                    DutySystem = "معافیت موقت";
                    break;
                case 5:
                    DutySystem = "غایب";
                    break;
                case 6:
                    DutySystem = "خانم";
                    break;
                default:
                    DutySystem = "نامعلوم";
                    break;
            }

            if (Model.tbUsers.usr_Degree != null)
            {
                switch (Model.tbUsers.usr_Degree)
                {
                    case 1:
                        MadrakTahsili = "بیسواد";
                        break;
                    case 2:
                        MadrakTahsili = "سیکل";
                        break;
                    case 3:
                        MadrakTahsili = "دیپلم";
                        break;
                    case 4:
                        MadrakTahsili = "کاردانی";
                        break;
                    case 5:
                        MadrakTahsili = "کارشناسی";
                        break;
                    case 6:
                        MadrakTahsili = "کارشناسی ارشد";
                        break;
                    case 8:
                        MadrakTahsili = "دکتری";
                        break;
                    default:
                        MadrakTahsili = "نامعلوم";
                        break;
                }
            }
            else
            {
                MadrakTahsili = "نامعلوم";
            }

            try
            {

                using (var package = new ExcelPackage())
                {
                    using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
                    {
                        package.Load(stream);
                    }

                    var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);

                    if (worksheet2 != null)
                    {
                        var stringList = new List<string>();

                        #region fill string list
                        stringList.Add(Model.usc_ContractNumber);
                        stringList.Add(Model.tbUsers.usr_Family);
                        stringList.Add(Model.tbUsers.usr_Name);
                        stringList.Add(Model.tbUsers.usr_SHCode.ToString());
                        stringList.Add(Model.tbUsers.usr_FatherName);
                        stringList.Add(Model.tbUsers.shamsiDateOfBirth);
                        stringList.Add(birthplace);
                        stringList.Add(PlaceOfRegister);
                        stringList.Add(maragestatus);
                        stringList.Add(Model.tbUsers.usr_Child_Allowance.ToString());
                        stringList.Add(DutySystem);
                        stringList.Add(MadrakTahsili);
                        stringList.Add("---");
                        stringList.Add(Model.tbUsers.usr_NationalCode);
                        stringList.Add(Model.tbUsers.usr_Personal_ID.ToString());
                        stringList.Add(Model.tbCompanies.CompanyName);
                        stringList.Add(Model.tbCompanies.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies.Company_Address);
                        stringList.Add(Model.tbCompanies.RegistrationNumber.ToString());
                        stringList.Add(Model.tbCompanies.PlaceOfRegister.ToString());
                        stringList.Add(Model.usc_Jobtitle);
                        stringList.Add(Model.usc_JobCode.ToString());
                        stringList.Add(Model.FK_JobGroup.ToString());
                        stringList.Add(Model.tbCompanies1.CompanyName);
                        stringList.Add(Model.tbCompanies1.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies1.Company_Address);
                        stringList.Add(Model.usc_ShmasiStartTime);
                        stringList.Add(Model.usc_ShmasiENDTime);
                        var diff = Model.usc_EndTime - Model.usc_StartTime;
                        stringList.Add(diff.HasValue ? diff.Value.Days.ToString() : "0");
                        stringList.Add(typeofcontract);
                        //qwer
                        //stringList.Add(Model.jobgroup_ValueMozdGroup.ToString("N0"));
                        stringList.Add(System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup).ToString("N0"));
                        stringList.Add(System.Convert.ToInt32(Model.jobgroup_ValueSanavat).ToString("N0"));
                        stringList.Add("0");
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) + System.Convert.ToInt32(Model.jobgroup_ValueSanavat)).ToString("N0"));
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30).ToString("N0"));
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30).ToString("N0"));
                        stringList.Add("0");
                        stringList.Add(((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30) + (System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30)).ToString("N0"));
                        stringList.Add(System.Convert.ToInt32(Model.jobgroup_HagheMaskan).ToString("N0"));
                        stringList.Add(System.Convert.ToInt32(Model.jobgroup_KharoBar).ToString("N0"));
                        stringList.Add(System.Convert.ToInt32(Model.jobgroup_HagheOlad).ToString("N0"));
                  
                        stringList.Add("شرح وظیفه:"+Model.Sharhvazayf);
                        stringList.Add("توضیحات:" + Model.sharh);
                        stringList.Add("شرابط خاتمه قرارداد :" + Model.shraytfasgh);

                        #endregion

                        var gharar = Model.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == Model.usc_ID).ToList();
                        double valueee = 0;
                        foreach (var it in gharar)
                        {
                            if (it != null)
                            {
                                //stringList.Add(gharar.Value.ToString());
                                //stringList.Add(System.Convert.ToInt32(it.Value).ToString("N0"));
                                valueee += (double)it.Value;
                                //stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + gharar.Value + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());
                            }
                            else
                            {
                                stringList.Add("0");
                                //stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString("N0"));
                            }
                        }
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + valueee + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString("N0"));

                        foreach (var it in gharar)
                        {
                            if (it != null)
                            {
                                //stringList.Add(gharar.Value.ToString());
                                stringList.Add(System.Convert.ToInt32(it.Value).ToString("N0"));
                                //valueee +=(double) it.Value;
                                //stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + gharar.Value + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());
                            }
                            else
                            {
                                stringList.Add("0");
                                //stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString("N0"));
                            }
                        }
                        //var count1 = newFilePath.Sheets[1].Rows.Count();

                        int count = 45;
                        count = count + gharar.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            var cell = worksheet2.Cells[$"C{i}"];
                            cell.Value = stringList[i - 1];
                        }
                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }

                    var worksheet3 = package.Workbook.Worksheets[0];
                    if (worksheet3 != null)
                    {

                        worksheet3.Calculate();


                        var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address]
                            .Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();

                        if (cellsWithFormulas.Any())
                        {
                            foreach (var cell in cellsWithFormulas)
                            {
                                var value = cell.Value;
                                cell.Value = value;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
                        }

                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }
                    string namepyn = "";
                    string city = "";
                    var findpymn = db.Link_User_And_Peyman.Where(s => s.FK_User_ID == Model.tbUsers.usr_ID && s.Status == true).FirstOrDefault();
                    if (findpymn != null)
                    {
                        namepyn = findpymn.tbPeymanContracts.pec_Title;
                    }
                    var findcity = db.tbCities.Where(s => s.ID == Model.tbUsers.usr_City_Dutysystem).FirstOrDefault();
                    if (findcity != null)
                    {
                        city = findcity.Name;
                    }
                    SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                    GemBox.Spreadsheet.ExcelFile workbook1 = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);
                    string safeName = Model.tbUsers.FullName.Replace("/", "_").Replace("\\", "_");
                    string safeDateAz = shamsiDateaz.Replace("/", "_");
                    string safeDateTa = shamsiDateta.Replace("/", "_");
                    while (workbook1.Worksheets.Count > 1)
                        workbook1.Worksheets.Remove(1);
                    string outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/ghardad/" +
                        Model.tbUsers.usr_Personal_ID + "_" + safeName + "_"+ namepyn + "_"+ city + "_" + safeDateAz + "_" + safeDateTa + ".pdf");
                    workbook1.Save(outputPdfPath);
                    System.Diagnostics.Debug.WriteLine("تبدیل به PDF با موفقیت انجام شد.");
                    return Content("True");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("خطا در تبدیل فایل: " + ex.Message);
                System.Diagnostics.Debug.WriteLine("جزئیات خطا: " + ex.StackTrace);
                return Content("یک خطا رخ داده است: " + ex.Message);
            }
        }







        public ActionResult vi1(List<string> Fkpym, int month, int year, int count)
        {
            try
            {
                count = 0;
                {
                    var userIds = Fkpym.Where(x => int.TryParse(x, out _)) // فقط مقادیر عددی معتبر
                         .Select(int.Parse).ToList();

                    if(userIds.Count == 1)
                    {
                        foreach(var it in userIds)
                        {
                            if (it ==-1)
                            {
                                var ex = db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)" 
                                 && p.tbUsers.usr_amani != true && p.mlfvlfsh_Month == month && p.mlfvlfsh_Year == year).OrderBy(p => p.mlfvlfsh_ID)
                                  .Skip(count).ToList();
                                foreach (var item in ex)
                                {
                                    count += 1;

                                    var result = ShowFish3332(item.tbUsers.usr_ID, month, year, count);

                                    //if (result == -1)
                                    //{
                                    //    ViewBag.Month = month;
                                    //    ViewBag.Year = year;
                                    //    ViewBag.CountPdf = count;
                                    //    return View("~/Areas/Contracts/Views/UserContract/viewcontinue.cshtml");

                                    //}
                                }

                            }
                            else
                            {
                                var ex = db.tbMoalefeValueFish
                  .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)"
                     && userIds.Contains(p.FK_User.Value) // فرض بر این که FK_User از نوع int? هست
                      && p.tbUsers.usr_amani != true
                       && p.mlfvlfsh_Month == month
                        && p.mlfvlfsh_Year == year)
                         .OrderBy(p => p.mlfvlfsh_ID)
                           .Skip(count)
                            .ToList();
                                foreach (var item in ex)
                                {
                                    count += 1;

                                    var result = ShowFish3332(item.tbUsers.usr_ID, month, year, count);

                                    //if (result == -1)
                                    //{
                                    //    ViewBag.Month = month;
                                    //    ViewBag.Year = year;
                                    //    ViewBag.CountPdf = count;
                                    //    return View("~/Areas/Contracts/Views/UserContract/viewcontinue.cshtml");

                                    //}
                                }
                            }
                        }

                    }
                    else
                    {
                        var ex = db.tbMoalefeValueFish
                        .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)"
                           && userIds.Contains(p.FK_User.Value) // فرض بر این که FK_User از نوع int? هست
                            && p.tbUsers.usr_amani != true
                             && p.mlfvlfsh_Month == month
                              && p.mlfvlfsh_Year == year)
                               .OrderBy(p => p.mlfvlfsh_ID)
                                 .Skip(count)
                                  .ToList();
                        foreach (var item in ex)
                        {
                            count += 1;

                            var result = ShowFish3332(item.tbUsers.usr_ID, month, year, count);

                            //if (result == -1)
                            //{
                            //    ViewBag.Month = month;
                            //    ViewBag.Year = year;
                            //    ViewBag.CountPdf = count;
                            //    return View("~/Areas/Contracts/Views/UserContract/viewcontinue.cshtml");

                            //}
                        }

                    }



                    //var ex = db.tbMoalefeValueFish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)" && Fkpym.Contains(p.FK_User.Value)
                    // && p.tbUsers.usr_amani != true && p.mlfvlfsh_Month == month && p.mlfvlfsh_Year == year).OrderBy(p => p.mlfvlfsh_ID)
                    //  .Skip(count).ToList();



                }
            }
            catch (Exception ex)

            {
                return Content("Error: " + ex.Message);
            }
            //else
            //{
            //    ShowFish333(usr_ID, month, year, 1);

            //}

            return Content("True");
        }




        /// <summary>
        /// تشخیص میدهد قرار داد ارسالی با قرار های دیگر این فردهمپوشانی نداشته باشد
        /// </summary>
        /// <returns></returns>
        [AuthorizeAAA]
        public bool CheckForOverLapConracts(DateTime? startDate, DateTime? endDate, int? user_ID)
        {
            var itsContractUser = rep_userContracts.Listt(user_ID);
            if (itsContractUser != null)
            {
                foreach (var item in itsContractUser)
                {
                    if (startDate >= endDate)
                        return false;
                    if (startDate >= item.usc_StartTime && startDate <= item.usc_EndTime) // زمان شروع بین قرارداد قبلی است
                        return false;
                    if (startDate < item.usc_StartTime && endDate >= item.usc_StartTime && endDate <= item.usc_EndTime)
                        return false;

                    return true;
                }
                return true;
            }
            else
                return true;
        }
        ////http://localhost:6061/Contracts/UserContract/vi?usr_ID=0&year=1404&month=1&count=0
        public ActionResult vi(int usr_ID,int month,int year,int count)
        {
            if (usr_ID== 0)
            {
                var ex = db.tbMoalefeValueFish
       .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)"
                && p.tbUsers.usr_amani != true
                && p.mlfvlfsh_Month == month
                && p.mlfvlfsh_Year == year)
       .OrderBy(p => p.mlfvlfsh_ID) // حتماً ترتیب بده تا Skip معنی داشته باشه
       .Skip(count)
       .ToList();
                var ex2 = db.tbMoalefeValueFish
  .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)"
           && p.tbUsers.usr_amani != true
           && p.mlfvlfsh_Month == month
           && p.mlfvlfsh_Year == year)
  //.OrderBy(p => p.mlfvlfsh_ID) // حتماً ترتیب بده تا Skip معنی داشته باشه
  //.Skip(count)
  .ToList();

                //count = 0;
                foreach (var item in ex)
                {
                    //if (db.tbcheackfishforusers.Where(p => p.FK_User == item.FK_User && p.Month == month && p.Year == year && p.HasFish == true).FirstOrDefault() != null)
                    //{
                    count += 1;

                    int result = ShowFish333(item.tbUsers.usr_ID, month, year, count);


                    // بررسی نتیجه‌ی بازگشتی از تابع
                    if (result  ==-1)
                    {
                        ViewBag.Month = month;
                        ViewBag.Year = year;
                        ViewBag.CountPdf = count;
                        return View("~/Areas/Contracts/Views/UserContract/viewcontinue.cshtml");
                        // اگر تابع موفق بود

                    }
                 
                    //}


                }
                
            }
            else
            {
                var findusr = db.tbUsers.Where(s => s.usr_Personal_ID == usr_ID).FirstOrDefault();
                ShowFish333(findusr.usr_ID, month, year,1);

            }

            return View();
        }

        public async Task<ActionResult> DeleteDuplicateFish(int usr_ID, int month, int year)
        {
            bool hasDuplicates = true;

            while (hasDuplicates)
            {
                // پیدا کردن رکوردها با شرایط مشخص (ماه خاص)
                var records = await db.tbMoalefeValueFish
                    .Where(s => s.mlfvlfsh_Month == 11 && s.FK_EXCel != null)
                    .ToListAsync();

                // گروه‌بندی بر اساس FK_Moalefe، mlfvlfsh_Month، FK_User و FK_EXCel
                var groupedDuplicates = records
                    .GroupBy(s => new { s.FK_Moalefe, s.mlfvlfsh_Month, s.FK_User, s.FK_EXCel })
                    .Where(g => g.Count() > 1); // فقط گروه‌هایی که بیشتر از یک رکورد دارند

                // بررسی اینکه آیا گروه تکراری وجود دارد یا نه
                var firstDuplicateGroup = groupedDuplicates.FirstOrDefault();
                if (firstDuplicateGroup != null)
                {
                    var firstRecord = firstDuplicateGroup.FirstOrDefault();
                    if (firstRecord != null)
                    {
                        db.tbMoalefeValueFish.Remove(firstRecord);
                        await db.SaveChangesAsync();
                    }
                }
                else
                {
                    // اگر هیچ گروه تکراری پیدا نشد، حلقه متوقف می‌شود
                    hasDuplicates = false;
                }
            }

            return Content("همه رکوردهای تکراری با موفقیت حذف شدند");
        }
        public ActionResult DeleteDuplicateFish1(int usr_ID, int month, int year)
        {
            // پیدا کردن رکوردها با شرایط مشخص (ماه خاص)
            var records = db.tbmoalfefishexcel
                .Where(s => s.tbMoalefeValueFish.Any(d => d.mlfvlfsh_Month == 11 &&d.Finalaccept==true))
                .ToList();

            foreach (var record in records)
            {
                // پیدا کردن رکوردهایی که FK_Moalefe برابر با 1960 دارند
                var duplicates = db.tbMoalefeValueFish
                    .Where(s => s.FK_EXCel == record.ID)
                    .ToList();

                // گروه‌بندی بر اساس FK_Moalefe، mlfvlfsh_Month و FK_User
                var groupedDuplicates = duplicates
                    .GroupBy(s => new { s.FK_Moalefe, s.mlfvlfsh_Month, s.FK_User })
                    .Where(g => g.Count() > 1); // فقط گروه‌هایی که بیشتر از یک رکورد دارند

                foreach (var group in groupedDuplicates)
                {
                    // حذف اولین رکورد از گروه
                    var firstRecord = group.FirstOrDefault();
                    if (firstRecord != null)
                    {
                        db.tbMoalefeValueFish.Remove(firstRecord);
                        db.SaveChanges();
                        break;

                    }
                }
            }

            // ذخیره تغییرات در دیتابیس

            return Content("حذف رکوردهای تکراری با موفقیت انجام شد");
        }


        public int ShowFish333(int UserID, int Month, int Year,int count)
        {
            try
            {
                var x = db.tbVisibleFish.FirstOrDefault(p => p.Month == Month && p.Year == Year);
                if (x != null)
                {
                    if (x.IsVisibile)
                    {
                        var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);
                        ManualFishDetail Model = new ManualFishDetail();
                        var lstmoalefe = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();

                        if (Year >= 1404)
                        {
                            lstmoalefe = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList();


                        }
                        #region فیش دارید یا خیر
                        if (Year  <1404)
                        {
                            var IDkhales = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                            if (!db.tbMoalefeValueFish.Any(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == IDkhales && p.FK_User == UserID))
                            {
                                return -1;
                            }

                        }
                        else
                        {
                            var IDkhales = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                            if (!db.tbMoalefeValueFish.Any(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == IDkhales && p.FK_User == UserID&&p.FK_EXCel==null))
                            {
                                return -1;
                            }
                        }
                     
                        #endregion
                        Model.month = Month;
                        var user = db.tbUsers.FirstOrDefault(p => p.usr_ID == UserID);
                        int peymanID = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == UserID && p.Status == true).FK_Peyman_ID.Value;
                        var t = db.FinancialDocuments.Where(p => p.User_ID == UserID).ToList();
                        List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
                        var asnad = db.tbfkfinancial.ToList();

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



                        long value = 0;
                        long value2 = 0;


                        foreach (var item2 in asnadstr)
                        {
                            var t3 = db.FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();

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
                                            value += (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value2 += (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value += (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }
                        }


























                        var title = db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl != null).FirstOrDefault();
                        string job = "";
                        if (title != null)
                        {
                            job = title.tbjob.Name;
                        }
                        else
                        {
                            job = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID).usc_Jobtitle;
                        }
                        FishHeader fishHeader = new FishHeader
                        {
                            PeymanName = db.tbPeymanContracts.FirstOrDefault(p => p.pec_ID == peymanID && p.Inactive != true).pec_Title,

                        
                            OnvanShoql = job,

                     
                        TedadRoozMonth = fishuti.TeadaroozMonth(Month, Year),
                            MonthName = fishuti.MonthName(Month),
                            PersonalCode = user.usr_Personal_ID.Value,
                            UserName = user.FullName,
                            NationalCode = user.usr_NationalCode,
                            year = Year,

                            VALUE = value,
                            VALUE2 = value2
                        };
                        Model.FishHeader = fishHeader;
                        #region مولفه داینامیک قرارداد و ریختن در فیش ولیو
                        var contractinfo = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_EndTime >= myDatetime && p.usc_StartTime <= myDatetime).FirstOrDefault();
                        var listmoalefeqarardadi = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == contractinfo.usc_ID).ToList();
                        List<FishValue> lstmoalefeqarardadfish = new List<FishValue>();
                        foreach (var item in listmoalefeqarardadi)
                        {
                            var moalefeObj = db.tbContractMoalefeDastmozdi
             .FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi);
                            FishValue moalefeqaradadfish = new FishValue
                            {
                                Title = moalefeObj?.md_Title ?? "",
                                Value = item.Value.HasValue ? (int)item.Value.Value : 0
                                //Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi).md_Title,
                                //Value = (int)item.Value
                            };
                            lstmoalefeqarardadfish.Add(moalefeqaradadfish);

                        }

                        Model.FishValueQaradad = lstmoalefeqarardadfish;
                        #endregion
                        #region پر کردن اضافات و ریختن در فیش ولیو
                        var ezafat = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
                        List<FishValue> ezafatValue = new List<FishValue>();
                        foreach (var item in ezafat)
                        {
                            var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID);
                            if (Year >= 1404)
                            {
                                 varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID&&p.FK_EXCel==null);

                            }
                            FishValue ezafvalue = new FishValue
                            {
                                Title = item.md_Title ?? "",
                                Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
                            };
                            ezafatValue.Add(ezafvalue);

                        }
                        Model.FishValuesEzafat = ezafatValue;
                        #endregion
                        #region  پر کردن کسورات و ریختن در فیش ولیو
                        var kosoratlist = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 1 && p.IsMain != true).ToList();
                        List<FishValue> koosratvalue = new List<FishValue>();
                        foreach (var item in kosoratlist)
                        {
                            var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID);
                            if (Year >= 1404)
                            {
                                 varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID && p.FK_EXCel == null);


                            }
                            FishValue kasrvalue = new FishValue
                            {
                                Title = item.md_Title ?? "",
                                Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
                            };
                            koosratvalue.Add(kasrvalue);

                        }
                        Model.FishValuesKosoorat = koosratvalue;
                        #endregion
                        #region محاسبه اقساط و ریختن در فیش ولیو
                        var tamamaqsat = fishuti.CalculateAqsat(Year, Month, UserID);
                        List<FishValue> aqsatformohasebekasriha = new List<FishValue>();
                        foreach (var item in tamamaqsat)
                        {
                            FishValue aqsat = new FishValue
                            {
                                Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.Item2).md_Title,
                                Value = item.Item1
          
                            };
                            aqsatformohasebekasriha.Add(aqsat);

                        }
                        //TODO:Felan badan bayad pakshavad
                        //FishValue felan = new FishValue
                        //{
                        //    Title = "صندوق خیریه شرکت",
                        //    Value = 200000
                        //};
                        //aqsatformohasebekasriha.Add(felan);
                        ///
                        Model.FishValuesAqsat = aqsatformohasebekasriha;

                        #endregion


                        #region فیش ولیوهای اصلی

                        List<FishValue> lstfshvalue = new List<FishValue>();


                        foreach (var item in lstmoalefe)
                        {
                            FishValue fshvalue = new FishValue
                            {
                                Value = item.mlfvlfsh_Value != null ? (double)item.mlfvlfsh_Value : 0,
                                Title = item.tbContractMoalefeDastmozdi != null
                                            ? item.tbContractMoalefeDastmozdi.md_Title
                                            : ""
                            };
                            lstfshvalue.Add(fshvalue);
                        }

                        Model.FishValues = lstfshvalue;
                        //excelfish(Model, UserID, Month, Year);
                         //excelfishtest1(Model, UserID, Month, Year, count);
                        return 0;

                        // بررسی نتیجه‌ی بازگشتی از تابع

                        // اگر تابع موفق بود
                        /*    return 0; */ // موفقیت


                        #endregion
                        //return Content("0");
                    }
                    else
                    {
                        return 0;
                    }
                }

                else
                {
                    return 0;
                }


            }
            catch(Exception ex)
            {
                string rx = ex.Message;
                return -2;//در نمایش فیش خطایی رخ داده است

            }
          

        }


        //------
        //public int excelfishtest12(ManualFishDetail Model, int UserID, int Month, int Year, int countpdf)
        //{
        //    var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();
        //    var caran = db.tbCaranSettings.ToList();
        //    var final = db.FinancialDocuments.ToList();
        //    var moalf = db.tbContractMoalefeDastmozdi.ToList();
        //    var fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();
        //    if (Year >= 1404)
        //    {
        //        fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList();
        //    }
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan1.xlsx"));
        //    var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
        //    var asnad = db.tbfkfinancial.ToList();
        //    List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
        //    foreach (var it in asnad)
        //    {
        //        PersianCalendar persianCalendar = new PersianCalendar();
        //        DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //        int persianYear = persianCalendar.GetYear(dataDocument);
        //        int persianMonth = persianCalendar.GetMonth(dataDocument);
        //        if (persianYear == Year && persianMonth == Month)
        //        {
        //            asnadstr.Add(it);
        //        }
        //    }
        //    long value = 0;
        //    long value2 = 0;
        //    long value3 = 0;
        //    long value4 = 0;
        //    long value5 = 0;
        //    long value6 = 0;
        //    long value344 = 0;
        //    long ezafehvalue6001 = 0; long kamvalue6001 = 0;
        //    long ezafehvalue7001 = 0; long kamvalue7001 = 0;

        //    foreach (var item2 in asnadstr)
        //    {
        //        var t3 = final.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();
        //        foreach (var item1 in t3)
        //        {
        //            if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
        //            {
        //                PersianCalendar persianCalendar = new PersianCalendar();
        //                DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //                int persianYear = persianCalendar.GetYear(dataDocument);
        //                int persianMonth = persianCalendar.GetMonth(dataDocument);
        //                if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
        //                {
        //                    if (item1.Debtore != 0 && item1.Creditor != 0)
        //                    {
        //                        value += (long)(item1.Creditor - item1.Debtore);
        //                    }
        //                    else if (item1.Debtore != 0 && item1.Debtore != null)
        //                    {
        //                        value2 = (long)(item1.Debtore ?? 0);
        //                    }
        //                    else if (item1.Creditor != 0 && item1.Creditor != null)
        //                    {
        //                        value += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value3 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value344 += (long)(item1.Debtore ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5002)
        //                    {
        //                        value4 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value344 += (long)(item1.Debtore ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 6001)
        //                    {
        //                        ezafehvalue6001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 6001)
        //                    {
        //                        kamvalue6001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 7001)
        //                    {
        //                        ezafehvalue7001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 7001)
        //                    {
        //                        kamvalue7001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5003)
        //                    {
        //                        value6 += (long)(item1.Creditor ?? 0);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    try
        //    {
        //        using (var package = new ExcelPackage())
        //        {

        //            using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
        //            {
        //                package.Load(stream);
        //            }
        //            // Get the second sheet of the new file if it exists
        //            var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);
        //            // Check if the worksheet exists
        //            if (worksheet2 != null)
        //            {
        //                // Define a list of strings to be added to the cells in columns C1 to C42
        //                var stringList = new List<string>();
        //                var stringList2 = new List<string>();
        //                var stringList3 = new List<string>();
        //                var stringList4 = new List<string>();
        //                var stringList5 = new List<string>();
        //                #region fill string list
        //                stringList.Add(Model.FishHeader.PersonalCode.ToString(""));
        //                stringList.Add(Model.FishHeader.OnvanShoql);
        //                stringList.Add(Model.FishHeader.NationalCode);
        //                stringList.Add(Model.FishHeader.UserName.ToString());
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد شغل (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کار ( هرساعت-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع مزد مبنا ( روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishHeader.MonthName);
        //                stringList.Add(Model.FishHeader.year.ToString());
        //                stringList.Add(Model.FishHeader.TedadRoozMonth.ToString("N0"));
        //                var endDate = new DateTime(2025, 3, 20);
        //                var finddd = db.Link_User_And_Peyman.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.Status == true).FirstOrDefault();
        //                var findob = db.tbUserContracts.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.usc_EndTime == endDate && p.FK_KarfarmaID != null).FirstOrDefault();
        //                string peymanName2 = Model.FishHeader.PeymanName;
        //                if (findob != null)
        //                {
        //                    peymanName2 = findob.tbCompanies.CompanyName;
        //                }
        //                if (finddd != null)
        //                {
        //                    var finj = db.tbPeymanContracts.Where(p => p.FK_UserTarafDovvom != null && p.pec_ID == finddd.FK_Peyman_ID).FirstOrDefault();
        //                    if (finj != null)
        //                    {
        //                        peymanName2 = db.tbCompanies.Where(p => p.ID == finj.FK_UserTarafDovvom).Select(s => s.CompanyName).FirstOrDefault();
        //                    }
        //                }
        //                // جداکننده مورد نظر
        //                char separator2 = '-';
        //                // یافتن موقعیت جداکننده در رشته
        //                int startIndex2 = peymanName2.IndexOf(separator2);
        //                string extractedFileName = peymanName2.Substring(startIndex2 + 1).Trim();
        //                stringList.Add(extractedFileName.ToString());
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز کارکرد ").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد ساعات اضافه کار").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز ماموریت").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز استعلاجی").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز غیبت").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد گروه (شغل)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق تاهل").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ماموریت (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
        //                {
        //                    stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)").Value.ToString("N0"));
        //                }
        //                else
        //                {
        //                    stringList3.Add("0");
        //                }
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "پاداش نوع 2").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع ناخالص حقوق و مزایا (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول بیمه (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول مالیات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)") != null)
        //                {
        //                    stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)").Value.ToString("N0"));
        //                }
        //                else
        //                {
        //                    stringList4.Add("0");
        //                }
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جریمه نوع 2").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کار مازاد (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل کسورات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "خالص قابل دریافت (ریال)").Value.ToString("N0"));
        //                var ex333 = db.tbJaremehandpadash.Where(p => p.Fk_usr == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
        //                string title = "";
        //                if (ex333 != null)
        //                {
        //                    title = ex333.Titel;
        //                }
        //                var tb = noalf.Where(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Value != 0).ToList();
        //                List<string> mode = new List<string>();
        //                foreach (var it in tb)
        //                {
        //                    var yuuu = caran
        //                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
        //                        .FirstOrDefault();

        //                    if (yuuu == null)
        //                    {
        //                        var u = caran
        //                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
        //                            .FirstOrDefault();

        //                        if (u != null)
        //                        {
        //                            mode.Add(
        //                                $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
        //                                $"( {(u.CaranAyabOZahab ?? 0).ToString("N0")}/ 0 /{(it.MoalfeVal_Value ?? 0)} )----"
        //                            );
        //                        }
        //                    }
        //                    else
        //                    {
        //                        mode.Add(
        //                            $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
        //                            $"( {(yuuu.CaranAyabOZahab ?? 0).ToString("N0")}/ {(yuuu.CaranStandard ?? 0)} /{(it.MoalfeVal_Value ?? 0)} )---"
        //                        );
        //                    }
        //                }

        //                //foreach (var it in tb)
        //                //{
        //                //    var yuuu = caran.Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0).FirstOrDefault();
        //                //    if (yuuu == null)
        //                //    {
        //                //        var u = caran.Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0).FirstOrDefault();
        //                //        if (u != null)
        //                //        {
        //                //            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)u.CaranAyabOZahab).ToString("N0")}/ 0 /{it.MoalfeVal_Value} )----");
        //                //        }
        //                //    }
        //                //    else
        //                //    {
        //                //        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)yuuu.CaranAyabOZahab).ToString("N0")}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
        //                //    }
        //                //}
        //                if (db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.FKUser == UserID && p.Year == Year && p.RemainValue != 0 && p.RemainValue != null).FirstOrDefault() != null)
        //                {
        //                    mode.Add($"(مازاد اضافه کار در حال بررسی میباشد  )----");
        //                }
        //                var List = worksheet2.Rows;
        //                var count = worksheet2.Rows[0].Count();
        //                var count2 = worksheet2.Rows.Count();
        //                int rowCount = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 1].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 1].Value.ToString()))
        //                    {
        //                        rowCount++;
        //                    }
        //                }
        //                int rowCount2 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 3].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 3].Value.ToString()))
        //                    {
        //                        rowCount2++;
        //                    }
        //                }

        //                int rowCount3 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 5].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 5].Value.ToString()))
        //                    {
        //                        rowCount3++;
        //                    }
        //                }
        //                int rowCount4 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 9].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 9].Value.ToString()))
        //                    {
        //                        rowCount4++;
        //                    }
        //                }
        //                var countColumnARows = rowCount;
        //                for (int i = 13; i <= countColumnARows; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 1].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                        else
        //                        {
        //                            stringList.Add("0");
        //                        }
        //                    }
        //                }
        //                #endregion

        //                for (int i = 6; i <= rowCount2; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 3].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList2.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                for (int i = 16; i <= rowCount3; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 5].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList3.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                for (int i = 10; i <= rowCount4; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 9].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                // Loop through each row in columns C1 to C42 and add the corresponding string from the list
        //                for (int i = 1; i <= stringList.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"B{i}"];
        //                    cell.Value = stringList[i - 1];
        //                }
        //                for (int i = 1; i <= stringList2.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"D{i}"];
        //                    cell.Value = stringList2[i - 1];
        //                }
        //                for (int i = 1; i <= stringList3.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"F{i}"];
        //                    cell.Value = stringList3[i - 1];
        //                }
        //                for (int i = 1; i <= stringList4.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"J{i}"];
        //                    cell.Value = stringList4[i - 1];
        //                }
        //                var cell2 = worksheet2.Cells[$"L{1}"];
        //                cell2.Value = string.Join(" ", mode);
        //                var cell3 = worksheet2.Cells[$"L{4}"];
        //                cell3.Value = value3;
        //                var cell32 = worksheet2.Cells[$"L{9}"];
        //                cell32.Value = value344;

        //                var cell10 = worksheet2.Cells[$"L{10}"];
        //                cell10.Value = ezafehvalue6001;
        //                var cell11 = worksheet2.Cells[$"L{11}"];
        //                cell11.Value = kamvalue6001;

        //                var cell12 = worksheet2.Cells[$"L{12}"];
        //                cell12.Value = ezafehvalue7001;
        //                var cell13 = worksheet2.Cells[$"L{13}"];
        //                cell13.Value = kamvalue7001;


        //                var cell4 = worksheet2.Cells[$"L{5}"];
        //                cell4.Value = value5;
        //                var cell5 = worksheet2.Cells[$"L{6}"];
        //                cell5.Value = value6;
        //                var cell6 = worksheet2.Cells[$"L{7}"];
        //                cell6.Value = value4;
        //                var cell7 = worksheet2.Cells[$"L{8}"];
        //                cell7.Value = value2;
        //                var cell8 = worksheet2.Cells[$"L{2}"];
        //                cell8.Value = title;
        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //                // استفاده از فضای خالی به جای Environment.NewLine
        //            }
        //            var worksheet3 = package.Workbook.Worksheets[0];
        //            if (worksheet3 != null)
        //            {
        //                worksheet3.Calculate();
        //                var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address].Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();
        //                if (cellsWithFormulas.Any())
        //                {
        //                    foreach (var cell in cellsWithFormulas)
        //                    {
        //                        var value1232 = cell.Value;
        //                        cell.Value = value1232;
        //                    }
        //                }
        //                else
        //                {
        //                    System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
        //                }
        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //            }
        //            // Save changes to the new file
        //        }
        //        using (ExcelEngine excelEngine = new ExcelEngine())
        //        {

        //            string peymanName = Model.FishHeader.PeymanName;
        //            string extractedFileName = peymanName;
        //            //Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(newFilePath);
        //            //for (int i = 1; i < workbook.Worksheets.Count; i++)
        //            //{
        //            //    workbook.Worksheets[i].IsVisible = false;
        //            //}
        //            string fullnam = "";
        //            var findusr = db.tbUsers.Where(s => s.usr_Personal_ID == Model.FishHeader.PersonalCode).FirstOrDefault();
        //            if (findusr != null)
        //            {
        //                fullnam = findusr.FullName;
        //            }
        //            string outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + fullnam + '-' + Month + '-' + Year + "_Clean.pdf");
        //            string pdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + fullnam + '-' + Month + '-' + Year + ".pdf");
        //            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        //            try
        //            {
        //                // بارگذاری فایل اکسل با GemBox
        //                GemBox.Spreadsheet.ExcelFile workbook1 = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);

        //                // فقط شیت اول را نگه دار، بقیه را حذف کن (چون در Aspose پنهان می‌کردی)
        //                while (workbook1.Worksheets.Count > 1)
        //                    workbook1.Worksheets.Remove(1);

        //                // تبدیل به PDF
        //                workbook1.Save(outputPdfPath);
        //            }
        //            catch (Exception ex)
        //            {
        //                // ثبت خطا یا ادامه کار
        //            }

        //            //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
        //            //workbook.Dispose();

        //            try
        //            {
        //                //workbook.Save(outputPdfPath);
        //                //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
        //                //workbook.Dispose();
        //            }
        //            catch (Exception ex)
        //            {
        //            }
        //            //RemoveWatermarkFromPDF(pdfPath, outputPdfPath);
        //            System.Diagnostics.Debug.WriteLine("واترمارک با موفقیت حذف شد!");




        //            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPdfPath);
        //            return 0;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        viewcontinue(countpdf);
        //        return -1;
        //    }
        //}
        //------
        public ActionResult ShowFish3332(int UserID, int Month, int Year, int count)
        {
            try
            {
                var x = db.tbVisibleFish.FirstOrDefault(p => p.Month == Month && p.Year == Year);
                if (x != null)
                {
                    if (x.IsVisibile)
                    {
                        var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);
                        ManualFishDetail Model = new ManualFishDetail();
                        var lstmoalefe = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();

                        if (Year >= 1404)
                        {
                            lstmoalefe = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList();


                        }
                        #region فیش دارید یا خیر
                        if (Year < 1404)
                        {
                            var IDkhales = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                            if (!db.tbMoalefeValueFish.Any(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == IDkhales && p.FK_User == UserID))
                            {
                                return Content("-1");
                            }

                        }
                        else
                        {
                            var IDkhales = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                            if (!db.tbMoalefeValueFish.Any(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == IDkhales && p.FK_User == UserID && p.FK_EXCel == null))
                            {
                                return Content("-1");
                            }
                        }

                        #endregion
                        Model.month = Month;
                        var user = db.tbUsers.FirstOrDefault(p => p.usr_ID == UserID);
                        int peymanID = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == UserID && p.Status == true).FK_Peyman_ID.Value;
                        var t = db.FinancialDocuments.Where(p => p.User_ID == UserID).ToList();
                        List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
                        var asnad = db.tbfkfinancial.ToList();

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



                        long value = 0;
                        long value2 = 0;


                        foreach (var item2 in asnadstr)
                        {
                            var t3 = db.FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();

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
                                            value += (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value2 += (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value += (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }
                        }


























                        var title = db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl != null).FirstOrDefault();
                        string job = "";
                        if (title != null)
                        {
                            job = title.tbjob.Name;
                        }
                        else
                        {
                            job = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID).usc_Jobtitle;
                        }
                        FishHeader fishHeader = new FishHeader
                        {
                            PeymanName = db.tbPeymanContracts.FirstOrDefault(p => p.pec_ID == peymanID && p.Inactive != true).pec_Title,


                            OnvanShoql = job,


                            TedadRoozMonth = fishuti.TeadaroozMonth(Month, Year),
                            MonthName = fishuti.MonthName(Month),
                            PersonalCode = user.usr_Personal_ID.Value,
                            UserName = user.FullName,
                            NationalCode = user.usr_NationalCode,
                            year = Year,

                            VALUE = value,
                            VALUE2 = value2
                        };
                        Model.FishHeader = fishHeader;
                        #region مولفه داینامیک قرارداد و ریختن در فیش ولیو
                        var contractinfo = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_EndTime >= myDatetime && p.usc_StartTime <= myDatetime).FirstOrDefault();
                        var listmoalefeqarardadi = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == contractinfo.usc_ID).ToList();
                        List<FishValue> lstmoalefeqarardadfish = new List<FishValue>();
                        foreach (var item in listmoalefeqarardadi)
                        {
                            var moalefeObj = db.tbContractMoalefeDastmozdi
             .FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi);
                            FishValue moalefeqaradadfish = new FishValue
                            {
                                Title = moalefeObj?.md_Title ?? "",
                                Value = item.Value.HasValue ? (int)item.Value.Value : 0
                                //Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi).md_Title,
                                //Value = (int)item.Value
                            };
                            lstmoalefeqarardadfish.Add(moalefeqaradadfish);

                        }

                        Model.FishValueQaradad = lstmoalefeqarardadfish;
                        #endregion
                        #region پر کردن اضافات و ریختن در فیش ولیو
                        var ezafat = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
                        List<FishValue> ezafatValue = new List<FishValue>();
                        foreach (var item in ezafat)
                        {
                            var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID);
                            if (Year >= 1404)
                            {
                                varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID && p.FK_EXCel == null);

                            }
                            FishValue ezafvalue = new FishValue
                            {
                                Title = item.md_Title ?? "",
                                Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
                            };
                            ezafatValue.Add(ezafvalue);

                        }
                        Model.FishValuesEzafat = ezafatValue;
                        #endregion
                        #region  پر کردن کسورات و ریختن در فیش ولیو
                        var kosoratlist = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 1 && p.IsMain != true).ToList();
                        List<FishValue> koosratvalue = new List<FishValue>();
                        foreach (var item in kosoratlist)
                        {
                            var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID);
                            if (Year >= 1404)
                            {
                                varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID && p.FK_EXCel == null);


                            }
                            FishValue kasrvalue = new FishValue
                            {
                                Title = item.md_Title ?? "",
                                Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
                            };
                            koosratvalue.Add(kasrvalue);

                        }
                        Model.FishValuesKosoorat = koosratvalue;
                        #endregion
                        #region محاسبه اقساط و ریختن در فیش ولیو
                        var tamamaqsat = fishuti.CalculateAqsat(Year, Month, UserID);
                        List<FishValue> aqsatformohasebekasriha = new List<FishValue>();
                        foreach (var item in tamamaqsat)
                        {
                            FishValue aqsat = new FishValue
                            {
                                Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.Item2).md_Title,
                                Value = item.Item1

                            };
                            aqsatformohasebekasriha.Add(aqsat);

                        }
                        //TODO:Felan badan bayad pakshavad
                        //FishValue felan = new FishValue
                        //{
                        //    Title = "صندوق خیریه شرکت",
                        //    Value = 200000
                        //};
                        //aqsatformohasebekasriha.Add(felan);
                        ///
                        Model.FishValuesAqsat = aqsatformohasebekasriha;

                        #endregion


                        #region فیش ولیوهای اصلی

                        List<FishValue> lstfshvalue = new List<FishValue>();


                        foreach (var item in lstmoalefe)
                        {
                            FishValue fshvalue = new FishValue
                            {
                                Value = item.mlfvlfsh_Value != null ? (double)item.mlfvlfsh_Value : 0,
                                Title = item.tbContractMoalefeDastmozdi != null
                                            ? item.tbContractMoalefeDastmozdi.md_Title
                                            : ""
                            };
                            lstfshvalue.Add(fshvalue);
                        }
                        //مهم مهم 

                        Model.FishValues = lstfshvalue;
                        //excelfish(Model, UserID, Month, Year);
                        return excelfishtest1(Model, UserID, Month, Year, count);

                        // بررسی نتیجه‌ی بازگشتی از تابع

                        // اگر تابع موفق بود
                        /*    return 0; */ // موفقیت

                        //return Content("0");

                        #endregion
                        //return Content("0");
                    }
                    else
                    {
                        return Content("0");

                    }
                }

                else
                {
                    return Content("0");

                }


            }
            catch (Exception ex)
            {
                string rx = ex.Message;
                return Content(ex.Message);

            }


        }
        public ActionResult DownloadSingleFish(int userId, int year, int month)
        {
            string folderPath = Server.MapPath("~/Areas/Contracts/Contents/fish/");

            var user = db.tbUsers.Find(userId);
            if (user == null) return Content("کاربر یافت نشد");

            string pattern = $"{user.usr_Personal_ID}-*-{month}-{year}_Clean.pdf";

            var file = Directory.GetFiles(folderPath, pattern).FirstOrDefault();
            if (file == null) return Content("فیش یافت نشد");

            byte[] bytes = System.IO.File.ReadAllBytes(file);
            return File(bytes, "application/pdf", Path.GetFileName(file));
        }

        //#فیش  قیش
        //public ActionResult DownloadFishByMonth(int year, int month)
        //{
        //    string basePath = Server.MapPath("~/Areas/Contracts/Contents/fish/");

        //    string monthFolder = Path.Combine(basePath, year + "-" + month.ToString("00"));

        //    if (!Directory.Exists(monthFolder))
        //        return Content("پوشه مربوط به این ماه وجود ندارد");

        //    // جستجو داخل تمام زیرپوشه‌ها
        //    var files = Directory.GetFiles(
        //        monthFolder,
        //        "*_Clean.pdf",
        //        SearchOption.AllDirectories
        //    );

        //    if (files.Length == 0)
        //        return Content("فایلی برای این ماه یافت نشد");

        //    string zipPath = Path.Combine(basePath, $"Fish_{year}_{month}.zip");

        //    if (System.IO.File.Exists(zipPath))
        //        System.IO.File.Delete(zipPath);

        //    using (FileStream fs = new FileStream(zipPath, FileMode.Create))
        //    using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create))
        //    {
        //        foreach (var filePath in files)
        //        {
        //            string entryName = Path.GetFileName(filePath); // فقط اسم فایل داخل zip
        //            ZipArchiveEntry entry = archive.CreateEntry(entryName);

        //            using (var entryStream = entry.Open())
        //            using (var fileStream = System.IO.File.OpenRead(filePath))
        //            {
        //                fileStream.CopyTo(entryStream);
        //            }
        //        }
        //    }

        //    byte[] zipBytes = System.IO.File.ReadAllBytes(zipPath);
        //    return File(zipBytes, "application/zip", $"Fish_{year}_{month}.zip");
        //}
        [HttpPost]
        public ActionResult DeleteFishByMonth(int year, int month)
        {
            try
            {
                string basePath = Server.MapPath("~/Areas/Contracts/Contents/fish/");
                string monthFolder = Path.Combine(basePath, year + "-" + month.ToString("00"));

                if (!Directory.Exists(monthFolder))
                    return Content("پوشه این ماه وجود ندارد");

                // حذف کامل فولدر ماه (همه پیمانکارها و شهرها)
                Directory.Delete(monthFolder, true);

                return Content("فیش‌های این ماه با موفقیت حذف شدند");
            }
            catch (Exception ex)
            {
                return Content("خطا در حذف فایل‌ها: " + ex.Message);
            }
        }
        public ActionResult DownloadFishByMonth(int year, int month)
        {
            string basePath = Server.MapPath("~/Areas/Contracts/Contents/fish/");
            string monthFolder = Path.Combine(basePath, year + "-" + month.ToString("00"));

            if (!Directory.Exists(monthFolder))
                return Content("پوشه مربوط به این ماه وجود ندارد");

            // جستجو داخل تمام زیرپوشه‌ها
            var files = Directory.GetFiles(
                monthFolder,
                "*_Clean.pdf",
                SearchOption.AllDirectories
            );

            if (files.Length == 0)
                return Content("فایلی برای این ماه یافت نشد");

            string zipPath = Path.Combine(basePath, $"Fish_{year}_{month}.zip");

            if (System.IO.File.Exists(zipPath))
                System.IO.File.Delete(zipPath);

            using (FileStream fs = new FileStream(zipPath, FileMode.Create))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                foreach (var filePath in files)
                {
                    // مسیر داخل ZIP نسبت به فولدر ماه
                    string entryName = filePath.Substring(monthFolder.Length + 1) // مسیر نسبی
                                            .Replace("\\", "/"); // برای سازگاری با ZIP
                    ZipArchiveEntry entry = archive.CreateEntry(entryName);

                    using (var entryStream = entry.Open())
                    using (var fileStream = System.IO.File.OpenRead(filePath))
                    {
                        fileStream.CopyTo(entryStream);
                    }
                }
            }

            byte[] zipBytes = System.IO.File.ReadAllBytes(zipPath);
            return File(zipBytes, "application/zip", $"Fish_{year}_{month}.zip");
        }
        //public ActionResult DownloadFishByMonth(int year, int month)
        //{
        //    string basePath = Server.MapPath("~/Areas/Contracts/Contents/fish/");

        //    string monthFolder = Path.Combine(basePath, year + "-" + month.ToString("00"));

        //    if (!Directory.Exists(monthFolder))
        //        return Content("پوشه مربوط به این ماه وجود ندارد");

        //    // جستجو داخل تمام زیرپوشه‌ها
        //    var files = Directory.GetFiles(
        //        monthFolder,
        //        "*_Clean.pdf",
        //        SearchOption.AllDirectories
        //    );

        //    if (files.Length == 0)
        //        return Content("فایلی برای این ماه یافت نشد");

        //    string zipPath = Path.Combine(basePath, $"Fish_{year}_{month}.zip");

        //    if (System.IO.File.Exists(zipPath))
        //        System.IO.File.Delete(zipPath);

        //    using (FileStream fs = new FileStream(zipPath, FileMode.Create))
        //    using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create))
        //    {
        //        foreach (var filePath in files)
        //        {
        //            string entryName = Path.GetFileName(filePath); // فقط اسم فایل داخل zip
        //            ZipArchiveEntry entry = archive.CreateEntry(entryName);

        //            using (var entryStream = entry.Open())
        //            using (var fileStream = System.IO.File.OpenRead(filePath))
        //            {
        //                fileStream.CopyTo(entryStream);
        //            }
        //        }
        //    }

        //    byte[] zipBytes = System.IO.File.ReadAllBytes(zipPath);
        //    return File(zipBytes, "application/zip", $"Fish_{year}_{month}.zip");
        //}
        [HttpPost]
        public ActionResult UploadExcelFiles(HttpPostedFileBase OriginalExcel, HttpPostedFileBase NewExcel)
        {
            if (OriginalExcel == null || NewExcel == null)
                return Content("لطفاً هر دو فایل اکسل را آپلود کنید");

            // مسیر ذخیره روی سرور
            string saveFolder = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/");
            if (!Directory.Exists(saveFolder))
                Directory.CreateDirectory(saveFolder);

            // مسیر کامل برای هر فایل با حفظ نام اصلی
            string originalPath = Path.Combine(saveFolder, OriginalExcel.FileName);
            string newPath = Path.Combine(saveFolder, NewExcel.FileName);

            // ذخیره فایل‌ها
            OriginalExcel.SaveAs(originalPath);
            NewExcel.SaveAs(newPath);

            return Content($"فایل‌ها با موفقیت ذخیره شدند: {OriginalExcel.FileName} و {NewExcel.FileName}");
        }
        public ActionResult DownloadSampleFile(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return Content("نام فایل معتبر نیست");

            string folderPath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/");
            string filePath = Path.Combine(folderPath, fileName);

            if (!System.IO.File.Exists(filePath))
                return Content("فایل پیدا نشد");

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        public ActionResult excelfishtest1(ManualFishDetail Model, int UserID, int Month, int Year, int countpdf)
        {
            try
            {
                // داده‌های پایه
                var noalf = db.tbMoalefeDastmozdiValueFromExcel
                    .Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();
                var caran = db.tbCaranSettings.ToList();
                var final = db.FinancialDocuments.ToList();
                var moalf = db.tbContractMoalefeDastmozdi.ToList();
                var fish = db.tbMoalefeValueFish
                    .Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();

                if (Year >= 1404)
                    fish = fish.Where(p => p.FK_EXCel == null).ToList();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan1.xlsx"));
                var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
                //var final = db.FinancialDocuments.ToList();
                // فیلتر اسناد مالی برای ماه و سال
                PersianCalendar persianCalendar = new PersianCalendar();
                //var asnadstr = db.tbfkfinancial
                //    .Where(it => it.DataDocument.HasValue &&
                //                 persianCalendar.GetYear(it.DataDocument.Value) == Year &&
                //                 persianCalendar.GetMonth(it.DataDocument.Value) == Month)
                //    .ToList();
                List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();

                var asnad = db.tbfkfinancial.ToList();
                foreach (var it in asnad)
                {
                    DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                    int persianYear = persianCalendar.GetYear(dataDocument);
                    int persianMonth = persianCalendar.GetMonth(dataDocument);
                    if (persianYear == Year && persianMonth == Month)
                    {
                        asnadstr.Add(it);
                    }
                }
                long value2 = 0, value3 = 0, value4 = 0, value5 = 0, value6 = 0, value344 = 0;
                long ezafeh6001 = 0, kam6001 = 0, ezafeh7001 = 0, kam7001 = 0;

                foreach (var item2 in asnadstr)
                {
                    var t3 = final.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();

                    foreach (var item1 in t3)
                    {
                        if (!item1.DataDocument.HasValue) continue;

                        var d = item1.DataDocument.Value;
                        if (persianCalendar.GetYear(d) != Year || persianCalendar.GetMonth(d) != Month) continue;

                        int code = (int)item1.tbfkfinancial.numbershomar;

                        if (item1.Debtore != 0 && item1.Debtore != null)
                            value2 = (long)item1.Debtore;

                        if (item1.Creditor != 0 && item1.Creditor != null)
                        {
                            switch (code)
                            {
                                case 5001: value3 += (long)item1.Creditor; break;
                                case 5002: value4 += (long)item1.Creditor; break;
                                case 6001: ezafeh6001 += (long)item1.Creditor; break;
                                case 7001: ezafeh7001 += (long)item1.Creditor; break;
                                case 5003: value6 += (long)item1.Creditor; break;
                            }
                        }

                        if (item1.Debtore != 0 && item1.Debtore != null)
                        {
                            switch (code)
                            {
                                case 5001: value344 += (long)item1.Debtore; break;
                                case 6001: kam6001 += (long)item1.Debtore; break;
                                case 7001: kam7001 += (long)item1.Debtore; break;
                            }
                        }
                    }
                }

                using (var package = new ExcelPackage(originalFile))
                {
                    var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(s => s.Index == 1);
                    if (worksheet2 != null)
                    {
                        var stringList = new List<string>();
                        var stringList2 = new List<string>();
                        var stringList3 = new List<string>();
                        var stringList4 = new List<string>();
                        var stringList5 = new List<string>();
                        var stringList6 = new List<string>();

                        #region پر کردن لیست‌ها از مدل

                        stringList.Add(Model.FishHeader.PersonalCode.ToString());
                        stringList.Add(Model.FishHeader.OnvanShoql);
                        stringList.Add(Model.FishHeader.NationalCode);
                        stringList.Add(Model.FishHeader.UserName);
                        stringList.Add(Model.FishValues.First(p => p.Title == "مزد شغل (روزانه-ریال)").Value.ToString("N0"));
                        stringList.Add(Model.FishValues.First(p => p.Title == "مزد سنوات (روزانه-ریال)").Value.ToString("N0"));
                        stringList.Add(Model.FishValues.First(p => p.Title == "اضافه کار ( هرساعت-ریال)").Value.ToString("N0"));
                        stringList.Add(Model.FishValues.First(p => p.Title == "جمع مزد مبنا ( روزانه-ریال)").Value.ToString("N0"));
                        stringList.Add(Model.FishHeader.MonthName);
                        stringList.Add(Model.FishHeader.year.ToString());
                        stringList.Add(Model.FishHeader.TedadRoozMonth.ToString("N0"));
                        stringList.Add(Model.FishHeader.PeymanName);

                        stringList2.Add(Model.FishValues.First(p => p.Title == "تعداد روز کارکرد ").Value.ToString("N0"));
                        stringList2.Add(Model.FishValues.First(p => p.Title == "تعداد ساعات اضافه کار").Value.ToString("N0"));
                        stringList2.Add(Model.FishValues.First(p => p.Title == "تعداد روز ماموریت").Value.ToString("N0"));
                        stringList2.Add(Model.FishValues.First(p => p.Title == "تعداد روز استعلاجی").Value.ToString("N0"));
                        stringList2.Add(Model.FishValues.First(p => p.Title == "نوبتکاری (ریال)").Value.ToString("N0"));
                        stringList2.Add(Model.FishValues.First(p => p.Title == "تعداد روز غیبت").Value.ToString("N0"));

                        stringList3.Add(Model.FishValues.First(p => p.Title == "مزد گروه (شغل)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "مزد سنوات (ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه مسکن (ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه اولاد(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "اضافه کاری (ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "حق تاهل").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "ماموریت (ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value.ToString("N0"));

                        var ayab = Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)");
                        stringList3.Add(ayab != null ? ayab.Value.ToString("N0") : "0");

                        stringList3.Add(Model.FishValues.First(p => p.Title == "پاداش نوع 2").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value.ToString("N0"));
                        stringList3.Add(Model.FishValues.First(p => p.Title == "جمع ناخالص حقوق و مزایا (ریال)").Value.ToString("N0"));

                        stringList4.Add(Model.FishValues.First(p => p.Title == "جمع کل مشمول بیمه (ریال)").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "جمع کل مشمول مالیات (ریال)").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "مالیات سهم کارمند (ریال)").Value.ToString("N0"));

                        var takmili = Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)");
                        stringList4.Add(takmili != null ? takmili.Value.ToString("N0") : "0");

                        stringList4.Add(Model.FishValues.First(p => p.Title == "جریمه نوع 2").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "ذخیره کار مازاد (ریال)").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "جمع کل کسورات (ریال)").Value.ToString("N0"));
                        stringList4.Add(Model.FishValues.First(p => p.Title == "خالص قابل دریافت (ریال)").Value.ToString("N0"));
                        var ex333 = db.tbDetailFish.Where(p => p.FK_usr == UserID && p.month == Month && p.year == Year).FirstOrDefault();
                        string title = "";
                        if (ex333 != null)
                        {
                            title = ex333.sharh;
                        }
                        var tb = noalf.Where(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Value != 0).ToList();
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
                                    mode.Add(
                                        $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
                                        $"( {(u.CaranAyabOZahab ?? 0).ToString("N0")}/ 0 /{(it.MoalfeVal_Value ?? 0)} )----"
                                    );
                                }
                            }
                            else
                            {
                                mode.Add(
                                    $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
                                    $"( {(yuuu.CaranAyabOZahab ?? 0).ToString("N0")}/ {(yuuu.CaranStandard ?? 0)} /{(it.MoalfeVal_Value ?? 0)} )---"
                                );
                            }
                        }

                        #endregion
                        int lastRow = worksheet2.Dimension.End.Row;

                        // شمارش ردیف‌های دارای مقدار در ستون‌ها
                        int rowCount = 0, rowCount2 = 0, rowCount3 = 0, rowCount4 = 0, rowCount6=0, rowCount5 = 0;

                        for (int row = 1; row <= lastRow; row++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 1].Text)) rowCount++;
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 3].Text)) rowCount2++;
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 5].Text)) rowCount3++;
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 9].Text)) rowCount4++;
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 13].Text)) rowCount5++;
                            if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 13].Text)) rowCount6++;

                        }

                        // -------- ستون A (از ردیف 13) ----------
                        for (int i = 13; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 1].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            var ex = fish.FirstOrDefault(p =>
                                p.tbContractMoalefeDastmozdi != null &&
                                p.tbContractMoalefeDastmozdi.md_Title == name &&
                                p.mlfvlfsh_Month == Month &&
                                p.mlfvlfsh_Year == Year &&
                                p.FK_User == UserID);

                            stringList.Add(ex?.mlfvlfsh_Value != null
                                ? ((decimal)ex.mlfvlfsh_Value).ToString("N0")
                                : "0");
                        }

                        // -------- ستون C (از ردیف 6) ----------
                        for (int i = 6; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 3].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            var ex = fish.FirstOrDefault(p =>
                                p.tbContractMoalefeDastmozdi != null &&
                                p.tbContractMoalefeDastmozdi.md_Title == name &&
                                p.mlfvlfsh_Month == Month &&
                                p.mlfvlfsh_Year == Year &&
                                p.FK_User == UserID);

                            if (ex?.mlfvlfsh_Value != null)
                                stringList2.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
                        }

                        // -------- ستون E (از ردیف 16) ----------
                        for (int i = 16; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 5].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            var ex = fish.FirstOrDefault(p =>
                                p.tbContractMoalefeDastmozdi != null &&
                                p.tbContractMoalefeDastmozdi.md_Title == name &&
                                p.mlfvlfsh_Month == Month &&
                                p.mlfvlfsh_Year == Year &&
                                p.FK_User == UserID);

                            if (ex?.mlfvlfsh_Value != null)
                                stringList3.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
                        }

                        // -------- ستون I (از ردیف 10) ----------
                        for (int i = 10; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 9].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            var ex = fish.FirstOrDefault(p =>
                                p.tbContractMoalefeDastmozdi != null &&
                                p.tbContractMoalefeDastmozdi.md_Title == name &&
                                p.mlfvlfsh_Month == Month &&
                                p.mlfvlfsh_Year == Year &&
                                p.FK_User == UserID);

                            if (ex?.mlfvlfsh_Value != null)
                                stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
                        }
                        for (int i = 1; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 13].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            // Try to parse the string to an int
                            if (!int.TryParse(name, out int nameNumber))
                            {
                                // Invalid number, skip this row
                                continue;
                            }

                            // Find the matching item
                            var find = asnadstr.FirstOrDefault(s => s.numbershomar == nameNumber);

                            if (find != null)
                            {
                                var find2 = final.Where(p => p.User_ID == UserID && p.FK_final == find.ID).FirstOrDefault();
                                if (find2 != null && find2.Creditor!=0)
                                {
                                    if (find2?.Creditor != null)
                                    {
                                        stringList5.Add(((decimal)find2.Creditor).ToString("N0"));

                                    }
                                    else
                                    {
                                        stringList5.Add("0");

                                    }
                                }
                                else
                                {
                                    stringList5.Add("0");

                                }

                                // Do something with 'find'
                            }

                     

                            //var ex = fish.FirstOrDefault(p =>
                            //    p.tbContractMoalefeDastmozdi != null &&
                            //    p.tbContractMoalefeDastmozdi.md_Title == name &&
                            //    p.mlfvlfsh_Month == Month &&
                            //    p.mlfvlfsh_Year == Year &&
                            //    p.FK_User == UserID);

                            //if (ex?.mlfvlfsh_Value != null)
                            //    stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
                        }
                        for (int i = 1; i <= lastRow; i++)
                        {
                            var name = worksheet2.Cells[i, 15].Text?.Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            // Try to parse the string to an int
                            if (!int.TryParse(name, out int nameNumber))
                            {
                                // Invalid number, skip this row
                                continue;
                            }

                            // Find the matching item
                            var find = asnadstr.FirstOrDefault(s => s.numbershomar == nameNumber);

                            if (find != null)
                            {
                                var find2 = final.Where(p => p.User_ID == UserID && p.FK_final == find.ID).FirstOrDefault();
                                if (find2 != null && find2.Debtore != 0)
                                {
                                    if (find2?.Debtore != null)
                                    {
                                        stringList6.Add(((decimal)find2.Debtore).ToString("N0"));

                                    }
                                    else
                                    {
                                        stringList6.Add("0");


                                    }
                                }
                                else
                                {
                                    stringList6.Add("0");

                                }
                                // Do something with 'find'
                            }

      

                            //var ex = fish.FirstOrDefault(p =>
                            //    p.tbContractMoalefeDastmozdi != null &&
                            //    p.tbContractMoalefeDastmozdi.md_Title == name &&
                            //    p.mlfvlfsh_Month == Month &&
                            //    p.mlfvlfsh_Year == Year &&
                            //    p.FK_User == UserID);

                            //if (ex?.mlfvlfsh_Value != null)
                            //    stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
                        }
                        // نوشتن مقادیر در سلول‌ها
                        for (int i = 1; i <= stringList.Count; i++) worksheet2.Cells[$"B{i}"].Value = stringList[i - 1];
                        for (int i = 1; i <= stringList2.Count; i++) worksheet2.Cells[$"D{i}"].Value = stringList2[i - 1];
                        for (int i = 1; i <= stringList3.Count; i++) worksheet2.Cells[$"F{i}"].Value = stringList3[i - 1];
                        for (int i = 1; i <= stringList4.Count; i++) worksheet2.Cells[$"J{i}"].Value = stringList4[i - 1];
                        for (int i = 1; i <= stringList5.Count; i++) worksheet2.Cells[$"N{i}"].Value = stringList5[i - 1];
                        for (int i = 1; i <= stringList6.Count; i++) worksheet2.Cells[$"P{i}"].Value = stringList6[i - 1];

                        var cell2 = worksheet2.Cells[$"L{1}"];
                        cell2.Value = string.Join(" ", mode);
                        worksheet2.Cells["L4"].Value = value3;
                        var cell4 = worksheet2.Cells[$"L{5}"];
                        cell4.Value = value5;
                        worksheet2.Cells["L9"].Value = value344;
                        worksheet2.Cells["L10"].Value = ezafeh6001;
                        worksheet2.Cells["L11"].Value = kam6001;
                        worksheet2.Cells["L12"].Value = ezafeh7001;
                        worksheet2.Cells["L13"].Value = kam7001;
                        worksheet2.Cells["L6"].Value = value6;
                        worksheet2.Cells["L7"].Value = value4;
                        worksheet2.Cells["L8"].Value = value2;
                        var cell8 = worksheet2.Cells[$"L{2}"];
                        cell8.Value = title;

                        package.SaveAs(new FileInfo(newFilePath));
                    }
                    var worksheet3 = package.Workbook.Worksheets[0];
                    if (worksheet3 != null)
                    {
                        worksheet3.Calculate();
                        foreach (var cell in worksheet3.Cells[worksheet3.Dimension.Address])
                            if (!string.IsNullOrEmpty(cell.Formula))
                                cell.Value = cell.Value;
                    }

                    // ذخیره فایل اکسل
                    package.SaveAs(new FileInfo(newFilePath));
                    // محاسبه شیت اول
                    //var worksheet3 = package.Workbook.Worksheets[0];
                    //worksheet3.Calculate();
                    //package.SaveAs(new FileInfo(newFilePath));
                }
                string fullnam = "";
                var findusr = db.tbUsers.Where(s => s.usr_Personal_ID == Model.FishHeader.PersonalCode).FirstOrDefault();
                string cit = "";
                if (findusr != null)
                {
                    fullnam = findusr.FullName;
                    var findci = db.tbCities.Where(s => s.ID == findusr.usr_City_Dutysystem).FirstOrDefault();
                    if (findci != null)
                    {
                        cit = findci.Name;

                    }
                }

                //
                var findpymn = db.Link_User_And_Peyman.Where(s => s.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && s.Status == true).FirstOrDefault();
                string extractedFileName = "";
                if (findpymn != null)
                {
                    extractedFileName = findpymn.tbPeymanContracts.pec_Title;
                }
                string basePath = Server.MapPath("~/Areas/Contracts/Contents/fish/");

                string monthFolderName = Year + "-" + Month.ToString("00");
                string monthFolderPath = Path.Combine(basePath, monthFolderName);

                if (!Directory.Exists(monthFolderPath))
                    Directory.CreateDirectory(monthFolderPath);
                string safePeyman = string.Concat(extractedFileName
                    .Where(c => !Path.GetInvalidFileNameChars().Contains(c)));
                string peymanFolderPath = Path.Combine(monthFolderPath, safePeyman);

                if (!Directory.Exists(peymanFolderPath))
                    Directory.CreateDirectory(peymanFolderPath);
                string safeCity = string.Concat(cit
    .Where(c => !Path.GetInvalidFileNameChars().Contains(c)));

                string cityFolderPath = Path.Combine(peymanFolderPath, safeCity);

                if (!Directory.Exists(cityFolderPath))
                    Directory.CreateDirectory(cityFolderPath);

                string fileNameBase =
         Model.FishHeader.PersonalCode + "-" +
         safePeyman + "-" +
         safeCity + "-" +
         fullnam + "-" +
         Month + "-" +
         Year;

                string outputPdfPath = Path.Combine(cityFolderPath, fileNameBase + "_Clean.pdf");
                string pdfPath = Path.Combine(cityFolderPath, fileNameBase + ".pdf");


                //string outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + cit + '-' + fullnam + '-' + Month + '-' + Year + "_Clean.pdf");
                //string pdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + cit + '-' + fullnam + '-' + Month + '-' + Year + ".pdf");
                // تبدیل به PDF با GemBox
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                //var workbook1 = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);
                //while (workbook1.Worksheets.Count > 1) workbook1.Worksheets.Remove(1);

                //string outputPdfPath = Server.MapPath($"~/Areas/Contracts/Contents/fish/{UserID}-{Month}-{Year}.pdf");
                //string extractedFileName = peymanName2.Substring(startIndex2 + 1).Trim();
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                try
                {
                    // بارگذاری فایل اکسل با GemBox
                    GemBox.Spreadsheet.ExcelFile workbook1 = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);

                    // فقط شیت اول را نگه دار، بقیه را حذف کن (چون در Aspose پنهان می‌کردی)
                    while (workbook1.Worksheets.Count > 1)
                        workbook1.Worksheets.Remove(1);

                    // تبدیل به PDF
                    workbook1.Save(outputPdfPath);
                }
                catch (Exception ex)
                {
                    // ثبت خطا یا ادامه کار
                }

                //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
                //workbook.Dispose();

                try
                {
                    //workbook.Save(outputPdfPath);
                    //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
                    //workbook.Dispose();
                }
                catch (Exception ex)
                {
                }
                //RemoveWatermarkFromPDF(pdfPath, outputPdfPath);
                System.Diagnostics.Debug.WriteLine("واترمارک با موفقیت حذف شد!");



                byte[] fileBytes = System.IO.File.ReadAllBytes(outputPdfPath);
                string fileName = Path.GetFileName(outputPdfPath);

                return File(fileBytes, "application/pdf", fileName);

                //workbook1.Save(outputPdfPath);

                //byte[] fileBytes = System.IO.File.ReadAllBytes(outputPdfPath);
                //return File(fileBytes, "application/pdf", Path.GetFileName(outputPdfPath));
            }
            catch (Exception ex)
            {
                viewcontinue(countpdf);
                return Content(ex.Message);
            }
            //catch (Exception ex)
            //{
            //    string rx = ex.Message;
            //    return Content(ex.Message);

            //}
        }

        [AuthorizeAAA]
        //S(------------------------------------------------------------|Mk|
        //public ActionResult excelfishtest1(ManualFishDetail Model, int UserID, int Month, int Year, int countpdf)
        //{
        //    var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();
        //    var caran = db.tbCaranSettings.ToList();
        //    var final = db.FinancialDocuments.ToList();
        //    var moalf = db.tbContractMoalefeDastmozdi.ToList();
        //    var fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();
        //    if (Year >= 1404)
        //    {
        //        fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_EXCel == null).ToList();
        //    }
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan1.xlsx"));
        //    var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
        //    var asnad = db.tbfkfinancial.ToList();
        //    List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
        //    foreach (var it in asnad)
        //    {
        //        PersianCalendar persianCalendar = new PersianCalendar();
        //        DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //        int persianYear = persianCalendar.GetYear(dataDocument);
        //        int persianMonth = persianCalendar.GetMonth(dataDocument);
        //        if (persianYear == Year && persianMonth == Month)
        //        {
        //            asnadstr.Add(it);
        //        }
        //    }
        //    long value = 0;
        //    long value2 = 0;
        //    long value3 = 0;
        //    long value4 = 0;
        //    long value5 = 0;
        //    long value6 = 0;
        //    long value344 = 0;
        //    long ezafehvalue6001 = 0; long kamvalue6001 = 0;
        //    long ezafehvalue7001 = 0; long kamvalue7001 = 0;

        //    foreach (var item2 in asnadstr)
        //    {
        //        var t3 = final.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();
        //        foreach (var item1 in t3)
        //        {
        //            if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
        //            {
        //                PersianCalendar persianCalendar = new PersianCalendar();
        //                DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //                int persianYear = persianCalendar.GetYear(dataDocument);
        //                int persianMonth = persianCalendar.GetMonth(dataDocument);
        //                if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
        //                {
        //                    if (item1.Debtore != 0 && item1.Creditor != 0)
        //                    {
        //                        value += (long)(item1.Creditor - item1.Debtore);
        //                    }
        //                    else if (item1.Debtore != 0 && item1.Debtore != null)
        //                    {
        //                        value2 = (long)(item1.Debtore ?? 0);
        //                    }
        //                    else if (item1.Creditor != 0 && item1.Creditor != null)
        //                    {
        //                        value += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value3 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value344 += (long)(item1.Debtore ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5002)
        //                    {
        //                        value4 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value344 += (long)(item1.Debtore ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 6001)
        //                    {
        //                        ezafehvalue6001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 6001)
        //                    {
        //                        kamvalue6001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 7001)
        //                    {
        //                        ezafehvalue7001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Debtore != 0 && item1.tbfkfinancial.numbershomar == 7001)
        //                    {
        //                        kamvalue7001 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5003)
        //                    {
        //                        value6 += (long)(item1.Creditor ?? 0);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    try
        //    {
        //        using (var package = new ExcelPackage())
        //        {

        //            using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
        //            {
        //                package.Load(stream);
        //            }
        //            // Get the second sheet of the new file if it exists
        //            var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);
        //            // Check if the worksheet exists
        //            if (worksheet2 != null)
        //            {
        //                // Define a list of strings to be added to the cells in columns C1 to C42
        //                var stringList = new List<string>();
        //                var stringList2 = new List<string>();
        //                var stringList3 = new List<string>();
        //                var stringList4 = new List<string>();
        //                var stringList5 = new List<string>();
        //                #region fill string list
        //                stringList.Add(Model.FishHeader.PersonalCode.ToString(""));
        //                stringList.Add(Model.FishHeader.OnvanShoql);
        //                stringList.Add(Model.FishHeader.NationalCode);
        //                stringList.Add(Model.FishHeader.UserName.ToString());
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد شغل (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کار ( هرساعت-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع مزد مبنا ( روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishHeader.MonthName);
        //                stringList.Add(Model.FishHeader.year.ToString());
        //                stringList.Add(Model.FishHeader.TedadRoozMonth.ToString("N0"));
        //                var endDate = new DateTime(2025, 3, 20);
        //                var finddd = db.Link_User_And_Peyman.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.Status == true).FirstOrDefault();
        //                var findob = db.tbUserContracts.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.usc_EndTime == endDate && p.FK_KarfarmaID != null).FirstOrDefault();
        //                string peymanName2 = Model.FishHeader.PeymanName;
        //                if (findob != null)
        //                {
        //                    peymanName2 = findob.tbCompanies.CompanyName;
        //                }
        //                if (finddd != null)
        //                {
        //                    var finj = db.tbPeymanContracts.Where(p => p.FK_UserTarafDovvom != null && p.pec_ID == finddd.FK_Peyman_ID).FirstOrDefault();
        //                    if (finj != null)
        //                    {
        //                        peymanName2 = db.tbCompanies.Where(p => p.ID == finj.FK_UserTarafDovvom).Select(s => s.CompanyName).FirstOrDefault();
        //                    }
        //                }
        //                // جداکننده مورد نظر
        //                char separator2 = '-';
        //                // یافتن موقعیت جداکننده در رشته
        //                int startIndex2 = peymanName2.IndexOf(separator2);
        //                string extractedFileName = peymanName2.Substring(startIndex2 + 1).Trim();
        //                stringList.Add(extractedFileName.ToString());
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز کارکرد ").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد ساعات اضافه کار").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز ماموریت").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز استعلاجی").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز غیبت").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد گروه (شغل)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق تاهل").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ماموریت (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
        //                {
        //                    stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)").Value.ToString("N0"));
        //                }
        //                else
        //                {
        //                    stringList3.Add("0");
        //                }
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "پاداش نوع 2").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع ناخالص حقوق و مزایا (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول بیمه (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول مالیات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)") != null)
        //                {
        //                    stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)").Value.ToString("N0"));
        //                }
        //                else
        //                {
        //                    stringList4.Add("0");
        //                }
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جریمه نوع 2").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کار مازاد (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل کسورات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "خالص قابل دریافت (ریال)").Value.ToString("N0"));
        //                var ex333 = db.tbJaremehandpadash.Where(p => p.Fk_usr == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
        //                string title = "";
        //                if (ex333 != null)
        //                {
        //                    title = ex333.Titel;
        //                }
        //                var tb = noalf.Where(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Value != 0).ToList();
        //                List<string> mode = new List<string>();
        //                foreach (var it in tb)
        //                {
        //                    var yuuu = caran
        //                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
        //                        .FirstOrDefault();

        //                    if (yuuu == null)
        //                    {
        //                        var u = caran
        //                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
        //                            .FirstOrDefault();

        //                        if (u != null)
        //                        {
        //                            mode.Add(
        //                                $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
        //                                $"( {(u.CaranAyabOZahab ?? 0).ToString("N0")}/ 0 /{(it.MoalfeVal_Value ?? 0)} )----"
        //                            );
        //                        }
        //                    }
        //                    else
        //                    {
        //                        mode.Add(
        //                            $"{it.tbContractMoalefeDastmozdi?.md_Title} " +
        //                            $"( {(yuuu.CaranAyabOZahab ?? 0).ToString("N0")}/ {(yuuu.CaranStandard ?? 0)} /{(it.MoalfeVal_Value ?? 0)} )---"
        //                        );
        //                    }
        //                }

        //                //foreach (var it in tb)
        //                //{
        //                //    var yuuu = caran.Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0).FirstOrDefault();
        //                //    if (yuuu == null)
        //                //    {
        //                //        var u = caran.Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0).FirstOrDefault();
        //                //        if (u != null)
        //                //        {
        //                //            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)u.CaranAyabOZahab).ToString("N0")}/ 0 /{it.MoalfeVal_Value} )----");
        //                //        }
        //                //    }
        //                //    else
        //                //    {
        //                //        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)yuuu.CaranAyabOZahab).ToString("N0")}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
        //                //    }
        //                //}
        //                if (db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.FKUser == UserID && p.Year == Year && p.RemainValue != 0 && p.RemainValue != null).FirstOrDefault() != null)
        //                {
        //                    mode.Add($"(مازاد اضافه کار در حال بررسی میباشد  )----");
        //                }
        //                int totalRows = worksheet2.Dimension.End.Row;

        //                int rowCountCol1 = 0;
        //                int rowCountCol3 = 0;
        //                int rowCountCol5 = 0;
        //                int rowCountCol9 = 0;

        //                for (int row = 1; row <= totalRows; row++)
        //                {
        //                    if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 1].Text))
        //                        rowCountCol1++;

        //                    if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 3].Text))
        //                        rowCountCol3++;

        //                    if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 5].Text))
        //                        rowCountCol5++;

        //                    if (!string.IsNullOrWhiteSpace(worksheet2.Cells[row, 9].Text))
        //                        rowCountCol9++;
        //                }
        //                var List = worksheet2.Rows;
        //                var count = worksheet2.Rows[0].Count();
        //                var count2 = worksheet2.Rows.Count();
        //                int rowCount = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 1].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 1].Value.ToString()))
        //                    {
        //                        rowCount++;
        //                    }
        //                }
        //                int rowCount2 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 3].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 3].Value.ToString()))
        //                    {
        //                        rowCount2++;
        //                    }
        //                }

        //                int rowCount3 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 5].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 5].Value.ToString()))
        //                    {
        //                        rowCount3++;
        //                    }
        //                }
        //                int rowCount4 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 9].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 9].Value.ToString()))
        //                    {
        //                        rowCount4++;
        //                    }
        //                }
        //                var countColumnARows = rowCount;
        //                for (int i = 13; i <= countColumnARows; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 1].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                        else
        //                        {
        //                            stringList.Add("0");
        //                        }
        //                    }
        //                }
        //                #endregion

        //                for (int i = 6; i <= rowCount2; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 3].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList2.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                for (int i = 16; i <= rowCount3; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 5].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList3.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                for (int i = 10; i <= rowCount4; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 9].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }
        //                }
        //                // Loop through each row in columns C1 to C42 and add the corresponding string from the list
        //                for (int i = 1; i <= stringList.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"B{i}"];
        //                    cell.Value = stringList[i - 1];
        //                }
        //                for (int i = 1; i <= stringList2.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"D{i}"];
        //                    cell.Value = stringList2[i - 1];
        //                }
        //                for (int i = 1; i <= stringList3.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"F{i}"];
        //                    cell.Value = stringList3[i - 1];
        //                }
        //                for (int i = 1; i <= stringList4.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"J{i}"];
        //                    cell.Value = stringList4[i - 1];
        //                }
        //                var cell2 = worksheet2.Cells[$"L{1}"];
        //                cell2.Value = string.Join(" ", mode);
        //                var cell3 = worksheet2.Cells[$"L{4}"];
        //                cell3.Value = value3;
        //                var cell32 = worksheet2.Cells[$"L{9}"];
        //                cell32.Value = value344;

        //                var cell10 = worksheet2.Cells[$"L{10}"];
        //                cell10.Value = ezafehvalue6001;
        //                var cell11 = worksheet2.Cells[$"L{11}"];
        //                cell11.Value = kamvalue6001;

        //                var cell12 = worksheet2.Cells[$"L{12}"];
        //                cell12.Value = ezafehvalue7001;
        //                var cell13 = worksheet2.Cells[$"L{13}"];
        //                cell13.Value = kamvalue7001;


        //                var cell4 = worksheet2.Cells[$"L{5}"];
        //                cell4.Value = value5;
        //                var cell5 = worksheet2.Cells[$"L{6}"];
        //                cell5.Value = value6;
        //                var cell6 = worksheet2.Cells[$"L{7}"];
        //                cell6.Value = value4;
        //                var cell7 = worksheet2.Cells[$"L{8}"];
        //                cell7.Value = value2;
        //                var cell8 = worksheet2.Cells[$"L{2}"];
        //                cell8.Value = title;
        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //                // استفاده از فضای خالی به جای Environment.NewLine
        //            }
        //            var worksheet3 = package.Workbook.Worksheets[0];
        //            if (worksheet3 != null)
        //            {
        //                worksheet3.Calculate();
        //                var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address].Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();
        //                if (cellsWithFormulas.Any())
        //                {
        //                    foreach (var cell in cellsWithFormulas)
        //                    {
        //                        var value1232 = cell.Value;
        //                        cell.Value = value1232;
        //                    }
        //                }
        //                else
        //                {
        //                    System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
        //                }
        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //            }
        //            // Save changes to the new file
        //        }
        //        using (ExcelEngine excelEngine = new ExcelEngine())
        //        {

        //            string peymanName = Model.FishHeader.PeymanName;
        //            string extractedFileName = peymanName;
        //            //Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(newFilePath);
        //            //for (int i = 1; i < workbook.Worksheets.Count; i++)
        //            //{
        //            //    workbook.Worksheets[i].IsVisible = false;
        //            //}
        //            string fullnam = "";
        //            var findusr = db.tbUsers.Where(s => s.usr_Personal_ID == Model.FishHeader.PersonalCode).FirstOrDefault();
        //            if (findusr != null)
        //            {
        //                fullnam = findusr.FullName;
        //            }
        //            string outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + fullnam + '-' + Month + '-' + Year + "_Clean.pdf");
        //            string pdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + fullnam + '-' + Month + '-' + Year + ".pdf");
        //            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        //            try
        //            {
        //                // بارگذاری فایل اکسل با GemBox
        //                GemBox.Spreadsheet.ExcelFile workbook1 = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);

        //                // فقط شیت اول را نگه دار، بقیه را حذف کن (چون در Aspose پنهان می‌کردی)
        //                while (workbook1.Worksheets.Count > 1)
        //                    workbook1.Worksheets.Remove(1);

        //                // تبدیل به PDF
        //                workbook1.Save(outputPdfPath);
        //            }
        //            catch (Exception ex)
        //            {
        //                // ثبت خطا یا ادامه کار
        //            }

        //            //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
        //            //workbook.Dispose();

        //            try
        //            {
        //                //workbook.Save(outputPdfPath);
        //                //workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);
        //                //workbook.Dispose();
        //            }
        //            catch (Exception ex)
        //            {
        //            }
        //            //RemoveWatermarkFromPDF(pdfPath, outputPdfPath);
        //            System.Diagnostics.Debug.WriteLine("واترمارک با موفقیت حذف شد!");




        //            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPdfPath);
        //            string fileName = Path.GetFileName(outputPdfPath);

        //            return File(fileBytes, "application/pdf", fileName);
        //            //return 0;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        viewcontinue(countpdf);
        //        return Content(ex.Message);
        //    }
        //}
        //E)------------------------------------------------------------|Mk|









        //S(------------------------------------------------------------|Mk|
        public ActionResult TestGemBox()
        {
            var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
            var outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/output.pdf");
            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var file = GemBox.Spreadsheet.ExcelFile.Load(newFilePath);

            while (file.Worksheets.Count > 1)
                file.Worksheets.Remove(1);

            file.Save(outputPdfPath);

            return Content("تبدیل با موفقیت انجام شد.");
        }
        //E)------------------------------------------------------------|Mk|
        //public void excelfish(ManualFishDetail Model, int UserID, int Month, int Year)
        //{
        //    var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();
        //    var caran = db.tbCaranSettings.ToList();
        //    var final = db.FinancialDocuments.ToList();
        //    var moalf = db.tbContractMoalefeDastmozdi.ToList();
        //    var fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year).ToList();
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //    var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan1.xlsx"));
        //    var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
        //    var asnad = db.tbfkfinancial.ToList();
        //    List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
        //    foreach (var it in asnad)
        //    {
        //        PersianCalendar persianCalendar = new PersianCalendar();
        //        DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //        int persianYear = persianCalendar.GetYear(dataDocument);
        //        int persianMonth = persianCalendar.GetMonth(dataDocument);
        //        if (persianYear == Year && persianMonth == Month)
        //        {
        //            asnadstr.Add(it);

        //        }
        //    }



        //    long value = 0;
        //    long value2 = 0;
        //    long value3 = 0;
        //    long value4 = 0;
        //    long value5 = 0;
        //    long value6 = 0;

        //    foreach (var item2 in asnadstr)
        //    {


        //        var t3 = final.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();

        //        foreach (var item1 in t3)
        //        {
        //            if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
        //            {
        //                PersianCalendar persianCalendar = new PersianCalendar();
        //                DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //                int persianYear = persianCalendar.GetYear(dataDocument);
        //                int persianMonth = persianCalendar.GetMonth(dataDocument);
        //                if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
        //                {
        //                    if (item1.Debtore != 0 && item1.Creditor != 0)
        //                    {
        //                        value += (long)(item1.Creditor - item1.Debtore);
        //                    }


        //                    else if (item1.Debtore != 0 && item1.Debtore != null)
        //                    {
        //                        value2 = (long)(item1.Debtore ?? 0);
        //                    }
        //                    else if (item1.Creditor != 0 && item1.Creditor != null)
        //                    {
        //                        value += (long)(item1.Creditor ?? 0);

        //                    }

        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5001)
        //                    {
        //                        value3 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5002)
        //                    {
        //                        value4 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5004)
        //                    {
        //                        value5 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5003)
        //                    {
        //                        value6 += (long)(item1.Creditor ?? 0);
        //                    }



        //                }

        //            }
        //        }
        //    }




        //    //var originalFile = new FileInfo(Server.MapPath("~/Areas/Content/ExcelFiles/excelfish.xlsx"));
        //    //var newFilePath = Server.MapPath("~/Areas/Content/ExcelFiles/excelfish.xlsx");
        //    try
        //    {
        //        using (var package = new ExcelPackage())
        //        {
        //            // Copy data and formatting from the original file to the new file
        //            using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
        //            {
        //                package.Load(stream);
        //            }

        //            // Get the second sheet of the new file if it exists
        //            var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);

        //            // Check if the worksheet exists
        //            if (worksheet2 != null)
        //            {
        //                // Define a list of strings to be added to the cells in columns C1 to C42
        //                var stringList = new List<string>();
        //                var stringList2 = new List<string>();

        //                var stringList3 = new List<string>();

        //                var stringList4 = new List<string>();
        //                var stringList5 = new List<string>();


        //                #region fill string list
        //                stringList.Add(Model.FishHeader.PersonalCode.ToString(""));
        //                stringList.Add(Model.FishHeader.OnvanShoql);
        //                stringList.Add(Model.FishHeader.NationalCode);
        //                stringList.Add(Model.FishHeader.UserName.ToString());
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد شغل (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کار ( هرساعت-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع مزد مبنا ( روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishHeader.MonthName);
        //                stringList.Add(Model.FishHeader.year.ToString());
        //                stringList.Add(Model.FishHeader.TedadRoozMonth.ToString("N0"));
        //                var endDate = new DateTime(2025, 3, 20);
        //                var finddd = db.Link_User_And_Peyman.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.Status == true).FirstOrDefault();


        //                var findob = db.tbUserContracts
        //                    .Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode
        //                                && p.usc_EndTime == endDate
        //                                && p.FK_KarfarmaID != null)
        //                    .FirstOrDefault();
        //                string peymanName2 = Model.FishHeader.PeymanName;

        //                if (findob != null)
        //                {
        //                    peymanName2 = findob.tbCompanies.CompanyName;
        //                }
        //                if (finddd != null)
        //                {
        //                    var finj = db.tbPeymanContracts.Where(p => p.FK_UserTarafDovvom != null && p.pec_ID == finddd.FK_Peyman_ID).FirstOrDefault();
        //                    if (finj != null)
        //                    {
        //                        peymanName2 = db.tbCompanies.Where(p => p.ID == finj.FK_UserTarafDovvom).Select(s => s.CompanyName).FirstOrDefault();

        //                    }

        //                }

        //                // جداکننده مورد نظر
        //                char separator2 = '-';

        //                // یافتن موقعیت جداکننده در رشته
        //                int startIndex2 = peymanName2.IndexOf(separator2);
        //                string extractedFileName = peymanName2.Substring(startIndex2 + 1).Trim();
        //                stringList.Add(extractedFileName.ToString());







        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز کارکرد ").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد ساعات اضافه کار").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز ماموریت").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز استعلاجی").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز غیبت").Value.ToString("N0"));





        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد گروه (شغل)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق تاهل").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ماموریت (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "پاداش نوع 2").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع ناخالص حقوق و مزایا (ریال)").Value.ToString("N0"));
        //                //stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));




        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول بیمه (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول مالیات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)") != null)
        //                {
        //                    stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)").Value.ToString("N0"));

        //                }
        //                else
        //                {
        //                    stringList4.Add("0");

        //                }
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جریمه نوع 2").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کار مازاد (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل کسورات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "خالص قابل دریافت (ریال)").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));

        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));


        //                var ex333 = db.tbJaremehandpadash.Where(p => p.Fk_usr == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
        //                string title = "";
        //                if (ex333 != null)
        //                {
        //                    title = ex333.Titel;
        //                }


        //                var tb = noalf
        //         .Where(p => p.MoalfeVal_FKUser == UserID

        //                     && p.MoalfeVal_Value != 0
        //                     )
        //         .ToList();
        //                List<string> mode = new List<string>();

        //                foreach (var it in tb)
        //                {
        //                    var yuuu = caran
        //                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
        //                        .FirstOrDefault();

        //                    if (yuuu == null)
        //                    {
        //                        var u = caran
        //                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
        //                            .FirstOrDefault();

        //                        if (u != null)
        //                        {
        //                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)u.CaranAyabOZahab).ToString("N0")}/ 0 /{it.MoalfeVal_Value} )----");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)yuuu.CaranAyabOZahab).ToString("N0")}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
        //                    }
        //                }
        //                if (db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.FKUser == UserID && p.Year == Year && p.RemainValue != 0 && p.RemainValue != null).FirstOrDefault() != null)
        //                {
        //                    mode.Add($"(مازاد اضافه کار در حال بررسی میباشد  )----");

        //                }




        //                //var title = worksheet2.Rows[0].Cells;
        //                var List = worksheet2.Rows;
        //                var count = worksheet2.Rows[0].Count();
        //                var count2 = worksheet2.Rows.Count();

        //                //for (int i = 1; i < count; i++)
        //                //{
        //                //    var row = worksheet2.Rows[i];
        //                //    var Name = row.Cells[0];
        //                //} 
        //                int rowCount = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 1].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 1].Value.ToString()))
        //                    {
        //                        rowCount++;
        //                    }
        //                }

        //                int rowCount2 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 3].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 3].Value.ToString()))
        //                    {
        //                        rowCount2++;
        //                    }
        //                }


        //                int rowCount3 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 5].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 5].Value.ToString()))
        //                    {
        //                        rowCount3++;
        //                    }
        //                }

        //                int rowCount4 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 9].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 9].Value.ToString()))
        //                    {
        //                        rowCount4++;
        //                    }
        //                }




        //                var countColumnARows = rowCount;
        //                for (int i = 13; i <= countColumnARows; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 1].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                        else
        //                        {
        //                            stringList.Add("0");
        //                        }


        //                    }

        //                }

        //                #endregion

        //                for (int i = 6; i <= rowCount2; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 3].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList2.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }
        //                for (int i = 16; i <= rowCount3; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 5].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList3.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }
        //                for (int i = 10; i <= rowCount4; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 9].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }

        //                // Loop through each row in columns C1 to C42 and add the corresponding string from the list
        //                for (int i = 1; i <= stringList.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"B{i}"];
        //                    cell.Value = stringList[i - 1];
        //                }
        //                for (int i = 1; i <= stringList2.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"D{i}"];
        //                    cell.Value = stringList2[i - 1];
        //                }
        //                for (int i = 1; i <= stringList3.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"F{i}"];
        //                    cell.Value = stringList3[i - 1];
        //                }
        //                for (int i = 1; i <= stringList4.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"J{i}"];
        //                    cell.Value = stringList4[i - 1];
        //                }
        //                var cell2 = worksheet2.Cells[$"L{1}"];
        //                cell2.Value = string.Join(" ", mode);

        //                var cell3 = worksheet2.Cells[$"L{4}"];
        //                cell3.Value = value3;
        //                var cell4 = worksheet2.Cells[$"L{5}"];
        //                cell4.Value = value5;
        //                var cell5 = worksheet2.Cells[$"L{6}"];
        //                cell5.Value = value6;
        //                var cell6 = worksheet2.Cells[$"L{7}"];
        //                cell6.Value = value4;
        //                var cell7 = worksheet2.Cells[$"L{8}"];
        //                cell7.Value = value2;



        //                var cell8 = worksheet2.Cells[$"L{2}"];
        //                cell8.Value = title;


        //                // استفاده از فضای خالی به جای Environment.NewLine
        //            }

        //            // Save changes to the new file
        //            var newFile = new FileInfo(newFilePath);
        //            package.SaveAs(newFile);
        //        }
        //        using (ExcelEngine excelEngine = new ExcelEngine())
        //        {
        //            //IApplication application = excelEngine.Excel;
        //            //application.DefaultVersion = ExcelVersion.Xlsx;
        //            //FileStream excelStream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx"), FileMode.Open, FileAccess.Read);
        //            //IWorkbook workbook = application.Workbooks.Open(excelStream);
        //            //IWorksheet worksheet = workbook.Worksheets[0];

        //            //Initialize XlsIO renderer.
        //            //   XlsIORenderer renderer = new XlsIORenderer();

        //            //Convert Excel document into PDF document 
        //            //   PdfDocument pdfDocument = renderer.ConvertToPDF(worksheet);

        //            //    Stream stream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), FileMode.Create, FileAccess.ReadWrite);
        //            //    pdfDocument.Save(stream);
        //            //  pdfDocument.Close();
        //            string peymanName = Model.FishHeader.PeymanName;

        //            // جداکننده مورد نظر
        //            char separator = '-';

        //            // یافتن موقعیت جداکننده در رشته
        //            int startIndex = peymanName.IndexOf(separator);
        //            string extractedFileName = peymanName.Substring(startIndex + 1).Trim();



        //            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        //            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx"));
        //            string outputPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + ".pdf");
        //            //outputPath = outputPath.Replace("\\", "/");

        //            wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath);

        //            // Save the workbook as PDF
        //            //string pdfPath = Server.MapPath($"~/Areas/Contracts/Contents/fish/{Model.FishHeader.PersonalCode}_{Model.FishHeader.PeymanName}.pdf");
        //            //wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, pdfPath);
        //            //wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode  + ".pdf"));





        //            //  pdfDocument.Dispose();
        //            //   stream.Dispose();
        //            wb.Close(false);

        //            // Quit Excel application
        //            excelApp.Quit();


        //            // byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"));
        //            // return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "frm77.pdf");

        //            // return File(stream, "Forum77.pdf");

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //wb.Close(false);
        //        // Handle any exceptions that occur during the file creation process

        //    }

        //}

        public void FillExcel(tbUserContracts Model)
        {

            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.xlsx"));
            var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx");





            var obj = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
          var birthplace=  obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfBrith).Select(p => p.cityName).FirstOrDefault();
            var PlaceOfRegister = obj.Where(p => p.cityId == Model.tbUsers.usr_PlaceOfIssue).Select(p => p.cityName).FirstOrDefault();
            var typeofcontract = "";
            string maragestatus = "";
            string DutySystem = "";
            string MadrakTahsili = "";

            if (Model.usc_TypeOfContract == false)
            {
                typeofcontract = "وقت مزدی";
            }
            else
            {
                typeofcontract = "کار مزدی";
            }


            Row Row;
            if (Model.tbUsers.usr_MaritaIStatus == 1)
            {
                maragestatus = "متاهل";
            }
            else if (Model.tbUsers.usr_MaritaIStatus == 2)
            {
                maragestatus = "مجرد";
            }

            switch (Model.tbUsers.usr_Dutysystem)
            {
                case 1:
                    DutySystem = "مشمول";
                    break;
                case 2:
                    DutySystem = "پایان خدمت";
                    break;
                case 3:
                    DutySystem = "معافیت دائم";
                    break;
                case 4:
                    DutySystem = "معافیت موقت";
                    break;
                case 5:
                    DutySystem = "غایب";
                    break;
                case 6:
                    DutySystem = "خانم";
                    break;
                default:
                    DutySystem = "نامعلوم";
                    break;

            }

            ////////////////////
            if (Model.tbUsers.usr_Degree != null)
            {
                switch (Model.tbUsers.usr_Degree)
                {

                    case 1:
                        MadrakTahsili = "بیسواد";
                        break;
                    case 2:
                        MadrakTahsili = "سیکل";
                        break;
                    case 3:
                        MadrakTahsili = "دیپلم";
                        break;
                    case 4:
                        MadrakTahsili = "کاردانی";
                        break;
                    case 5:
                        MadrakTahsili = "کارشناسی";
                        break;
                    case 6:
                        MadrakTahsili = "کارشناسی ارشد";
                        break;

                    case 8:
                        MadrakTahsili = "دکتری";
                        break;
                    default:
                        MadrakTahsili = "نامعلوم ";

                        break;
                }
            }
            else
            {
                MadrakTahsili = "نامعلوم ";
            }


            try
            {
                using (var package = new ExcelPackage())
                {
                    // Copy data and formatting from the original file to the new file
                    using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
                    {
                        package.Load(stream);
                    }

                    // Get the second sheet of the new file if it exists
                    var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);

                    // Check if the worksheet exists
                    if (worksheet2 != null)
                    {
                        // Define a list of strings to be added to the cells in columns C1 to C42
                        var stringList = new List<string>();
                        #region fill string list
                        stringList.Add(Model.usc_ContractNumber);
                        stringList.Add(Model.tbUsers.usr_Family);
                        stringList.Add(Model.tbUsers.usr_Name);
                        stringList.Add(Model.tbUsers.usr_SHCode.ToString());
                        stringList.Add(Model.tbUsers.usr_FatherName);
                        stringList.Add(Model.tbUsers.shamsiDateOfBirth);
                        stringList.Add(birthplace);
                        stringList.Add(PlaceOfRegister);
                        stringList.Add(maragestatus);
                        stringList.Add(Model.tbUsers.usr_Child_Allowance.ToString());
                        stringList.Add(DutySystem);
                        stringList.Add(MadrakTahsili);
                        stringList.Add("---");
                        stringList.Add(Model.tbUsers.usr_NationalCode);
                        stringList.Add(Model.tbUsers.usr_Personal_ID.ToString());
                        stringList.Add(Model.tbCompanies.CompanyName);
                        stringList.Add(Model.tbCompanies.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies.Company_Address);
                        stringList.Add(Model.tbCompanies.RegistrationNumber.ToString());
                        stringList.Add(Model.tbCompanies.PlaceOfRegister.ToString());
                        stringList.Add(Model.usc_Jobtitle);
                        stringList.Add(Model.usc_JobCode.ToString());
                        stringList.Add(Model.FK_JobGroup.ToString());
                        stringList.Add(Model.tbCompanies1.CompanyName);
                        stringList.Add(Model.tbCompanies1.tbUsers.FullName);
                        stringList.Add(Model.tbCompanies1.Company_Address);
                        stringList.Add(Model.usc_ShmasiStartTime);
                        stringList.Add(Model.usc_ShmasiENDTime);
                        stringList.Add((Model.usc_EndTime - Model.usc_StartTime).ToString());
                        stringList.Add(typeofcontract);
                        stringList.Add(Model.jobgroup_ValueMozdGroup);
                        stringList.Add(Model.jobgroup_ValueSanavat);
                        stringList.Add("0");
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) + System.Convert.ToInt32(Model.jobgroup_ValueSanavat)).ToString());
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30).ToString());
                        stringList.Add((System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30).ToString());
                        stringList.Add("0");
                        stringList.Add(((System.Convert.ToInt32(Model.jobgroup_ValueMozdGroup) * 30) + (System.Convert.ToInt32(Model.jobgroup_ValueSanavat) * 30)).ToString());
                        stringList.Add(Model.jobgroup_HagheMaskan);
                        stringList.Add(Model.jobgroup_KharoBar);
                        stringList.Add(Model.jobgroup_HagheOlad);
                        var gharar = Model.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == Model.usc_ID).FirstOrDefault();
                        if (gharar != null)
                        {
                            stringList.Add(gharar.Value.ToString());
                            stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + gharar.Value + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());

                        }
                        else
                        {
                            stringList.Add("0");

                            stringList.Add((System.Convert.ToInt32(Model.jobgroup_HagheOlad) + System.Convert.ToInt32(Model.jobgroup_KharoBar) + System.Convert.ToInt32(Model.jobgroup_HagheMaskan)).ToString());

                        }


                        #endregion


                  

                            // Loop through each row in columns C1 to C42 and add the corresponding string from the list
                            for (int i = 1; i <= 43; i++)
                        {
                            var cell = worksheet2.Cells[$"C{i}"];
                            cell.Value = stringList[i - 1];
                        }

                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }
                    var worksheet3 = package.Workbook.Worksheets[0];

                    if (worksheet3 != null)
                    {

                        worksheet3.Calculate();


                        var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address]
                            .Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();

                        if (cellsWithFormulas.Any())
                        {
                            foreach (var cell in cellsWithFormulas)
                            {
                                var value = cell.Value;
                                cell.Value = value;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
                        }

                        var newFile = new FileInfo(newFilePath);
                        package.SaveAs(newFile);
                    }
                    // Save changes to the new file




                }
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //IApplication application = excelEngine.Excel;
                    //application.DefaultVersion = ExcelVersion.Xlsx;
                    //FileStream excelStream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx"), FileMode.Open, FileAccess.Read);
                    //IWorkbook workbook = application.Workbooks.Open(excelStream);
                    //IWorksheet worksheet = workbook.Worksheets[0];
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(newFilePath);

                    for (int i = 1; i < workbook.Worksheets.Count; i++)
                    {
                        workbook.Worksheets[i].IsVisible = false;
                    }


                    workbook.Save(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), Aspose.Cells.SaveFormat.Pdf);

                    System.Diagnostics.Debug.WriteLine("تبدیل به PDF با موفقیت انجام شد.");
                    //Initialize XlsIO renderer.
                    //   XlsIORenderer renderer = new XlsIORenderer();

                    //Convert Excel document into PDF document 
                    //   PdfDocument pdfDocument = renderer.ConvertToPDF(worksheet);

                    //    Stream stream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), FileMode.Create, FileAccess.ReadWrite);
                    //    pdfDocument.Save(stream);
                    //  pdfDocument.Close();
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx"));

                    // Save the workbook as PDF
                   
                    wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"));

                    //  pdfDocument.Dispose();
                    //   stream.Dispose();
                    wb.Close(false);

                    // Quit Excel application
                    excelApp.Quit();


                    // byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"));
                    // return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "frm77.pdf");

                    // return File(stream, "Forum77.pdf");

                }
            } 
            catch (Exception ex) 
            {
                // Handle any exceptions that occur during the file creation process
                
            }

          
            
        }






      


        //public int excelfishtest(ManualFishDetail Model, int UserID, int Month, int Year,int countpdf)
        //{
         
        //    var noalf = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).ToList();
        //    var caran = db.tbCaranSettings.ToList();
        //    var final = db.FinancialDocuments.ToList();
        //    var moalf = db.tbContractMoalefeDastmozdi.ToList();
        //    var fish = db.tbMoalefeValueFish.Where(p=>p.mlfvlfsh_Month==Month&&p.mlfvlfsh_Year==Year).ToList();
        //    if (Year >= 1404)
        //    {
        //        fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year&&p.FK_EXCel==null).ToList();
        //    }
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //    var originalFile = new FileInfo(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan1.xlsx"));
        //    var newFilePath = Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx");
        //    var asnad = db.tbfkfinancial.ToList();
        //    List<tbfkfinancial> asnadstr = new List<tbfkfinancial>();
        //    foreach (var it in asnad)
        //    {
        //        PersianCalendar persianCalendar = new PersianCalendar();
        //        DateTime dataDocument = it.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //        int persianYear = persianCalendar.GetYear(dataDocument);
        //        int persianMonth = persianCalendar.GetMonth(dataDocument);
        //        if (persianYear == Year && persianMonth == Month)
        //        {
        //            asnadstr.Add(it);

        //        }
        //    }



        //    long value = 0;
        //    long value2 = 0;
        //    long value3 = 0;
        //    long value4= 0;
        //    long value5 = 0;
        //    long value6 = 0;

        //    foreach (var item2 in asnadstr)
        //    {
           

        //        var t3 = final.Where(p => p.User_ID == UserID && p.FK_final == item2.ID).ToList();

        //        foreach (var item1 in t3)
        //        {
        //            if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
        //            {
        //                PersianCalendar persianCalendar = new PersianCalendar();
        //                DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
        //                int persianYear = persianCalendar.GetYear(dataDocument);
        //                int persianMonth = persianCalendar.GetMonth(dataDocument);
        //                if (persianYear == Year && persianMonth == Month && item1.FK_final == item2.ID)
        //                {
        //                    if (item1.Debtore != 0 && item1.Creditor != 0)
        //                    {
        //                        value += (long)(item1.Creditor - item1.Debtore);
        //                    }


        //                    else if (item1.Debtore != 0 && item1.Debtore != null)
        //                    {
        //                        value2 = (long)(item1.Debtore ?? 0);
        //                    }
        //                    else if (item1.Creditor != 0 && item1.Creditor != null)
        //                    {
        //                        value += (long)(item1.Creditor ?? 0);

        //                    }

        //                    if(item1.Creditor != 0&& item1.tbfkfinancial.numbershomar== 5001)
        //                    {
        //                        value3+= (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 &&  item1.tbfkfinancial.numbershomar == 5002)
        //                    {
        //                        value4+= (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5004)
        //                    {
        //                        value5 += (long)(item1.Creditor ?? 0);
        //                    }
        //                    if (item1.Creditor != 0 && item1.tbfkfinancial.numbershomar == 5003)
        //                    {
        //                        value6 += (long)(item1.Creditor ?? 0);
        //                    }



        //                }

        //            }
        //        }
        //    }




        //        //var originalFile = new FileInfo(Server.MapPath("~/Areas/Content/ExcelFiles/excelfish.xlsx"));
        //        //var newFilePath = Server.MapPath("~/Areas/Content/ExcelFiles/excelfish.xlsx");
        //        try
        //    {
        //        using (var package = new ExcelPackage())
        //        {
        //            if (countpdf == 101)
        //            {

        //                // بین هر ۲۰۰ فایل، یه مکث و پاکسازی حافظه انجام بده
        //                GC.Collect();
        //                System.Threading.Thread.Sleep(2000); // ۲ ثانیه مکث

        //            }
        //            // Copy data and formatting from the original file to the new file
        //            using (var stream = new FileStream(originalFile.FullName, FileMode.Open, FileAccess.Read))
        //            {
        //                package.Load(stream);
        //            }

        //            // Get the second sheet of the new file if it exists
        //            var worksheet2 = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Index == 1);

        //            // Check if the worksheet exists
        //            if (worksheet2 != null)
        //            {
        //                // Define a list of strings to be added to the cells in columns C1 to C42
        //                var stringList = new List<string>();
        //                var stringList2 = new List<string>();

        //                var stringList3 = new List<string>();

        //                var stringList4 = new List<string>();
        //                var stringList5 = new List<string>();


        //                #region fill string list
        //                stringList.Add(Model.FishHeader.PersonalCode.ToString(""));
        //                stringList.Add(Model.FishHeader.OnvanShoql);
        //                stringList.Add(Model.FishHeader.NationalCode);
        //                stringList.Add(Model.FishHeader.UserName.ToString());
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد شغل (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کار ( هرساعت-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع مزد مبنا ( روزانه-ریال)").Value.ToString("N0"));
        //                stringList.Add(Model.FishHeader.MonthName);
        //                stringList.Add(Model.FishHeader.year.ToString());
        //                stringList.Add(Model.FishHeader.TedadRoozMonth.ToString("N0"));
        //                var endDate = new DateTime(2025, 3, 20);
        //                var finddd = db.Link_User_And_Peyman.Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode && p.Status == true).FirstOrDefault();


        //                var findob = db.tbUserContracts
        //                    .Where(p => p.tbUsers.usr_Personal_ID == Model.FishHeader.PersonalCode
        //                                && p.usc_EndTime == endDate
        //                                && p.FK_KarfarmaID != null)
        //                    .FirstOrDefault();
        //                string peymanName2 = Model.FishHeader.PeymanName;

        //                if (findob != null)
        //                {
        //                    peymanName2 = findob.tbCompanies.CompanyName;
        //                }
        //                if (finddd != null)
        //                {
        //                    var finj =  db.tbPeymanContracts.Where(p => p.FK_UserTarafDovvom != null && p.pec_ID == finddd.FK_Peyman_ID).FirstOrDefault();
        //                    if (finj != null)
        //                    {
        //                        peymanName2 = db.tbCompanies.Where(p=>p.ID==finj.FK_UserTarafDovvom).Select(s=>s.CompanyName).FirstOrDefault();

        //                    }

        //                }

        //                // جداکننده مورد نظر
        //                char separator2 = '-';

        //                // یافتن موقعیت جداکننده در رشته
        //                int startIndex2 = peymanName2.IndexOf(separator2);
        //                string extractedFileName = peymanName2.Substring(startIndex2 + 1).Trim();
        //                stringList.Add(extractedFileName.ToString());







        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز کارکرد ").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد ساعات اضافه کار").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز ماموریت").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز استعلاجی").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)").Value.ToString("N0"));
        //                stringList2.Add(Model.FishValues.FirstOrDefault(p => p.Title == "تعداد روز غیبت").Value.ToString("N0"));





        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد گروه (شغل)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق تاهل").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ماموریت (ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value.ToString("N0"));
        //                if (Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
        //                {
        //                    stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)").Value.ToString("N0"));

        //                }
        //                else
        //                {
        //                    stringList3.Add("0");

        //                }
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "پاداش نوع 2").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value.ToString("N0"));
        //                stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع ناخالص حقوق و مزایا (ریال)").Value.ToString("N0"));
        //                //stringList3.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));




        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول بیمه (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل مشمول مالیات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value.ToString("N0"));
        //                if(Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)") != null)
        //                {
        //                    stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "بیمه تکمیلی(ریال)").Value.ToString("N0"));

        //                }
        //                else
        //                {
        //                    stringList4.Add("0");

        //                }
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جریمه نوع 2").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "ذخیره کار مازاد (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "جمع کل کسورات (ریال)").Value.ToString("N0"));
        //                stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "خالص قابل دریافت (ریال)").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));
        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));

        //                //stringList4.Add(Model.FishValues.FirstOrDefault(p => p.Title == "").Value.ToString("N0"));


        //                var ex333=db.tbJaremehandpadash.Where(p=>p.Fk_usr==UserID&&p.Month==Month&&p.Year==Year).FirstOrDefault();
        //                string title = "";
        //                if (ex333 != null)
        //                {
        //                    title = ex333.Titel;
        //                }


        //                var tb = noalf
        //         .Where(p => p.MoalfeVal_FKUser ==  UserID

        //                     && p.MoalfeVal_Value != 0
        //                     )
        //         .ToList();
        //                List<string> mode = new List<string>();
                    
        //                foreach (var it in tb)
        //                {
        //                    var yuuu = caran
        //                        .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranStandard != 0)
        //                        .FirstOrDefault();

        //                    if (yuuu == null)
        //                    {
        //                        var u = caran
        //                            .Where(p => p.FK_Moalefe_ID == it.MoalfeVal_FKMoalafeDastmozdi && p.CaranAyabOZahab != 0)
        //                            .FirstOrDefault();

        //                        if (u != null)
        //                        {
        //                            mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)u.CaranAyabOZahab).ToString("N0")}/ 0 /{it.MoalfeVal_Value} )----");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        mode.Add($"{it.tbContractMoalefeDastmozdi.md_Title} ( {((decimal)yuuu.CaranAyabOZahab).ToString("N0")}/ {yuuu.CaranStandard} /{it.MoalfeVal_Value} )---");
        //                    }
        //                }
        //                if (db.tbMaxkarkardMonth.Where(p => p.Month == Month && p.FKUser == UserID && p.Year == Year && p.RemainValue != 0 && p.RemainValue != null).FirstOrDefault() != null)
        //                {
        //                    mode.Add($"(مازاد اضافه کار در حال بررسی میباشد  )----");

        //                }




        //                //var title = worksheet2.Rows[0].Cells;
        //                var List = worksheet2.Rows;
        //                var count = worksheet2.Rows[0].Count();
        //                var count2 = worksheet2.Rows.Count();

        //                //for (int i = 1; i < count; i++)
        //                //{
        //                //    var row = worksheet2.Rows[i];
        //                //    var Name = row.Cells[0];
        //                //} 
        //                int rowCount = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 1].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 1].Value.ToString()))
        //                    {
        //                        rowCount++;
        //                    }
        //                }

        //                int rowCount2 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 3].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 3].Value.ToString()))
        //                    {
        //                        rowCount2++;
        //                    }
        //                }


        //                int rowCount3 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 5].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 5].Value.ToString()))
        //                    {
        //                        rowCount3++;
        //                    }
        //                }

        //                int rowCount4 = 0;
        //                for (int row = 1; row <= worksheet2.Dimension.End.Row; row++)
        //                {
        //                    if (worksheet2.Cells[row, 9].Value != null && !string.IsNullOrEmpty(worksheet2.Cells[row, 9].Value.ToString()))
        //                    {
        //                        rowCount4++;
        //                    }
        //                }




        //                var countColumnARows = rowCount;
        //                for (int i = 13; i <= countColumnARows; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i,1].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex= fish.Where(p=>p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year&&p.FK_User==UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                        else
        //                        {
        //                            stringList.Add("0");
        //                        }


        //                    }

        //                }

        //                #endregion

        //                for (int i = 6; i <= rowCount2; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 3].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year&& p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList2.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }
        //                for (int i = 16; i <= rowCount3; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 5].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year&&p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList3.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }
        //                for (int i = 10; i <= rowCount4; i++)
        //                {
        //                    var row = worksheet2.Rows[i];
        //                    var Name = worksheet2.Cells[i, 9].Value.ToString();
        //                    if (Name != null)
        //                    {
        //                        var ex = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == Name && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_User == UserID).FirstOrDefault();
        //                        if (ex != null && ex.mlfvlfsh_Value != null)
        //                        {
        //                            stringList4.Add(((decimal)ex.mlfvlfsh_Value).ToString("N0"));
        //                        }
        //                    }

        //                }

        //                // Loop through each row in columns C1 to C42 and add the corresponding string from the list
        //                for (int i = 1; i <= stringList.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"B{i}"];
        //                    cell.Value = stringList[i - 1];
        //                }
        //                for (int i = 1; i <= stringList2.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"D{i}"];
        //                    cell.Value = stringList2[i - 1];
        //                }
        //                for (int i = 1; i <= stringList3.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"F{i}"];
        //                    cell.Value = stringList3[i - 1];
        //                }
        //                for (int i = 1; i <= stringList4.Count; i++)
        //                {
        //                    var cell = worksheet2.Cells[$"J{i}"];
        //                    cell.Value = stringList4[i - 1];
        //                }
        //                var cell2 = worksheet2.Cells[$"L{1}"];
        //                cell2.Value = string.Join(" ", mode);
                   
        //                var cell3 = worksheet2.Cells[$"L{4}"];
        //                cell3.Value = value3;
        //                var cell4 = worksheet2.Cells[$"L{5}"];
        //                cell4.Value = value5;
        //                var cell5= worksheet2.Cells[$"L{6}"];
        //                cell5.Value = value6;
        //                var cell6 = worksheet2.Cells[$"L{7}"];
        //                cell6.Value = value4;
        //                var cell7 = worksheet2.Cells[$"L{8}"];
        //                cell7.Value = value2;



        //                var cell8 = worksheet2.Cells[$"L{2}"];
        //                cell8.Value = title;

        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //                // استفاده از فضای خالی به جای Environment.NewLine
        //            }
        //            var worksheet3 = package.Workbook.Worksheets[0];
        //            if (worksheet3 != null)
        //            {

        //                worksheet3.Calculate();


        //                var cellsWithFormulas = worksheet3.Cells[worksheet3.Dimension.Address]
        //                    .Where(c => !string.IsNullOrEmpty(c.Formula)).ToList();

        //                if (cellsWithFormulas.Any())
        //                {
        //                    foreach (var cell in cellsWithFormulas)
        //                    {
        //                        var value1232 = cell.Value;
        //                        cell.Value = value1232;
        //                    }
        //                }
        //                else
        //                {
        //                    System.Diagnostics.Debug.WriteLine("هیچ فرمولی در شیت پیدا نشد!");
        //                }

        //                var newFile = new FileInfo(newFilePath);
        //                package.SaveAs(newFile);
        //            }
        //            // Save changes to the new file

        //        }
        //        using (ExcelEngine excelEngine = new ExcelEngine())
        //        {
        //            //IApplication application = excelEngine.Excel;
        //            //application.DefaultVersion = ExcelVersion.Xlsx;
        //            //FileStream excelStream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77New.xlsx"), FileMode.Open, FileAccess.Read);
        //            //IWorkbook workbook = application.Workbooks.Open(excelStream);
        //            //IWorksheet worksheet = workbook.Worksheets[0];

        //            //Initialize XlsIO renderer.
        //            //   XlsIORenderer renderer = new XlsIORenderer();

        //            //Convert Excel document into PDF document 
        //            //   PdfDocument pdfDocument = renderer.ConvertToPDF(worksheet);

        //            //    Stream stream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), FileMode.Create, FileAccess.ReadWrite);
        //            //    pdfDocument.Save(stream);
        //            //  pdfDocument.Close();
        //            string peymanName = Model.FishHeader.PeymanName;

        //            // جداکننده مورد نظر
        //            //char separator = '-';

        //            // یافتن موقعیت جداکننده در رشته
        //            //int startIndex = peymanName.IndexOf(separator);
        //            //string extractedFileName = peymanName.Substring(startIndex + 1).Trim();

        //            string extractedFileName = peymanName;

        //            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(newFilePath);
                   
        //            for (int i = 1; i < workbook.Worksheets.Count; i++)
        //            {
        //                workbook.Worksheets[i].IsVisible = false;
        //            }


        //            //workbook.Save(Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName+'-'+Month+'-'+Year + ".pdf"), Aspose.Cells.SaveFormat.Pdf);

        //            System.Diagnostics.Debug.WriteLine("تبدیل به PDF با موفقیت انجام شد.");
        //            //Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        //            //Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/sistan2.xlsx"));
        //            //string outputPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode +'-'+ extractedFileName + ".pdf");
        //            ////outputPath = outputPath.Replace("\\", "/");

        //            //wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath);

        //            //// Save the workbook as PDF
        //            ////string pdfPath = Server.MapPath($"~/Areas/Contracts/Contents/fish/{Model.FishHeader.PersonalCode}_{Model.FishHeader.PeymanName}.pdf");
        //            ////wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, pdfPath);
        //            ////wb.Sheets[1].ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode  + ".pdf"));

        //            string outputPdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + Month + '-' + Year + "_Clean.pdf");
        //            string pdfPath = Server.MapPath("~/Areas/Contracts/Contents/fish/" + Model.FishHeader.PersonalCode + '-' + extractedFileName + '-' + Month + '-' + Year + ".pdf");
        //            workbook.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);

        //            RemoveWatermarkFromPDF(pdfPath, outputPdfPath);

        //            System.Diagnostics.Debug.WriteLine("واترمارک با موفقیت حذف شد!");

        //            // 🛑 اینجا می‌توانی فایل بدون واترمارک را به کاربر برگردانی، مثلا:
        //            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPdfPath);
        //            //return File(fileBytes, "application/pdf", "Cleaned_File.pdf");

        //            // 🔹 تابع حذف واترمارک

        //            return 0;


        //            ////  pdfDocument.Dispose();
        //            ////   stream.Dispose();
        //            //wb.Close(false);

        //            //// Quit Excel application
        //            //excelApp.Quit();


        //            // byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"));
        //            // return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "frm77.pdf");

        //            // return File(stream, "Forum77.pdf");

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        viewcontinue(countpdf);
        //        return -1;

        //        //wb.Close(false);
        //        // Handle any exceptions that occur during the file creation process

        //    }

        //}
        public ActionResult viewcontinue(int countpdf)
        {
            ViewBag.CountPdf = countpdf;
            return View("~/Areas/Contracts/Views/UserContract/viewcontinue.cshtml");
        }
        public void RemoveWatermarkFromPDF(string inputPdfPath, string outputPdfPath)
        {
            PdfReader reader = new PdfReader(inputPdfPath);
            PdfStamper stamper = new PdfStamper(reader, new FileStream(outputPdfPath, FileMode.Create));

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                PdfContentByte canvas = stamper.GetOverContent(i);
                canvas.SetColorFill(BaseColor.WHITE);
                //canvas.Rectangle(0, 830, 530, 40); // موقعیت تقریبی واترمارک
                canvas.Rectangle(0, 830, 530, 30);

                canvas.Fill();
            }

            stamper.Close();
            reader.Close();
        }






















        public ActionResult EXCEL_GHARARDAD()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel4();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "EXCEL_GHARARDAD" + extension);

        }
        private Workbook SetDataToMachinsOrToolsExcel4()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EXCELGHARADAD.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var usr=db.tbUsers.Where(p=>p.usr_Personal_ID!=null).ToList();
            var company = db.tbCompanies.ToList();
            var job = db.tbjob.ToList();
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
            foreach (var item in moalfe)
            {
                Row = new Row() { Height = 20, Index = count2 };
                Row.AddCells(new List<Cell>()
                {
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
                        Index = 0
                    },
                        
                });
                count2++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            foreach (var item in company)
            {
                Row = new Row() { Height = 20, Index = counter };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.CompanyName+"-"+(item.tbUsers.FullName),
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
                        Value = item.ID,
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
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            foreach (var item in job)
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
                        Index = 6
                    },
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
                        Index = 7
                    },
                });
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            return Moalefeexcelfile;








        }






        #endregion

        #region قرار داد های اجاره ای

        //public bool UserRent_Create(tbUserRentContracts obj_userrent)
        //{

        //    return rep_userrentContracts.Create(obj_userrent);
        //}

        //public bool UserRent_Update(tbUserRentContracts obj_userrent)
        //{

        //    return rep_userrentContracts.Create(obj_userrent);
        //}
        //[AuthorizeAAA]
        //public bool UserRent_Accept(int userContractID)
        //{
        //    return rep_userrentContracts.AcceptUserContract(userContractID);
        //}
        //[AuthorizeAAA]
        //public string UserRent_Delete(int userContractID)
        //{
        //    return rep_userrentContracts.Disable(userContractID); // delete
        //}
        //[AuthorizeAAA]
        //public bool UserRent_UploadFile(int ContractrentID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        //{
        //    if (files != null)
        //    {
        //        foreach (var file in files)
        //        {

        //            if (file.ContentLength > 0)
        //            {
        //                var segment = file.FileName.Split('.');
        //                string file_type = segment[segment.Length - 1];
        //                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + file_type).ToString();
        //                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Content/ContractFiles/" + filename));
        //                var contract = rep_userrentContracts.Find(ContractrentID);
        //                contract.urc_FileSystemName = filename;
        //                contract.urc_FileName = file.FileName;
        //                return rep_userrentContracts.Update(contract);
        //            }
        //            else
        //            {
        //                return true;
        //            }
        //        }
        //        return true;
        //    }
        //    else
        //    {
        //        return true;
        //    }

        //}
        #endregion

        [AuthorizeAAA]
        public ActionResult DownloadPDF()
        {
            try
            {



                string extension = ".pdf";

                //var stream2 = new MemoryStream();

                //var mimeType = MimeTypes.ByExtension[extension];

                //return File(stream2.ToArray(), mimeType, "frm77" + extension);

                Stream stream = new FileStream(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/Forum77.pdf"), FileMode.Open, FileAccess.ReadWrite);
                return File(stream,extension,"forum.pdf");

            }
            catch (Exception e)
            {

                return null;
            }
            //System.IO.MemoryStream stream = new System.IO.MemoryStream();
            //var workbook = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/ContractsFile/test2.xlsx"));
            //workbook.Save(stream, ".xlsx");

            //return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "form77.xlsx");



        }









        #region









        public class ServiceResult
        {
            public bool Success { get; set; }

            public string Message { get; set; }

            public object Data { get; set; }
        }



        public class BaleUpdate
        {
            public long update_id { get; set; }
            public BaleMessage message { get; set; }
        }

        public class BaleMessage
        {
            public long message_id { get; set; }
            public BaleUser from { get; set; }
            public BaleChat chat { get; set; }
            public string text { get; set; }
        }

        public class BaleUser
        {
            public long id { get; set; }
            public string first_name { get; set; }
            public string last_name { get; set; }
            public string username { get; set; }
        }

        public class BaleChat
        {
            public long id { get; set; }
            public string type { get; set; }
            public string username { get; set; }
            public string first_name { get; set; }
            public string last_name { get; set; }
        }




        public class BaleApiResponse<T>
        {
            public bool ok { get; set; }

            public T result { get; set; }

            public string description { get; set; }

            public int error_code { get; set; }
        }


        public class BaleBotService
        {
            private static readonly HttpClient _client = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(30)
            };

            private readonly string _token;
            private readonly string _baseUrl;

            public BaleBotService()
            {
                _token = System.Configuration.ConfigurationManager.AppSettings["BaleBotToken"];
                _baseUrl = "https://tapi.bale.ai";
            }

            public async Task<ServiceResult> SendMessageAsync(long chatId, string text)
            {
                var url = $"{_baseUrl}/bot{_token}/sendMessage";

                var payload = new
                {
                    chat_id = chatId,
                    text = text
                };

                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var response = await _client.PostAsync(url, content);
                    var responseText = await response.Content.ReadAsStringAsync();

                    var apiResponse = JsonConvert.DeserializeObject<BaleApiResponse<object>>(responseText);

                    if (apiResponse != null && apiResponse.ok)
                    {
                        return new ServiceResult { Success = true, Message = "پیام ارسال شد" };
                    }
                    else
                    {
                        return new ServiceResult { Success = false, Message = apiResponse?.description ?? "خطای ارسال" };
                    }
                }
                catch (Exception ex)
                {
                    return new ServiceResult { Success = false, Message = ex.Message };
                }
            }
        }


        private readonly SaabEntities _db = new SaabEntities();
 private readonly BaleBotService _baleService = new BaleBotService();

        public ActionResult baleview()
        {
            // فقط کاربرانی که در جدول UserBaleAccounts هستند را می‌آوریم
            var users = (from u in _db.tbUsers
                         join b in _db.UserBaleAccounts on u.usr_ID equals b.UserId
                         where b.IsActive && b.IsVerified
                         
            select new UserBaleViewModel
                         {
                             UserId = u.usr_ID,
                             FullName = u.usr_Name + " " + u.usr_Family,
                             ChatId = b.BaleChatId,
                             BaleUsername = b.BaleUsername
                         }).ToList();

            return View(users);
        }







        [HttpPost]
        public async Task<JsonResult> SendQuickMessage(long chatId, string message)
        {
            if (string.IsNullOrEmpty(message))
                return Json(new { success = false, message = "متن پیام خالی است" });

            var result = await _baleService.SendMessageAsync(chatId, message);
            return Json(new { success = result.Success, message = result.Message });
        }
    

    // یک کلاس ساده برای انتقال داده به View
    //public class UserBaleViewModel
    //{
    //    public int UserId { get; set; }
    //    public string FullName { get; set; }
    //    public long ChatId { get; set; }
    //    public string BaleUsername { get; set; }
    //}







    // GET: GenerateLinkCode
    public ActionResult GenerateLinkCode(int userId)
    {
        // یک کد رندوم ۴ رقمی بسازیم
        var random = new Random();
        var code = "BALE-" + random.Next(1000, 9999);

        // حذف کدهای قبلی استفاده نشده این کاربر
        var oldCodes = _db.BaleLinkRequests
                          .Where(x => x.UserId == userId && !x.IsUsed)
                          .ToList();

        foreach (var item in oldCodes)
        {
            item.IsUsed = true;
            item.UsedAt = DateTime.Now;
        }

        // ساخت رکورد جدید
        var newRequest = new BaleLinkRequests
        {
            UserId = userId,
            LinkCode = code,
            IsUsed = false,
            ExpiresAt = DateTime.Now.AddMinutes(10), // کد به مدت ۱۰ دقیقه معتبر است
            CreatedAt = DateTime.Now
        };

        db.BaleLinkRequests.Add(newRequest);
        db.SaveChanges();

        // نمایش کد به کاربر
        return Json(new
        {
            success = true,
            code = code,
            message = "لطفا این کد را در ربات بله ارسال کنید."
        }, JsonRequestBehavior.AllowGet);
    }


















    [HttpPost]
    public async Task<ActionResult> Index5()
    {
        string body;
        using (var reader = new StreamReader(Request.InputStream))
        {
            body = await reader.ReadToEndAsync();
        }

        var update = JsonConvert.DeserializeObject<BaleUpdate>(body);

        if (update?.message == null || string.IsNullOrWhiteSpace(update.message.text))
            return new HttpStatusCodeResult(200);

        var chatId = update.message.chat.id;
        var fromUserId = update.message.from.id;
        var text = update.message.text.Trim();

        var baleService = new BaleBotService();

        if (text.StartsWith("BALE-"))
        {
            // بررسی کد اتصال
            var linkRequest = _db.BaleLinkRequests
                .FirstOrDefault(x => x.LinkCode == text && !x.IsUsed && x.ExpiresAt > DateTime.Now);

            if (linkRequest == null)
            {
                await baleService.SendMessageAsync(chatId, "کد اتصال نامعتبر یا منقضی شده است.");
                return new HttpStatusCodeResult(200);
            }

            var userAccount = _db.UserBaleAccounts.FirstOrDefault(x => x.UserId == linkRequest.UserId);
            if (userAccount == null)
            {
                userAccount = new UserBaleAccounts
                {
                    UserId = linkRequest.UserId,
                    BaleChatId = chatId,
                    BaleUserId = fromUserId,
                    BaleUsername = update.message.from?.username,
                    BaleFirstName = update.message.from?.first_name,
                    BaleLastName = update.message.from?.last_name,
                    IsActive = true,
                    IsVerified = true,
                    ConnectedAt = DateTime.Now
                };

                _db.UserBaleAccounts.Add(userAccount);
            }
            else
            {
                userAccount.BaleChatId = chatId;
                userAccount.BaleUserId = fromUserId;
                userAccount.BaleUsername = update.message.from?.username;
                userAccount.BaleFirstName = update.message.from?.first_name;
                userAccount.BaleLastName = update.message.from?.last_name;
                userAccount.IsActive = true;
                userAccount.IsVerified = true;
                userAccount.LastInteractionAt = DateTime.Now;
            }

            linkRequest.IsUsed = true;
            linkRequest.UsedAt = DateTime.Now;

            await _db.SaveChangesAsync();

            await baleService.SendMessageAsync(chatId, "اتصال حساب شما با موفقیت انجام شد.");
        }
        else
        {
            // پیام‌های دیگر
            await baleService.SendMessageAsync(chatId, "پیام دریافت شد");
        }

        return new HttpStatusCodeResult(200);
    }









    public ActionResult karbar()
    {
        var cookie = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
        if (cookie == null) return new HttpStatusCodeResult(HttpStatusCode.BadRequest);

        string nationalCode = Utility.Base64.Base64Decode(cookie.Value);
        var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalCode);

        if (user == null) return HttpNotFound();

        return View(user.usr_ID);
    }
    #endregion
}
}