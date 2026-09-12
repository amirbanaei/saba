using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Functions.Salaries.Formula;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Areas.Contracts.Models.Classes;
using SaabWebProject.Utility;
using SaabWebProject.Models.Utilitis;
using Telerik.Web.Spreadsheet;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using Rotativa.MVC;
using System.Data.Entity;
using SaabWebProject.Models.Repositories.Salaries.Formula;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using SaabWebProject.Models.ViewModels.Statements;
using System;

using static Stimulsoft.Report.StiRecentConnections;
using System.Windows.Controls;
using System.Diagnostics.Eventing.Reader;
using Stimulsoft.Controls.Win.DotNetBar;
using Syncfusion.XlsIO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using static Stimulsoft.Report.StiOptions.Designer.ComponentsTypes;
using OfficeOpenXml;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.Contracts.UserContracts;
using SaabWebProject.Utility;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Web.Spreadsheet;
using Microsoft.Office.Interop;
using Telerik.Windows.Documents.Spreadsheet.Model.DataValidation;
using System.Windows.Data;
using OfficeOpenXml.Drawing.Style.ThreeD;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Web.Configuration;
using System.Diagnostics;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Web.Mvc;
namespace SaabWebProject.Areas.Salaries.Controllers
{
    public class printfishController : Controller
    {
        // GET: Salaries/printfish
        private readonly IConverter _converter;

        public printfishController(IConverter converter)
        {
            _converter = converter;
        }
        public ActionResult Index()
        {
            return View();
        }
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
        #region متغیرها
        SaabEntities db;
        tbFormulaRepositories FormulaRepo;
        tbContractMoalefeDastMozdiRepository MoalefeDastmozdiRepo;
        tbUsersRepository userRepo;
        tbMoalefeValueRepository MoalefeValueRepo;
        tbMoalefeDastmozdiValueFromExcelRepository moalefewithexcelRepo;
        static DateTime datetime = new DateTime();
        #endregion
        public ActionResult ShowFish(int UserID, int Month, int Year)
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
                        #region فیش دارید یا خیر
                        var IDkhales = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                        if (!db.tbMoalefeValueFish.Any(p => p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == IDkhales && p.FK_User == UserID))
                        {
                            return Content("-1");
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
                                            value = (long)(item1.Creditor - item1.Debtore);
                                        }


                                        else if (item1.Debtore != 0 && item1.Debtore != null)
                                        {
                                            value2 = (long)(item1.Debtore ?? 0);
                                        }
                                        else if (item1.Creditor != 0 && item1.Creditor != null)
                                        {
                                            value = (long)(item1.Creditor ?? 0);

                                        }

                                    }

                                }
                            }
                        }



























                        FishHeader fishHeader = new FishHeader
                        {
                            PeymanName = db.tbPeymanContracts.FirstOrDefault(p => p.pec_ID == peymanID && p.Inactive != true).pec_Title,
                            OnvanShoql = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID).usc_Jobtitle,
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
                            FishValue moalefeqaradadfish = new FishValue
                            {
                                Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi).md_Title,
                                Value = (int)item.Value
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

                            FishValue ezafvalue = new FishValue
                            {
                                Title = item.md_Title,
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

                            FishValue kasrvalue = new FishValue
                            {
                                Title = item.md_Title,
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
                                Value = (double)item.mlfvlfsh_Value,
                                Title = item.tbContractMoalefeDastmozdi.md_Title
                            };
                            lstfshvalue.Add(fshvalue);
                        }
                        Model.FishValues = lstfshvalue;
                        #endregion
                        return PartialView("FishView2", Model);
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
            catch (Exception)
            {

                return Content("-2");//در نمایش فیش خطایی رخ داده است
            }

        }






        protected void btnSaveAndConvert_Click()
        {
            // مسیر فایل اکسل که می‌خواهیم آن را ایجاد کنیم و داده‌ها را در آن ذخیره کنیم
            string excelFilePath = Server.MapPath("~/Content/ExcelFiles/MAXKARKARD.xlsx");

            // ایجاد یک بسته جدید اکسل
            //using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
            //{
            //    // انتخاب ورق اول فایل
            //    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

            //    // اضافه کردن داده‌ها به ورق
            //    worksheet.Cells["A1"].Value = "نام";
            //    worksheet.Cells["B1"].Value = "نام خانوادگی";

            //    // ذخیره تغییرات
            //    excelPackage.Save();
            //}

            // حالا که داده‌ها را به فایل اکسل ذخیره کردید، می‌توانیم آن را به فایل PDF تبدیل کنیم
            ConvertExcelToPDF(excelFilePath);
        }

        private void ConvertExcelToPDF(string excelFilePath)
        {
            // مسیر فایل PDF که می‌خواهیم آن را ایجاد کنیم
            string pdfFilePath = Server.MapPath("~/Content/ExcelFiles/MAXKARKARD.xlsx");

            // خواندن داده‌ها از فایل اکسل
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // ورق اول فایل

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // ایجاد یک سند جدید PDF
                Document pdfDoc = new Document();
                PdfWriter.GetInstance(pdfDoc, new FileStream(pdfFilePath, FileMode.Create));

                pdfDoc.Open();

                // اضافه کردن داده‌ها به سند PDF
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        object cellValue = worksheet.Cells[row, col].Value;
                        pdfDoc.Add(new Paragraph(cellValue != null ? cellValue.ToString() : ""));
                    }
                    pdfDoc.Add(Chunk.NEWLINE); // خط جدید بین هر ردیف
                }

                pdfDoc.Close();
            }

            // نمایش لینک دانلود فایل PDF به کاربر
            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AppendHeader("Content-Disposition", "attachment; filename=ExcelData.pdf");
            Response.TransmitFile(pdfFilePath);
            Response.End();
        }









































       

    }
}