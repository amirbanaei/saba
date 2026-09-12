using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.ManagementAccounting;
using SaabWebProject.Utility;
using System;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.Utilitis;
using System.Net.NetworkInformation;
using ExcelLibrary.BinaryFileFormat;
using System.Globalization;
using System.IO;
using Telerik.Web.Spreadsheet;
using System.Data.Entity;
using System.Windows.Media.Animation;
using Stimulsoft.Controls.Win.UI.ContentManager;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Ajax.Utilities;
using Syncfusion.XlsIO.Implementation.XmlSerialization;
using DocumentFormat.OpenXml.VariantTypes;
using SaabWebProject.Models.ViewModels.Contracts.Function;
using System.Web.UI.WebControls;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.ComponentModel.DataAnnotations;
using Newtonsoft.Json;
using Org.BouncyCastle.Bcpg;

namespace SaabWebProject.Areas.ManagementAccounting.Controllers
{
    public class FinancialDocumentsController : Controller
    {
        SaabEntities db = new SaabEntities();
        #region Finnail
        // GET: ManagementAccounting/FinancialDocuments5






        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult _Submit_Detail()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/_Submit_Detail.cshtml");
        }

        FinancialDocumentsRepositories rep_def_rsn = new FinancialDocumentsRepositories();

        public ActionResult documents_List()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/documents_List.cshtml", rep_def_rsn.Update());
        }

        public ActionResult _ListAlltbDetailedtitle()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/_ListAlltbDetailedtitle.cshtml", rep_def_rsn.List());
        }

        public ActionResult _CreatetbDetailedtitle()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/_CreatetbDetailedtitle.cshtml");
        }
        public ActionResult Batchuser()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/Batchuser.cshtml");
        }
        public ActionResult Batchpymn()
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/Batchpymn.cshtml");
        }
        [HttpGet]
        public ActionResult GetMaxTarif()
        {
            var fii = db.dbRelatinafrad.Max(s => s.tarif);
            fii += 1; // افزایش یک واحد
            return Json(new { fii = fii }, JsonRequestBehavior.AllowGet);
        }
        public string Testpardas(DateTime? DataDocument = null, int creditline = 0, string EndTime2 = "", string title = "", int tar = 0, bool IsVisible = false, bool IsVisible2 = false)
        {
            TimeSpan? totalTimeDifference101 = TimeSpan.Zero;
            TimeSpan? totalTimeDifference104 = TimeSpan.Zero;
            List<dbRelatinafrad> Flters2 = new List<dbRelatinafrad>();

            TimeSpan? totalTimeDifference10 = TimeSpan.Zero;
            TimeSpan? totalTimeDifference103 = TimeSpan.Zero;
            totalTimeDifference103 = TimeSpan.Parse(EndTime2);
            totalTimeDifference10 = TimeSpan.Parse(EndTime2);
            totalTimeDifference101 = TimeSpan.Parse(EndTime2);
            if (IsVisible == true)
            {
                TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference54 = TimeSpan.Zero;

                totalTimeDifference2 = TimeSpan.Parse(EndTime2);

                // تبدیل تاریخ شمسی به میلادی
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime gregorianDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                int dayy;
                switch (dayOfWeek)
                {
                    case DayOfWeek.Saturday:
                        dayy = 1; // شنبه
                        break;
                    case DayOfWeek.Sunday:
                        dayy = 2; // یکشنبه
                        break;
                    case DayOfWeek.Monday:
                        dayy = 3; // دوشنبه
                        break;
                    case DayOfWeek.Tuesday:
                        dayy = 4; // سه‌شنبه
                        break;
                    case DayOfWeek.Wednesday:
                        dayy = 5; // چهارشنبه
                        break;
                    case DayOfWeek.Thursday:
                        dayy = 6; // پنج‌شنبه
                        break;
                    case DayOfWeek.Friday:
                        dayy = 7; // جمعه
                        break;
                    default:
                        dayy = 0; // خطا یا مقدار نامعتبر
                        break;
                }
                if (db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault() != null)
                {
                    return "3";

                }
                var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                foreach (var t in dayyu)
                {
                    totalTimeDifference3 += (t.end_clock - t.start_clock);

                }
                var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                if (moraghasi != null)
                {
                    if (moraghasi.end_clock != null)
                    {
                        totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);

                    }
                    if (moraghasi.end_clock2 != null)
                    {
                        totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);

                    }
                    if (totalTimeDifference3 <= TimeSpan.Zero)
                    {
                        return "4";

                    }

                }

                //// استفاده از HijriCalendar برای تبدیل تاریخ میلادی به قمری
                //HijriCalendar hijriCalendar = new HijriCalendar();
                //int hijriYear = hijriCalendar.GetYear(gregorianDate);
                //int hijriMonth = hijriCalendar.GetMonth(gregorianDate);
                //int hijriDay = hijriCalendar.GetDayOfMonth(gregorianDate);
                var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.DataDocument == gregorianDate && p.delet != true).ToList();
                foreach (var fiiii in find)
                {
                    totalTimeDifference1 += fiiii.clocktime;
                }
                totalTimeDifference3 -= totalTimeDifference1;
                totalTimeDifference3 -= totalTimeDifference2;
                if (totalTimeDifference3 <= TimeSpan.Zero)
                {
                    return totalTimeDifference3.ToString();


                }

                else
                {
                    return "0";

                }

                // نمایش و ارسال تاریخ قمری به دیتابیس
                //string hijriDate = $"{hijriYear}/{hijriMonth}/{hijriDay}";
                //Console.WriteLine($"تاریخ قمری: {hijriDate}");

                // ارسال به دیتابیس - SendToDatabase(hijriDate);
            }
            else if (IsVisible == false)
            {

                PersianCalendar persianCalendar = new PersianCalendar();

                DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                DateTime today2 = DateTime.Now.Date;
                TimeSpan currentTime = DateTime.Now.TimeOfDay;
                currentTime = currentTime.Add(new TimeSpan(0, 15, 0)); // اضافه کردن 15 دقیقه

                DateTime today = DateTime.Now.Date;
                while (today <= targetDate)
                {
                    if (today == today2)
                    {
                        bool vaz = true;


                        TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference54 = TimeSpan.Zero;
                        totalTimeDifference2 = TimeSpan.Parse(EndTime2);
                        DateTime gregorianDate = today;
                        var findday = db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault();
                        if (findday != null)
                        {
                            if (IsVisible2 == true)
                            {
                                vaz = true;
                            }
                            else
                            {
                                vaz = false;

                            }
                            //TimeSpan? totalTimeDifference300 = TimeSpan.Zero;

                            //var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                            //DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            //int dayy;
                            //switch (dayOfWeek)
                            //{
                            //    case DayOfWeek.Saturday:
                            //        dayy = 1; // شنبه
                            //        break;
                            //    case DayOfWeek.Sunday:
                            //        dayy = 2; // یکشنبه
                            //        break;
                            //    case DayOfWeek.Monday:
                            //        dayy = 3; // دوشنبه
                            //        break;
                            //    case DayOfWeek.Tuesday:
                            //        dayy = 4; // سه‌شنبه
                            //        break;
                            //    case DayOfWeek.Wednesday:
                            //        dayy = 5; // چهارشنبه
                            //        break;
                            //    case DayOfWeek.Thursday:
                            //        dayy = 6; // پنج‌شنبه
                            //        break;
                            //    case DayOfWeek.Friday:
                            //        dayy = 7; // جمعه
                            //        break;
                            //    default:
                            //        dayy = 0; // خطا یا مقدار نامعتبر
                            //        break;
                            //}
                            //var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            //foreach (var t in dayyu)
                            //{
                            //    totalTimeDifference300 += (t.end_clock - t.start_clock);
                            //}
                            //if (totalTimeDifference300 > TimeSpan.Zero)
                            //{
                            //    vaz = true;
                            //}
                            //else
                            //{
                            //    vaz = false;
                            //}


                        }
                        // تبدیل تاریخ شمسی به میلادی
                        PersianCalendar persianCalendar2 = new PersianCalendar();
                        if (vaz == true)
                        {
                            var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                            DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                            int dayy;
                            switch (dayOfWeek)
                            {
                                case DayOfWeek.Saturday:
                                    dayy = 1; // شنبه
                                    break;
                                case DayOfWeek.Sunday:
                                    dayy = 2; // یکشنبه
                                    break;
                                case DayOfWeek.Monday:
                                    dayy = 3; // دوشنبه
                                    break;
                                case DayOfWeek.Tuesday:
                                    dayy = 4; // سه‌شنبه
                                    break;
                                case DayOfWeek.Wednesday:
                                    dayy = 5; // چهارشنبه
                                    break;
                                case DayOfWeek.Thursday:
                                    dayy = 6; // پنج‌شنبه
                                    break;
                                case DayOfWeek.Friday:
                                    dayy = 7; // جمعه
                                    break;
                                default:
                                    dayy = 0; // خطا یا مقدار نامعتبر
                                    break;
                            }



                            var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            foreach (var t in dayyu)
                            {
                                if (t.end_clock > currentTime)
                                {
                                    if (t.start_clock > currentTime)
                                    {
                                        totalTimeDifference3 += (t.end_clock - t.start_clock);

                                    }
                                    else
                                    {
                                        totalTimeDifference3 += (t.end_clock - currentTime);

                                    }
                                }
                                //totalTimeDifference3 += (t.end_clock - t.start_clock);
                            }

                            var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                            if (moraghasi != null)
                            {
                                if (moraghasi.end_clock != null)
                                {

                                    if (moraghasi.end_clock > currentTime)
                                    {
                                        if (moraghasi.start_clock > currentTime)
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);

                                        }
                                        else
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock - currentTime);

                                        }
                                    }
                                }
                                if (moraghasi.end_clock2 != null)
                                {

                                    if (moraghasi.end_clock2 > currentTime)
                                    {
                                        if (moraghasi.start_clock2 > currentTime)
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);

                                        }
                                        else
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock2 - currentTime);

                                        }
                                    }
                                }
                                //if (moraghasi.end_clock != null)
                                //{
                                //    totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);
                                //}
                                //if (moraghasi.end_clock2 != null)
                                //{
                                //    totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);
                                //}

                            }

                            var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.delet != true && p.DataDocument == gregorianDate && p.FK_tbcreatfaal != null && p.delet != true).ToList();
                            foreach (var fiiii in find)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }
                            var find22 = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.delet != true && p.Title != title && p.DataDocument == gregorianDate && p.FK_tbcreatfaal == null && p.tarif == tar && p.delet != true).ToList();
                            foreach (var fiiii in find22)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }

                            totalTimeDifference3 -= totalTimeDifference1;

                            if (totalTimeDifference3 > TimeSpan.Zero)
                            {
                                if (totalTimeDifference10 > TimeSpan.Zero)
                                {

                                    totalTimeDifference10 -= totalTimeDifference3;

                                    if (totalTimeDifference10 <= TimeSpan.Zero)
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {
                                            dbRelatinafrad Flters = new dbRelatinafrad();
                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference103;
                                            Flters.Title = title;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters.delet = false;

                                            Flters.tarif = tar;
                                            Flters2.Add(Flters);
                                        }

                                        //return "0";
                                    }
                                    else
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {

                                            dbRelatinafrad Flters = new dbRelatinafrad();

                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference3;
                                            Flters.Title = title;
                                            totalTimeDifference103 -= totalTimeDifference3;
                                            Flters.tarif = tar;
                                            Flters.delet = false;

                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters2.Add(Flters);
                                        }

                                    }
                                }

                            }
                            else
                            {
                                totalTimeDifference104 -= totalTimeDifference10;


                            }

                            // اضافه کردن یک روز به تاریخ امروز
                        }
                        today = today.AddDays(1);

                    }
                    else
                    {
                        bool vaz = true;

                        TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference54 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference504 = TimeSpan.Zero;

                        totalTimeDifference2 = TimeSpan.Parse(EndTime2);
                        DateTime gregorianDate = today;
                        var findday = db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault();
                        if (findday != null)
                        {
                            //TimeSpan? totalTimeDifference300 = TimeSpan.Zero;

                            //var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                            //DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            //int dayy;
                            //switch (dayOfWeek)
                            //{
                            //    case DayOfWeek.Saturday:
                            //        dayy = 1; // شنبه
                            //        break;
                            //    case DayOfWeek.Sunday:
                            //        dayy = 2; // یکشنبه
                            //        break;
                            //    case DayOfWeek.Monday:
                            //        dayy = 3; // دوشنبه
                            //        break;
                            //    case DayOfWeek.Tuesday:
                            //        dayy = 4; // سه‌شنبه
                            //        break;
                            //    case DayOfWeek.Wednesday:
                            //        dayy = 5; // چهارشنبه
                            //        break;
                            //    case DayOfWeek.Thursday:
                            //        dayy = 6; // پنج‌شنبه
                            //        break;
                            //    case DayOfWeek.Friday:
                            //        dayy = 7; // جمعه
                            //        break;
                            //    default:
                            //        dayy = 0; // خطا یا مقدار نامعتبر
                            //        break;
                            //}
                            //var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            //foreach (var t in dayyu)
                            //{
                            //    totalTimeDifference300 += (t.end_clock - t.start_clock);
                            //}
                            //if (totalTimeDifference300 > TimeSpan.Zero)
                            //{
                            //    vaz = true;
                            //}
                            //else
                            //{
                            //    vaz = false;
                            //}
                            if (IsVisible2 == true)
                            {
                                vaz = true;
                            }
                            else
                            {
                                vaz = false;

                            }

                        }
                        // تبدیل تاریخ شمسی به میلادی
                        PersianCalendar persianCalendar2 = new PersianCalendar();
                        if (vaz == true)
                        {
                            var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                            DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                            int dayy;
                            switch (dayOfWeek)
                            {
                                case DayOfWeek.Saturday:
                                    dayy = 1; // شنبه
                                    break;
                                case DayOfWeek.Sunday:
                                    dayy = 2; // یکشنبه
                                    break;
                                case DayOfWeek.Monday:
                                    dayy = 3; // دوشنبه
                                    break;
                                case DayOfWeek.Tuesday:
                                    dayy = 4; // سه‌شنبه
                                    break;
                                case DayOfWeek.Wednesday:
                                    dayy = 5; // چهارشنبه
                                    break;
                                case DayOfWeek.Thursday:
                                    dayy = 6; // پنج‌شنبه
                                    break;
                                case DayOfWeek.Friday:
                                    dayy = 7; // جمعه
                                    break;
                                default:
                                    dayy = 0; // خطا یا مقدار نامعتبر
                                    break;
                            }



                            var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            foreach (var t in dayyu)
                            {
                                totalTimeDifference3 += (t.end_clock - t.start_clock);
                            }

                            var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                            if (moraghasi != null)
                            {
                                if (moraghasi.end_clock != null)
                                {
                                    totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);
                                }
                                if (moraghasi.end_clock2 != null)
                                {
                                    totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);
                                }

                            }

                            var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.DataDocument == gregorianDate && p.FK_tbcreatfaal != null && p.delet != true).ToList();
                            foreach (var fiiii in find)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }
                            var find22 = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.Title != title && p.DataDocument == gregorianDate && p.FK_tbcreatfaal == null && p.tarif == tar && p.delet != true).ToList();
                            foreach (var fiiii in find22)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }

                            totalTimeDifference3 -= totalTimeDifference1;

                            if (totalTimeDifference3 > TimeSpan.Zero)
                            {
                                if (totalTimeDifference10 > TimeSpan.Zero)
                                {
                                    totalTimeDifference504 = totalTimeDifference10;

                                    totalTimeDifference10 -= totalTimeDifference3;
                                    if (totalTimeDifference10 <= TimeSpan.Zero)
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {
                                            dbRelatinafrad Flters = new dbRelatinafrad();
                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference103;
                                            Flters.Title = title;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;

                                            Flters.tarif = tar;
                                            Flters2.Add(Flters);
                                        }

                                        //return "0";
                                    }
                                    else
                                    {
                                        //if(today== targetDate)
                                        //{
                                        //    if(totalTimeDifference10 == TimeSpan.Zero)
                                        //    {
                                        //        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        //        if (finnnn == null)
                                        //        {

                                        //            dbRelatinafrad Flters = new dbRelatinafrad();

                                        //            Flters.DataDocument = today;
                                        //            Flters.DataDocument2 = targetDate;
                                        //            Flters.clocktime30 = totalTimeDifference101;
                                        //            Flters.clocktime = totalTimeDifference3;
                                        //            Flters.Title = title;
                                        //            totalTimeDifference103 -= totalTimeDifference3;
                                        //            Flters.tarif = tar;
                                        //            Flters.FK_usr = fin.FK_usr;
                                        //            Flters.FK_Relationperson = creditline;
                                        //            Flters2.Add(Flters);
                                        //        }
                                        //    }
                                        //    else
                                        //    {
                                        //        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        //        if (finnnn == null)
                                        //        {

                                        //            dbRelatinafrad Flters = new dbRelatinafrad();

                                        //            Flters.DataDocument = today;
                                        //            Flters.DataDocument2 = targetDate;
                                        //            Flters.clocktime30 = totalTimeDifference101;
                                        //            Flters.clocktime = totalTimeDifference504;
                                        //            Flters.Title = title;
                                        //            totalTimeDifference103 -= totalTimeDifference3;
                                        //            Flters.tarif = tar;
                                        //            Flters.FK_usr = fin.FK_usr;
                                        //            Flters.FK_Relationperson = creditline;
                                        //            Flters2.Add(Flters);
                                        //        }
                                        //    }

                                        //}

                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {

                                            dbRelatinafrad Flters = new dbRelatinafrad();

                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference3;
                                            Flters.Title = title;
                                            totalTimeDifference103 -= totalTimeDifference3;
                                            Flters.tarif = tar;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters2.Add(Flters);

                                        }


                                    }
                                }

                            }
                            else
                            {
                                totalTimeDifference104 -= totalTimeDifference10;


                            }

                            // اضافه کردن یک روز به تاریخ امروز
                        }
                        today = today.AddDays(1);

                    }
                }

                //if (totalTimeDifference10 > TimeSpan.Zero)
                //{
                //    return (-totalTimeDifference10).ToString();

                //}
                db.dbRelatinafrad.AddRange(Flters2);

                db.SaveChanges();

                if (totalTimeDifference10 > TimeSpan.Zero)
                {
                    var fin2000 = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                    var fi2n = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                    var finnnn2 = db.dbRelatinafrad.Where(p => p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.FK_usr == fi2n.FK_usr && p.tarif == tar && p.delet != true).OrderByDescending(s => s.ID).FirstOrDefault();
                    var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.delet != true && p.tarif == tar).OrderByDescending(s => s.ID).FirstOrDefault();
                    if (finnnn != null)
                    {
                        finnnn.clocktime = finnnn.clocktime += totalTimeDifference10;
                        db.SaveChanges();
                        //dbRelatinafrad Flters = new dbRelatinafrad();
                        //Flters.DataDocument = today;
                        //Flters.DataDocument2 = targetDate;
                        //Flters.clocktime30 = totalTimeDifference101;
                        //Flters.clocktime = totalTimeDifference103;
                        //Flters.Title = title;
                        //Flters.FK_usr = fin2000.FK_usr;
                        //Flters.FK_Relationperson = creditline;

                        //Flters.tarif = tar;
                        //Flters2.Add(Flters);
                    }
                    else
                    {
                        dbRelatinafrad Flters = new dbRelatinafrad();

                        Flters.DataDocument = finnnn2.DataDocument;
                        Flters.DataDocument2 = targetDate;
                        Flters.clocktime30 = totalTimeDifference101;
                        Flters.clocktime = totalTimeDifference10;
                        Flters.Title = title;
                        Flters.tarif = tar;
                        Flters.FK_usr = fi2n.FK_usr;
                        Flters.FK_Relationperson = creditline;
                        Flters2.Add(Flters);
                        db.dbRelatinafrad.Add(Flters);
                        db.SaveChanges();

                    }




                    return (-totalTimeDifference10).ToString();

                }

                return "0";

                // نمایش و ارسال تاریخ قمری به دیتابیس
                //string hijriDate = $"{hijriYear}/{hijriMonth}/{hijriDay}";
                //Console.WriteLine($"تاریخ قمری: {hijriDate}");

                // ارسال به دیتابیس - SendToDatabase(hijriDate);
            }

            else
            {
                return "1";
            }









        }
        public string Testpardas2(DateTime? DataDocument = null, int creditline = 0, string EndTime2 = "", string title = "", int tar = 0, bool IsVisible = false, bool IsVisible2 = false)
        {
            TimeSpan? totalTimeDifference101 = TimeSpan.Zero;
            TimeSpan? totalTimeDifference104 = TimeSpan.Zero;
            List<dbRelatinafrad> Flters2 = new List<dbRelatinafrad>();

            TimeSpan? totalTimeDifference10 = TimeSpan.Zero;
            TimeSpan? totalTimeDifference103 = TimeSpan.Zero;
            totalTimeDifference103 = TimeSpan.Parse(EndTime2);
            totalTimeDifference10 = TimeSpan.Parse(EndTime2);
            totalTimeDifference101 = TimeSpan.Parse(EndTime2);
            if (IsVisible == true)
            {
                TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                TimeSpan? totalTimeDifference54 = TimeSpan.Zero;

                totalTimeDifference2 = TimeSpan.Parse(EndTime2);

                // تبدیل تاریخ شمسی به میلادی
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime gregorianDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day);
                var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                int dayy;
                switch (dayOfWeek)
                {
                    case DayOfWeek.Saturday:
                        dayy = 1; // شنبه
                        break;
                    case DayOfWeek.Sunday:
                        dayy = 2; // یکشنبه
                        break;
                    case DayOfWeek.Monday:
                        dayy = 3; // دوشنبه
                        break;
                    case DayOfWeek.Tuesday:
                        dayy = 4; // سه‌شنبه
                        break;
                    case DayOfWeek.Wednesday:
                        dayy = 5; // چهارشنبه
                        break;
                    case DayOfWeek.Thursday:
                        dayy = 6; // پنج‌شنبه
                        break;
                    case DayOfWeek.Friday:
                        dayy = 7; // جمعه
                        break;
                    default:
                        dayy = 0; // خطا یا مقدار نامعتبر
                        break;
                }
                if (db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault() != null)
                {
                    return "3";

                }
                var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                foreach (var t in dayyu)
                {
                    totalTimeDifference3 += (t.end_clock - t.start_clock);

                }
                var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                if (moraghasi != null)
                {
                    if (moraghasi.end_clock != null)
                    {
                        totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);

                    }
                    if (moraghasi.end_clock2 != null)
                    {
                        totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);

                    }
                    if (totalTimeDifference3 <= TimeSpan.Zero)
                    {
                        return "4";

                    }

                }

                //// استفاده از HijriCalendar برای تبدیل تاریخ میلادی به قمری
                //HijriCalendar hijriCalendar = new HijriCalendar();
                //int hijriYear = hijriCalendar.GetYear(gregorianDate);
                //int hijriMonth = hijriCalendar.GetMonth(gregorianDate);
                //int hijriDay = hijriCalendar.GetDayOfMonth(gregorianDate);
                var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.DataDocument == gregorianDate && p.delet != true).ToList();
                foreach (var fiiii in find)
                {
                    totalTimeDifference1 += fiiii.clocktime;
                }
                totalTimeDifference3 -= totalTimeDifference1;
                totalTimeDifference3 -= totalTimeDifference2;
                if (totalTimeDifference3 <= TimeSpan.Zero)
                {
                    return totalTimeDifference3.ToString();


                }

                else
                {
                    return "0";

                }

                // نمایش و ارسال تاریخ قمری به دیتابیس
                //string hijriDate = $"{hijriYear}/{hijriMonth}/{hijriDay}";
                //Console.WriteLine($"تاریخ قمری: {hijriDate}");

                // ارسال به دیتابیس - SendToDatabase(hijriDate);
            }
            else if (IsVisible == false)
            {

                PersianCalendar persianCalendar = new PersianCalendar();

                DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day);
                DateTime today2 = DateTime.Now.Date;
                TimeSpan currentTime = DateTime.Now.TimeOfDay;
                currentTime = currentTime.Add(new TimeSpan(0, 15, 0)); // اضافه کردن 15 دقیقه

                DateTime today = DateTime.Now.Date;
                while (today <= targetDate)
                {

                    if (today == today2)
                    {
                        bool vaz = true;
                        TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference54 = TimeSpan.Zero;
                        totalTimeDifference2 = TimeSpan.Parse(EndTime2);

                        // تبدیل تاریخ شمسی به میلادی
                        PersianCalendar persianCalendar2 = new PersianCalendar();
                        DateTime gregorianDate = today;
                        var findday = db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault();
                        if (findday != null)
                        {
                            //TimeSpan? totalTimeDifference300 = TimeSpan.Zero;

                            //var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                            //DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            //int dayy;
                            //switch (dayOfWeek)
                            //{
                            //    case DayOfWeek.Saturday:
                            //        dayy = 1; // شنبه
                            //        break;
                            //    case DayOfWeek.Sunday:
                            //        dayy = 2; // یکشنبه
                            //        break;
                            //    case DayOfWeek.Monday:
                            //        dayy = 3; // دوشنبه
                            //        break;
                            //    case DayOfWeek.Tuesday:
                            //        dayy = 4; // سه‌شنبه
                            //        break;
                            //    case DayOfWeek.Wednesday:
                            //        dayy = 5; // چهارشنبه
                            //        break;
                            //    case DayOfWeek.Thursday:
                            //        dayy = 6; // پنج‌شنبه
                            //        break;
                            //    case DayOfWeek.Friday:
                            //        dayy = 7; // جمعه
                            //        break;
                            //    default:
                            //        dayy = 0; // خطا یا مقدار نامعتبر
                            //        break;
                            //}
                            //var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            //foreach (var t in dayyu)
                            //{
                            //    totalTimeDifference300 += (t.end_clock - t.start_clock);
                            //}
                            //if(totalTimeDifference300> TimeSpan.Zero)
                            //{
                            //    vaz = true;
                            //}
                            //else
                            //{
                            //    vaz = false;
                            //}

                            if (IsVisible2 == true)
                            {
                                vaz = true;
                            }
                            else
                            {
                                vaz = false;

                            }
                        }
                        if (vaz == true)
                        {
                            var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                            DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                            int dayy;
                            switch (dayOfWeek)
                            {
                                case DayOfWeek.Saturday:
                                    dayy = 1; // شنبه
                                    break;
                                case DayOfWeek.Sunday:
                                    dayy = 2; // یکشنبه
                                    break;
                                case DayOfWeek.Monday:
                                    dayy = 3; // دوشنبه
                                    break;
                                case DayOfWeek.Tuesday:
                                    dayy = 4; // سه‌شنبه
                                    break;
                                case DayOfWeek.Wednesday:
                                    dayy = 5; // چهارشنبه
                                    break;
                                case DayOfWeek.Thursday:
                                    dayy = 6; // پنج‌شنبه
                                    break;
                                case DayOfWeek.Friday:
                                    dayy = 7; // جمعه
                                    break;
                                default:
                                    dayy = 0; // خطا یا مقدار نامعتبر
                                    break;
                            }



                            var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            foreach (var t in dayyu)
                            {
                                if (t.end_clock > currentTime)
                                {
                                    if (t.start_clock > currentTime)
                                    {
                                        totalTimeDifference3 += (t.end_clock - t.start_clock);

                                    }
                                    else
                                    {
                                        totalTimeDifference3 += (t.end_clock - currentTime);

                                    }
                                }
                                //totalTimeDifference3 += (t.end_clock - t.start_clock);
                            }

                            var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                            if (moraghasi != null)
                            {
                                if (moraghasi.end_clock != null)
                                {

                                    if (moraghasi.end_clock > currentTime)
                                    {
                                        if (moraghasi.start_clock > currentTime)
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);

                                        }
                                        else
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock - currentTime);

                                        }
                                    }
                                }
                                if (moraghasi.end_clock2 != null)
                                {

                                    if (moraghasi.end_clock2 > currentTime)
                                    {
                                        if (moraghasi.start_clock2 > currentTime)
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);

                                        }
                                        else
                                        {
                                            totalTimeDifference3 -= (moraghasi.end_clock2 - currentTime);

                                        }
                                    }
                                }
                                //if (moraghasi.end_clock != null)
                                //{
                                //    totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);
                                //}
                                //if (moraghasi.end_clock2 != null)
                                //{
                                //    totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);
                                //}

                            }

                            var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.DataDocument == gregorianDate && p.FK_tbcreatfaal != null && p.delet != true).ToList();
                            foreach (var fiiii in find)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }
                            var find22 = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.Title != title && p.DataDocument == gregorianDate && p.FK_tbcreatfaal == null && p.tarif == tar && p.delet != true).ToList();
                            foreach (var fiiii in find22)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }

                            totalTimeDifference3 -= totalTimeDifference1;

                            if (totalTimeDifference3 > TimeSpan.Zero)
                            {
                                if (totalTimeDifference10 > TimeSpan.Zero)
                                {

                                    totalTimeDifference10 -= totalTimeDifference3;

                                    if (totalTimeDifference10 <= TimeSpan.Zero)
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {
                                            dbRelatinafrad Flters = new dbRelatinafrad();
                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference103;
                                            Flters.Title = title;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters.delet = false;
                                            Flters.tarif = tar;
                                            Flters2.Add(Flters);
                                        }

                                        //return "0";
                                    }
                                    else
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {

                                            dbRelatinafrad Flters = new dbRelatinafrad();

                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference3;
                                            Flters.Title = title;
                                            totalTimeDifference103 -= totalTimeDifference3;
                                            Flters.tarif = tar;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.delet = false;

                                            Flters.FK_Relationperson = creditline;
                                            Flters2.Add(Flters);
                                        }

                                    }
                                }

                            }
                            else
                            {
                                totalTimeDifference104 -= totalTimeDifference10;


                            }

                            // اضافه کردن یک روز به تاریخ امروز
                        }
                        today = today.AddDays(1);
                    }
                    else
                    {
                        bool vaz = true;

                        TimeSpan? totalTimeDifference1 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference2 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference3 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference4 = TimeSpan.Zero;
                        TimeSpan? totalTimeDifference54 = TimeSpan.Zero;
                        totalTimeDifference2 = TimeSpan.Parse(EndTime2);
                        DateTime gregorianDate = today;
                        var findday = db.tbcraettatel.Where(p => p.Dtatime == gregorianDate).FirstOrDefault();
                        if (findday != null)
                        {
                            //TimeSpan? totalTimeDifference300 = TimeSpan.Zero;

                            //var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                            //DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            //int dayy;
                            //switch (dayOfWeek)
                            //{
                            //    case DayOfWeek.Saturday:
                            //        dayy = 1; // شنبه
                            //        break;
                            //    case DayOfWeek.Sunday:
                            //        dayy = 2; // یکشنبه
                            //        break;
                            //    case DayOfWeek.Monday:
                            //        dayy = 3; // دوشنبه
                            //        break;
                            //    case DayOfWeek.Tuesday:
                            //        dayy = 4; // سه‌شنبه
                            //        break;
                            //    case DayOfWeek.Wednesday:
                            //        dayy = 5; // چهارشنبه
                            //        break;
                            //    case DayOfWeek.Thursday:
                            //        dayy = 6; // پنج‌شنبه
                            //        break;
                            //    case DayOfWeek.Friday:
                            //        dayy = 7; // جمعه
                            //        break;
                            //    default:
                            //        dayy = 0; // خطا یا مقدار نامعتبر
                            //        break;
                            //}
                            //var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            //foreach (var t in dayyu)
                            //{
                            //    totalTimeDifference300 += (t.end_clock - t.start_clock);
                            //}
                            //if (totalTimeDifference300 > TimeSpan.Zero)
                            //{
                            //    vaz = true;
                            //}
                            //else
                            //{
                            //    vaz = false;
                            //}

                            if (IsVisible2 == true)
                            {
                                vaz = true;
                            }
                            else
                            {
                                vaz = false;

                            }
                        }
                        // تبدیل تاریخ شمسی به میلادی
                        PersianCalendar persianCalendar2 = new PersianCalendar();
                        if (vaz == true)
                        {
                            var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
                            DayOfWeek dayOfWeek = gregorianDate.DayOfWeek;

                            // تبدیل روز هفته میلادی به روز هفته شمسی (شنبه = 1، یکشنبه = 2، ...)
                            int dayy;
                            switch (dayOfWeek)
                            {
                                case DayOfWeek.Saturday:
                                    dayy = 1; // شنبه
                                    break;
                                case DayOfWeek.Sunday:
                                    dayy = 2; // یکشنبه
                                    break;
                                case DayOfWeek.Monday:
                                    dayy = 3; // دوشنبه
                                    break;
                                case DayOfWeek.Tuesday:
                                    dayy = 4; // سه‌شنبه
                                    break;
                                case DayOfWeek.Wednesday:
                                    dayy = 5; // چهارشنبه
                                    break;
                                case DayOfWeek.Thursday:
                                    dayy = 6; // پنج‌شنبه
                                    break;
                                case DayOfWeek.Friday:
                                    dayy = 7; // جمعه
                                    break;
                                default:
                                    dayy = 0; // خطا یا مقدار نامعتبر
                                    break;
                            }



                            var dayyu = db.dbtarifmovazaf.Where(p => p.FK_usr == fin.FK_usr && p.day == dayy).ToList();
                            foreach (var t in dayyu)
                            {
                                totalTimeDifference3 += (t.end_clock - t.start_clock);
                            }

                            var moraghasi = db.dbvaziathozor.Where(p => p.FK_usr == fin.FK_usr && p.datattime == gregorianDate).FirstOrDefault();
                            if (moraghasi != null)
                            {
                                if (moraghasi.end_clock != null)
                                {
                                    totalTimeDifference3 -= (moraghasi.end_clock - moraghasi.start_clock);
                                }
                                if (moraghasi.end_clock2 != null)
                                {
                                    totalTimeDifference3 -= (moraghasi.end_clock2 - moraghasi.start_clock2);
                                }

                            }

                            var find = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.DataDocument == gregorianDate && p.FK_tbcreatfaal != null && p.delet != true).ToList();
                            foreach (var fiiii in find)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }
                            var find22 = db.dbRelatinafrad.Where(p => p.FK_usr == fin.FK_usr && p.Title != title && p.DataDocument == gregorianDate && p.FK_tbcreatfaal == null && p.tarif == tar && p.delet != true).ToList();
                            foreach (var fiiii in find22)
                            {
                                totalTimeDifference1 += fiiii.clocktime;
                            }

                            totalTimeDifference3 -= totalTimeDifference1;

                            if (totalTimeDifference3 > TimeSpan.Zero)
                            {
                                if (totalTimeDifference10 > TimeSpan.Zero)
                                {

                                    totalTimeDifference10 -= totalTimeDifference3;

                                    if (totalTimeDifference10 <= TimeSpan.Zero)
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {
                                            dbRelatinafrad Flters = new dbRelatinafrad();
                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference103;
                                            Flters.Title = title;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters.delet = false;

                                            Flters.tarif = tar;
                                            Flters2.Add(Flters);
                                        }

                                        //return "0";
                                    }
                                    else
                                    {
                                        var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).FirstOrDefault();
                                        if (finnnn == null)
                                        {

                                            dbRelatinafrad Flters = new dbRelatinafrad();

                                            Flters.DataDocument = today;
                                            Flters.DataDocument2 = targetDate;
                                            Flters.clocktime30 = totalTimeDifference101;
                                            Flters.clocktime = totalTimeDifference3;
                                            Flters.Title = title;
                                            totalTimeDifference103 -= totalTimeDifference3;
                                            Flters.delet = false;

                                            Flters.tarif = tar;
                                            Flters.FK_usr = fin.FK_usr;
                                            Flters.FK_Relationperson = creditline;
                                            Flters2.Add(Flters);
                                        }

                                    }
                                }

                            }
                            else
                            {
                                totalTimeDifference104 -= totalTimeDifference10;


                            }

                            // اضافه کردن یک روز به تاریخ امروز
                        }
                        today = today.AddDays(1);
                    }
                }
                db.dbRelatinafrad.AddRange(Flters2);

                db.SaveChanges();
                if (totalTimeDifference10 > TimeSpan.Zero)
                {
                    var fin2000 = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();

                    var finnnn = db.dbRelatinafrad.Where(p => p.DataDocument == today && p.DataDocument2 == targetDate && p.clocktime30 == totalTimeDifference101 && p.Title == title && p.FK_Relationperson == creditline && p.FK_tbcreatfaal == null && p.tarif == tar).OrderByDescending(s => s.ID).FirstOrDefault();
                    if (finnnn != null)
                    {
                        finnnn.clocktime = finnnn.clocktime += totalTimeDifference10;
                        db.SaveChanges();
                        //dbRelatinafrad Flters = new dbRelatinafrad();
                        //Flters.DataDocument = today;
                        //Flters.DataDocument2 = targetDate;
                        //Flters.clocktime30 = totalTimeDifference101;
                        //Flters.clocktime = totalTimeDifference103;
                        //Flters.Title = title;
                        //Flters.FK_usr = fin2000.FK_usr;
                        //Flters.FK_Relationperson = creditline;

                        //Flters.tarif = tar;
                        //Flters2.Add(Flters);
                    }

                    return (-totalTimeDifference10).ToString();

                }
                //else if (totalTimeDifference10 <= TimeSpan.Zero)
                //{
                //    return (totalTimeDifference10).ToString();

                //}


                return "0";

                // نمایش و ارسال تاریخ قمری به دیتابیس
                //string hijriDate = $"{hijriYear}/{hijriMonth}/{hijriDay}";
                //Console.WriteLine($"تاریخ قمری: {hijriDate}");

                // ارسال به دیتابیس - SendToDatabase(hijriDate);
            }

            else
            {
                return "1";
            }









        }
        public string crateerjaja(DateTime? DataDocument = null, int creditline = 0, string EndTime2 = "", string title = "", int tar = 0, bool IsVisible = false, int delet = 0, string canornotcan = "")
        {
            TimeSpan? totalTimeDifference101 = TimeSpan.Zero;
            TimeSpan? totalTimeDifference104 = TimeSpan.Zero;
            List<dbRelatinafrad> Flters2 = new List<dbRelatinafrad>();
            dbRelatinafrad Flters20 = new dbRelatinafrad();
            var fin = db.dbrelatinperson.Where(p => p.ID == creditline).FirstOrDefault();
            var fin2 = db.dbRelatinafrad.Where(p => p.ID == delet).FirstOrDefault();
            int del = 0;
            if (fin2 != null)
            {
                del = (int)fin2.FK_tbcreatfaal;
            }
            Flters20.FK_usr = fin.FK_usr;

            Flters20.FK_Relationperson = creditline;
            Flters20.DataDocument = DataDocument;
            Flters20.canornotcan = canornotcan;
            Flters20.FK_tbcreatfaal = del;
            Flters20.FK_relatinafrad = fin2.ID;
            Flters20.clocktime = TimeSpan.Parse(EndTime2);
            Flters20.Title = title;
            Flters20.tarif = null;
            db.dbRelatinafrad.Add(Flters20);
            fin2.delet = true;
            db.SaveChanges();

            return "True";







        }

        // Helper function to convert Gregorian DayOfWeek to Persian Day numbers
        private int GetPersianDayOfWeek(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Saturday: return 1;
                case DayOfWeek.Sunday: return 2;
                case DayOfWeek.Monday: return 3;
                case DayOfWeek.Tuesday: return 4;
                case DayOfWeek.Wednesday: return 5;
                case DayOfWeek.Thursday: return 6;
                case DayOfWeek.Friday: return 7;
                default: return 0; // Invalid day
            }
        }

        public ActionResult _EditFinnail(int id)
        {
            var result = rep_def_rsn.Find(id);
            if (result.User_ID == null)
            {
                result.User_ID = 0;
            }
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/Edit_Finacial.cshtml", result);
        }
        public ActionResult _EditFinnail2(int id)
        {
            var result = rep_def_rsn.Find(id);
            if (result.User_ID == null)
            {
                result.User_ID = 0;
            }
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/Edit_Finacial.cshtml", result);
        }
        public ActionResult viewasnadfishmamor(int year = 0, int month = 1, int id = 0)
        {
            var t = db.FinancialDocuments.Where(p => p.User_ID == id).ToList();
            List<FinancialDocuments> result = new List<FinancialDocuments>();
            foreach (var item1 in t)
            {
                if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                    int persianYear = persianCalendar.GetYear(dataDocument);
                    int persianMonth = persianCalendar.GetMonth(dataDocument);

                    if (persianYear == year && persianMonth == month)
                    {
                        result.Add(item1);
                    }

                }
            }

            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/viewasnadfishmamor.cshtml", result);
        }
        public ActionResult Detail(int id)
        {
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/Detail.cshtml", rep_def_rsn.Find(id));
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            string macAddress = GetMacAddress();
            Response.Write("MAC Address: " + macAddress);
        }

        public static string GetMacAddress()
        {
            string macAddresses = string.Empty;
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    macAddresses += nic.GetPhysicalAddress().ToString();
                    break;
                }
            }
            return macAddresses;
        }
        public string creat(List<FinancialDocuments> Filters, IEnumerable<HttpPostedFileBase> files)
        {
            List<FinancialDocuments> Flters = new List<FinancialDocuments>();
            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        foreach (var it in Filters)
                        {
                            it.File_SystemName = filename;
                            it.File_RealName = file.FileName;

                            Flters.Add(it);
                        }
                    }
                    else
                    {

                    }
                }

            }

            foreach (var it in Filters)
            {
                var tt = db.tbfkfinancial.Where(p => p.ID == it.FK_final).FirstOrDefault();
                it.MacAdress = GetMacAddress();
                if (it.Peyman_ID == 0)
                {
                    it.Peyman_ID = null; // Change to nullable int
                }
                it.DataDocument = tt.DataDocument;
                Flters.Add(it);
            }

            return rep_def_rsn.Create(Flters);
        }



        public ActionResult crattattaril(DateTime? DataDocument = null, int creditline = 0, string titr = "")
        {
            tbcraettatel tbcraettatel = new tbcraettatel();
            tbcraettatel.Dtatime = DataDocument;
            tbcraettatel.day = titr;
            tbcraettatel.Title = creditline;
            db.tbcraettatel.Add(tbcraettatel);
            db.SaveChanges();
            return Content("True");
        }
        public ActionResult crattattariledit(DateTime? DataDocument = null, int creditline = 0, string titr = "", int ID = 0)

        {
            var find = db.tbcraettatel.Where(p => p.ID == ID).FirstOrDefault();
            if (find != null)
            {
                find.Dtatime = DataDocument;
                find.day = titr;
                find.Title = creditline;
                db.SaveChanges();

            }
            else
            {
                tbcraettatel tbcraettatel = new tbcraettatel();
                tbcraettatel.Dtatime = DataDocument;
                tbcraettatel.day = titr;
                tbcraettatel.Title = creditline;
                db.tbcraettatel.Add(tbcraettatel);
                db.SaveChanges();
            }

            return Content("True");
        }
        public ActionResult viewuplodsanad()
        {
            return View();
        }
        public ActionResult viewupload2(int id = 0)
        {
            var find = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == id).FirstOrDefault();

            return View(find);
        }
        public ActionResult VIEWLIST()
        {
            var fin = db.tbcraettatel.ToList();
            return View(fin);
        }
        public string creatfktitl(List<dbRelatinafrad> Filters, IEnumerable<HttpPostedFileBase> files)
        {
            List<dbRelatinafrad> Flters = new List<dbRelatinafrad>();
            //if (files != null)
            //{
            //    foreach (var file in files)
            //    {

            //        if (file.ContentLength > 0)
            //        {
            //            var segment = file.FileName.Split('.');
            //            string file_type = segment[segment.Length - 1];
            //            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
            //            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
            //            foreach (var it in Filters)
            //            {
            //                it.File_SystemName = filename;
            //                it.File_RealName = file.FileName;

            //                Flters.Add(it);
            //            }
            //        }
            //        else
            //        {

            //        }
            //    }

            //}




            PersianCalendar persianCalendar = new PersianCalendar();


            foreach (var it in Filters)
            {
                var fin = db.dbrelatinperson.Where(p => p.ID == it.FK_Relationperson).FirstOrDefault();

                DateTime gregorianDate2 = new DateTime(it.DataDocument.Value.Year, it.DataDocument.Value.Month, it.DataDocument.Value.Day, persianCalendar);
                TimeSpan? totalTimeDifference2 = TimeSpan.Parse(it.clocktime2);

                var findd23 = db.dbRelatinafrad.Where(p => p.FK_Relationperson == it.FK_Relationperson && p.Title == it.Title && p.DataDocument2 == gregorianDate2 && p.clocktime30 == totalTimeDifference2 && p.FK_tbcreatfaal == null && p.tarif == it.tarif).ToList();
                if (findd23.Count != 0)
                {
                    foreach (var findd2 in findd23)
                    {

                        var findd = db.tbperson.Where(p => p.ID == it.FK_Relationperson).FirstOrDefault();

                        if (findd != null)
                        {
                            findd2.FK_person = findd.ID;

                        }
                        else
                        {
                            findd2.FK_person = null;

                        }
                        findd2.tarif = null;
                        findd2.FK_usr = fin.FK_usr;
                        findd2.FK_Relationperson = it.FK_Relationperson;
                        findd2.FK_tbcreatfaal = it.FK_tbcreatfaal;
                        findd2.canornotcan = it.canornotcan;
                        db.SaveChanges();


                    }
                }


                else
                {

                    var tt = db.tbfkfinancial.Where(p => p.ID == it.FK_tbcreatfaal).FirstOrDefault();
                    var findd = db.tbperson.Where(p => p.ID == it.FK_Relationperson).FirstOrDefault();
                    if (findd != null)
                    {
                        it.FK_person = findd.ID;

                    }
                    else
                    {
                        it.FK_person = null;

                    }
                    it.tarif = null;
                    it.FK_usr = fin.FK_usr;
                    it.FK_Relationperson = it.FK_Relationperson;
                    DateTime gregorianDate = new DateTime(it.DataDocument.Value.Year, it.DataDocument.Value.Month, it.DataDocument.Value.Day, persianCalendar);
                    it.clocktime = TimeSpan.Parse(it.clocktime2);
                    it.DataDocument = gregorianDate;
                    it.clocktime2 = it.clocktime2;
                    it.canornotcan = it.canornotcan;
                    it.Title = it.Title;
                    it.delet = false;
                    Flters.Add(it);
                }

            }
            db.dbRelatinafrad.AddRange(Flters);
            db.SaveChanges();
            return "True";
        }

        public ActionResult viewasnadd(List<DataModel> dataList)
        {
            if (dataList != null && dataList.Any())
            {
                foreach (var data in dataList)
                {
                    var bindValue = data.Bind;
                    var containerValue = data.Container;
                    var name = data.Name;
                    var dbAsnadcratname = db.dbAsnadcratname.Where(s => s.codmahsol == bindValue).FirstOrDefault();
                    dbAsnadcratname dbAsnadcratname1 = new dbAsnadcratname();
                    if (dbAsnadcratname != null)
                    { var find = db.dbAsnadcratname.Where(s => s.codmahsol == containerValue).FirstOrDefault();
                        if (find != null)
                        {

                            find.name = name;
                            db.SaveChanges();
                        }
                        else
                        {
                            dbAsnadcratname1.nergh = null;
                            dbAsnadcratname1.codmahsol = containerValue;
                            dbAsnadcratname1.IDre = dbAsnadcratname.ID;
                            dbAsnadcratname1.name = name;
                            db.dbAsnadcratname.Add(dbAsnadcratname1);
                            db.SaveChanges();
                        }


                    }
                    else
                    {
                        var find = db.dbAsnadcratname.Where(s => s.codmahsol == containerValue).FirstOrDefault();
                        if (find != null)
                        {

                            find.name = name;
                            db.SaveChanges();
                        }
                        else
                        {
                            dbAsnadcratname1.nergh = null;
                            dbAsnadcratname1.codmahsol = containerValue;
                            dbAsnadcratname1.IDre = null;
                            dbAsnadcratname1.name = name;
                            db.dbAsnadcratname.Add(dbAsnadcratname1);
                            db.SaveChanges();
                        }


                    }
                    // پردازش داده‌های دریافتی


                }
                return Json("true"); // پاسخ موفقیت
            }
            return Json("false"); // پاسخ خطا
        }
        public ActionResult viewasnadduplad(List<DataModel> dataList)
        {
            if (dataList != null && dataList.Any())
            {
                foreach (var data in dataList)
                {
                    var bindValue = data.Bind;
                    var containerValue = data.Container;
                    var name = data.Name;
                    var dbAsnadcratname = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == bindValue).FirstOrDefault();
                    dbAsnaduplaodfil1 dbAsnadcratname1 = new dbAsnaduplaodfil1();
                    if (dbAsnadcratname != null)
                    {
                        var find = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == containerValue).FirstOrDefault();
                        if (find != null)
                        {

                            find.name = name;
                            db.SaveChanges();
                        }
                        else
                        {
                            dbAsnadcratname1.nergh = null;
                            dbAsnadcratname1.codmahsol = containerValue;
                            dbAsnadcratname1.IDre = dbAsnadcratname.ID;
                            dbAsnadcratname1.name = name;
                            db.dbAsnaduplaodfil1.Add(dbAsnadcratname1);
                            db.SaveChanges();
                        }


                    }
                    else
                    {
                        var find = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == containerValue).FirstOrDefault();
                        if (find != null)
                        {

                            find.name = name;
                            db.SaveChanges();
                        }
                        else
                        {
                            dbAsnadcratname1.nergh = null;
                            dbAsnadcratname1.codmahsol = containerValue;
                            dbAsnadcratname1.IDre = null;
                            dbAsnadcratname1.name = name;
                            db.dbAsnaduplaodfil1.Add(dbAsnadcratname1);
                            db.SaveChanges();
                        }


                    }
                    // پردازش داده‌های دریافتی


                }
                return Json("true"); // پاسخ موفقیت
            }
            return Json("false"); // پاسخ خطا
        }

        public class DataModel
        {
            public string Bind { get; set; }
            public string Container { get; set; }
            public string Name { get; set; }
        }

        public ActionResult viewasnadmaly454yupladtreatabel3(string name = "")
        {
            var findResult = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == name).FirstOrDefault();
            string nam = "";
            int idd = 0;

            if (findResult != null)
            {
                nam = findResult.name;
                var findsabt = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == findResult.ID).FirstOrDefault();
                if (findsabt != null)
                {
                    idd = 1;



                }
            }

            return Json(new { nam = nam, idd = idd });
        }


        public ActionResult viewasnadmaly454yuplad(string name = "")
        {
            var findResult = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == name).FirstOrDefault();
            string nam = "";
            int idd = 0;

            if (findResult != null)
            {
                nam = findResult.name;
                var findsabt = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == findResult.ID).FirstOrDefault();
                if (findsabt != null)
                {
                    var findsabtupllad = db.dbAsnaduplaodfil3.Where(s => s.FK_upload2 == findsabt.ID).FirstOrDefault();
                    if (findsabtupllad != null)
                    {
                        idd = 1;
                    }

                }
            }

            return Json(new { nam = nam, idd = idd });
        }
        public string viewasnadmaly454y(string name = "")
        {
            var findResult = db.dbAsnadcratname.Where(s => s.codmahsol == name).FirstOrDefault();
            string nam = "";
            if (findResult != null)
            {
                nam = findResult.name;

            }

            return nam;

        }
        public int viewasnadmaly454y22uplad(string name = "")
        {
            var findResult = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == name).FirstOrDefault();
            int count = 0;
            if (findResult != null)
            {
                var countt = db.dbAsnaduplaodfil1.Where(s => s.IDre == findResult.ID).ToList();
                count = countt.Count();
            }

            return count;

        }
        public int viewasnadmaly454y22(string name = "")
        {
            var findResult = db.dbAsnadcratname.Where(s => s.codmahsol == name).FirstOrDefault();
            int count = 0;
            if (findResult != null)
            {
                var countt = db.dbAsnadcratname.Where(s => s.IDre == findResult.ID).ToList();
                count = countt.Count();
            }

            return count;

        }

        public int viewasnadmalkol(string name = "")
        {
            var findResult = db.dbAsnadcratname.Where(s => s.codmahsol == name).FirstOrDefault();
            int count = 0;
            if (findResult != null)
            {
                var countt = db.dbAsnadcratname.Where(s => s.IDre == findResult.ID).ToList();
                count = countt.Count();
            }

            return count;

        }
        public ActionResult detailkolshomareh(string name = "")
        {
            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            var findResult = db.dbAsnadcratname.Where(s => s.codmahsol == name).FirstOrDefault();
            int count = 0;
            var currentItem = findResult; // شروع از آیتم فعلی

            bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر

            // ادامه جستجو تا زمانی که IDre برابر null شود



            var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
            foreach (var item in list11)
            {
                var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                foreach (var item1 in listt2)
                {
                    var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                    foreach (var item2 in list3)
                    {
                        var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                        foreach (var item20 in list4)
                        {
                            var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                            foreach (var item201 in list40)
                            {
                                var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                foreach (var item2011 in list401)
                                {
                                    var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                    var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                    dbAsnafsherkatsanad.AddRange(list5011);
                                }
                                var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list501);
                            }
                            var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                            dbAsnafsherkatsanad.AddRange(list50);
                        }

                        var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                        dbAsnafsherkatsanad.AddRange(list34);
                    }
                    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list);
                }
                var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                dbAsnafsherkatsanad.AddRange(list5);
            }




            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید


            if (isFirst)
            {
                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                dbAsnafsherkatsanad.AddRange(list);
                isFirst = false;
            }
            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(dbAsnafsherkatsanad);

        }


        public ActionResult gozareshmaly1()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود


            foreach(var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر

    
                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm = 0;
                    double summmbedh = 0;
                    double sumkol = 0;
                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                        dbAsnafsherkatsanad.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                    dbAsnafsherkatsanad.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                            dbAsnafsherkatsanad.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                        dbAsnafsherkatsanad.AddRange(list);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {
                        summm += (double)it.bestankar;
                        summmbedh += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm - summmbedh);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh,
                        SUMBESTENKAR = summm,
                        coding = item.codmahsol,
                        namecoding = item.name
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{
                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}

            }
         




            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید


          
            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }

        public ActionResult gozareshmaly12sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود


            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر

         
                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm = 0;
                    double summmbedh = 0;
                    double sumkol = 0;
                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                        dbAsnafsherkatsanad.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                    dbAsnafsherkatsanad.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                            dbAsnafsherkatsanad.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                        dbAsnafsherkatsanad.AddRange(list);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {
                        summm += (double)it.bestankar;
                        summmbedh += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm - summmbedh);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh,
                        SUMBESTENKAR = summm,
                        coding = item.codmahsol,
                        namecoding = item.name
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);
                }
                //if (isFirst)
                //{
                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}
               
            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }

        public ActionResult gozareshmaly22sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود
            long FK_ID = 0;

            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر
   
          
                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    FK_ID += 1;

                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm4 = 0;
                    double summmbedh4 = 0;
                    double sumkol = 0;
                 
                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        double summm = 0;
                        double summmbedh = 0;
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();

                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                        dbAsnafsherkatsanad2.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                    dbAsnafsherkatsanad2.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                                dbAsnafsherkatsanad2.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                            dbAsnafsherkatsanad2.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                        dbAsnafsherkatsanad2.AddRange(list);
                        foreach (var it in dbAsnafsherkatsanad2)
                        {
                            summm += (double)it.bestankar;
                            summmbedh += (double)it.bedehkar;
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }
                        sumkol = (double)(summm - summmbedh);
                        var gozareshmally1 = new gozareshmally
                        {
                            SUM = sumkol,
                            SUMBEDEHKAR = summmbedh,
                            SUMBESTENKAR = summm,
                            coding = item1.codmahsol,
                            namecoding = item1.name,
                            FK_ID = FK_ID
                        };
                        lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {

                        summm4 += (double)it.bestankar;
                        summmbedh4 += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm4 - summmbedh4);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh4,
                        SUMBESTENKAR = summm4,
                        coding = currentItem.codmahsol,
                        namecoding = currentItem.name,
                        FK_ID = 0,
                        ID = FK_ID
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{
            
                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}
              
            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }
        public ActionResult gozareshmaly32sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود
            long FK_ID = 0;
            long FK_ID2 = 0;

            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر


                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    FK_ID += 1;

                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm4 = 0;
                    double summmbedh4 = 0;
                    double sumkol = 0;

                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        FK_ID += 2;
                        double summm = 0;
                        double summmbedh = 0;
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();

                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            double summmtaf = 0;
                            double summmbedhtaf = 0;
                            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad3 = new List<dbAsnafsherkatsanad>();

                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                        dbAsnafsherkatsanad3.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                    dbAsnafsherkatsanad3.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                                dbAsnafsherkatsanad3.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                            dbAsnafsherkatsanad3.AddRange(list34);
                            foreach (var it in dbAsnafsherkatsanad3)
                            {
                                summmtaf += (double)it.bestankar;
                                summmbedhtaf += (double)it.bedehkar;
                                summm += (double)it.bestankar;
                                summmbedh += (double)it.bedehkar;
                                summm4 += (double)it.bestankar;
                                summmbedh4 += (double)it.bedehkar;
                            }
                            sumkol = (double)(summmtaf - summmbedhtaf);
                            var gozareshmally2 = new gozareshmally
                            {
                                SUM = sumkol,
                                SUMBEDEHKAR = summmbedhtaf,
                                SUMBESTENKAR = summmtaf,
                                coding = item2.codmahsol,
                                namecoding = item2.name,
                                FK_ID = FK_ID2,
                               
                            };
                            lstMoalefeexcel2.gozareshmally.Add(gozareshmally2);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                        dbAsnafsherkatsanad2.AddRange(list);
                        foreach (var it in dbAsnafsherkatsanad2)
                        {
                            summm += (double)it.bestankar;
                            summmbedh += (double)it.bedehkar;
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }

                        sumkol = (double)(summm - summmbedh);
                        var gozareshmally1 = new gozareshmally
                        {
                            SUM = sumkol,
                            SUMBEDEHKAR = summmbedh,
                            SUMBESTENKAR = summm,
                            coding = item1.codmahsol,
                            namecoding = item1.name,
                            FK_ID = FK_ID,
                            ID = FK_ID2
                        };
                        lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {

                        summm4 += (double)it.bestankar;
                        summmbedh4 += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm4 - summmbedh4);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh4,
                        SUMBESTENKAR = summm4,
                        coding = currentItem.codmahsol,
                        namecoding = currentItem.name,
                        FK_ID = 0,
                        ID = FK_ID
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{

                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}

            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }
        public ActionResult gozareshmaly34sotoneh()
        {
            var model = new FunctionModel { gozareshmally = new List<gozareshmally>() };
            long idCounter = 1;

            var roots = db.dbAsnadcratname.Where(x => x.IDre == null).ToList();

            foreach (var root in roots)
            {
                AddHierarchyItems(root, 0, null, model.gozareshmally, ref idCounter);
            }

            return View(model);
        }


        private (double totalBestankar, double totalBedehkar) AddHierarchyItems(
       dbAsnadcratname current,
       int level,
       long? parentId,
       List<gozareshmally> list,
       ref long idCounter)
        {
            // بدهکار و بستانکار خود نود فعلی
            double ownBestankar = 0, ownBedehkar = 0;

            var subAsnaf = db.dbAsnafsherkatsanad.Where(x => x.FK_name == current.ID).ToList();
            foreach (var item in subAsnaf)
            {
                ownBestankar += item.bestankar ?? 0;
                ownBedehkar += item.bedehkar ?? 0;
            }

            // مقادیر تجمیعی (خود نود + فرزندان)
            double totalBestankar = ownBestankar;
            double totalBedehkar = ownBedehkar;

            var rowId = idCounter++;

            var children = db.dbAsnadcratname.Where(x => x.IDre == current.ID).ToList();
            foreach (var child in children)
            {
                var (childBestankar, childBedehkar) = AddHierarchyItems(child, level + 1, rowId, list, ref idCounter);

                totalBestankar += childBestankar;
                totalBedehkar += childBedehkar;
            }

            // جمع کل = شامل خودش + همه زیرشاخه‌ها
            var row = new gozareshmally
            {
                ID = rowId,
                FK_ID = parentId ?? 0,
                coding = current.codmahsol,
                namecoding = current.name,
                SUMBEDEHKAR = totalBedehkar,
                SUMBESTENKAR = totalBestankar,
                SUM = totalBestankar - totalBedehkar
            };

            // حالا هر لایه‌ای رو که خواستی می‌تونی نگه‌داری (مثلاً فقط ۲ تا ۵)
           
                list.Add(row);
            

            return (totalBestankar, totalBedehkar);
        }








        public IEnumerable<gozareshmally> GetGozareshmallyLazy()
            {
                var level0List = db.dbAsnadcratname.Where(s => s.IDre == null);
                long FK_ID = 0;
                long FK_ID2 = 0;

                foreach (var currentItem in level0List)
                {
                    var listLevel1 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID);
                    foreach (var item in listLevel1)
                    {
                        FK_ID++;

                        double summm4 = 0;
                        double summmbedh4 = 0;

                        var listLevel2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID);
                        foreach (var item1 in listLevel2)
                        {
                            FK_ID += 2;

                            double summm = 0;
                            double summmbedh = 0;

                            var listLevel3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID);
                            foreach (var item2 in listLevel3)
                            {
                                double summmtaf = 0;
                                double summmbedhtaf = 0;

                                var children = GetAllChildren(item2.ID);
                                foreach (var it in children)
                                {
                                    summmtaf += (double)it.bestankar;
                                    summmbedhtaf += (double)it.bedehkar;
                                }

                                yield return new gozareshmally
                                {
                                    SUM = summmtaf - summmbedhtaf,
                                    SUMBEDEHKAR = summmbedhtaf,
                                    SUMBESTENKAR = summmtaf,
                                    coding = item2.codmahsol,
                                    namecoding = item2.name,
                                    FK_ID = FK_ID2
                                };

                                summm += summmtaf;
                                summmbedh += summmbedhtaf;
                                summm4 += summmtaf;
                                summmbedh4 += summmbedhtaf;
                            }

                            var directItems2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID);
                            foreach (var it in directItems2)
                            {
                                summm += (double)it.bestankar;
                                summmbedh += (double)it.bedehkar;
                                summm4 += (double)it.bestankar;
                                summmbedh4 += (double)it.bedehkar;
                            }

                            yield return new gozareshmally
                            {
                                SUM = summm - summmbedh,
                                SUMBEDEHKAR = summmbedh,
                                SUMBESTENKAR = summm,
                                coding = item1.codmahsol,
                                namecoding = item1.name,
                                FK_ID = FK_ID,
                                ID = FK_ID2
                            };
                        }

                        var directItems1 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID);
                        foreach (var it in directItems1)
                        {
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }

                        yield return new gozareshmally
                        {
                            SUM = summm4 - summmbedh4,
                            SUMBEDEHKAR = summmbedh4,
                            SUMBESTENKAR = summm4,
                            coding = currentItem.codmahsol,
                            namecoding = currentItem.name,
                            FK_ID = 0,
                            ID = FK_ID
                        };
                    }
                }
            }

            private IEnumerable<dbAsnafsherkatsanad> GetAllChildren(long parentId)
            {
                var queue = new Queue<long>();
                queue.Enqueue(parentId);

                while (queue.Count > 0)
                {
                    var id = queue.Dequeue();
                    var children = db.dbAsnadcratname.Where(p => p.IDre == id);

                    foreach (var child in children)
                    {
                        queue.Enqueue(child.ID);
                        foreach (var item in db.dbAsnafsherkatsanad.Where(s => s.FK_name == child.ID))
                            yield return item;
                    }

                    foreach (var item in db.dbAsnafsherkatsanad.Where(s => s.FK_name == id))
                        yield return item;
                }
            }
        



        public ActionResult gozareshmaly24sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود
            long FK_ID = 0;

            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر


                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    FK_ID += 1;

                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm4 = 0;
                    double summmbedh4 = 0;
                    double sumkol = 0;

                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        double summm = 0;
                        double summmbedh = 0;
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();

                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID).ToList();
                                        dbAsnafsherkatsanad2.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID).ToList();
                                    dbAsnafsherkatsanad2.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID).ToList();
                                dbAsnafsherkatsanad2.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID).ToList();
                            dbAsnafsherkatsanad2.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID).ToList();
                        dbAsnafsherkatsanad2.AddRange(list);
                        foreach (var it in dbAsnafsherkatsanad2)
                        {
                            summm += (double)it.bestankar;
                            summmbedh += (double)it.bedehkar;
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }
                        sumkol = (double)(summm - summmbedh);
                        var gozareshmally1 = new gozareshmally
                        {
                            SUM = sumkol,
                            SUMBEDEHKAR = summmbedh,
                            SUMBESTENKAR = summm,
                            coding = item1.codmahsol,
                            namecoding = item1.name,
                            FK_ID = FK_ID
                        };
                        lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {

                        summm4 += (double)it.bestankar;
                        summmbedh4 += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm4 - summmbedh4);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh4,
                        SUMBESTENKAR = summm4,
                        coding = currentItem.codmahsol,
                        namecoding = currentItem.name,
                        FK_ID = 0,
                        ID = FK_ID
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{

                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}

            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }
        public ActionResult gozareshmaly26sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود
            long FK_ID = 0;

            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر


                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    FK_ID += 1;
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad3 = new List<dbAsnafsherkatsanad>();

                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm4 = 0;
                    double summmbedh4 = 0;
                    double summm41 = 0;
                    double summmbedh41 = 0;
                    double sumkol = 0;

                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        double summm = 0;
                        double summmbedh = 0;
                        double summm11 = 0;
                        double summmbedh11 = 0;
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad4 = new List<dbAsnafsherkatsanad>();

                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                        var list50112 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                        dbAsnafsherkatsanad4.AddRange(list50112);

                                        dbAsnafsherkatsanad2.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                    var list5012 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                    dbAsnafsherkatsanad4.AddRange(list5012);

                                    dbAsnafsherkatsanad2.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                var list502 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                dbAsnafsherkatsanad4.AddRange(list502);

                                dbAsnafsherkatsanad2.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                            var list342 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                            dbAsnafsherkatsanad4.AddRange(list342);

                            dbAsnafsherkatsanad2.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                        var list2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                        dbAsnafsherkatsanad4.AddRange(list2);

                        dbAsnafsherkatsanad2.AddRange(list);
                        foreach (var it in dbAsnafsherkatsanad4)
                        {
                            summm += (double)it.bestankar;
                            summmbedh += (double)it.bedehkar;
                            summm41 += (double)it.bestankar;
                            summmbedh41 += (double)it.bedehkar;
                        }
                        foreach (var it in dbAsnafsherkatsanad2)
                        {
                            summm11 += (double)it.bestankar;
                            summmbedh11 += (double)it.bedehkar;
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }
                        sumkol = (double)(summm - summmbedh- summmbedh11+summm11);
                        var gozareshmally1 = new gozareshmally
                        {
                            SUM = sumkol,
                            SUMBEDEHKAR = summmbedh,
                            SUMBESTENKAR = summm,
                            SUMBEDEHKAR2= summmbedh11,
                            SUMBESTENKAR2 = summm11,

                            coding = item1.codmahsol,
                            namecoding = item1.name,
                            FK_ID = FK_ID
                        };
                        lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                    var list52 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                    dbAsnafsherkatsanad3.AddRange(list52);

                    dbAsnafsherkatsanad.AddRange(list5);
                    double sumx = summm4;
                    double bedhy = summmbedh4;
                    foreach (var it in dbAsnafsherkatsanad3)
                    {

                
                        summm41 += (double)it.bestankar;
                        summmbedh41 += (double)it.bedehkar;
                    }
                    foreach (var it in dbAsnafsherkatsanad)
                    {

                        summm4 += (double)it.bestankar;
                        summmbedh4 += (double)it.bedehkar;
               
                    }
                    sumkol = (double)(summm4 - summmbedh4+ summm41- summmbedh41);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh41,
                        SUMBESTENKAR = summm41,
                        SUMBEDEHKAR2 = summmbedh4,
                        SUMBESTENKAR2 = summm4,
                        coding = item.codmahsol,
                        namecoding = item.name,
                        FK_ID = 0,
                        ID = FK_ID
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{

                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}

            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }


        //public ActionResult gozareshmaly36sotoneh()
        //{
        //    FunctionModel lstMoalefeexcel2 = new FunctionModel();
        //    lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

        //    var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
        //    int count = 0;
        //    //var currentItem = findResult; // شروع از آیتم فعلی


        //    // ادامه جستجو تا زمانی که IDre برابر null شود
        //    long FK_ID = 0;
        //    long FK_ID2 = 0;

        //    foreach (var currentItem in findResult)
        //    {
        //        bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر


        //        var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
        //        foreach (var item in list11)
        //        {
        //            FK_ID += 1;
        //            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad3 = new List<dbAsnafsherkatsanad>();

        //            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
        //            double summm4 = 0;
        //            double summmbedh4 = 0;
        //            double summm41 = 0;
        //            double summmbedh41 = 0;
        //            double sumkol = 0;

        //            var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
        //            foreach (var item1 in listt2)
        //            {
        //                FK_ID2 += 1;

        //                double summm = 0;
        //                double summmbedh = 0;
        //                double summm11 = 0;
        //                double summmbedh11 = 0;
        //                List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
        //                List<dbAsnafsherkatsanad> dbAsnafsherkatsanad4 = new List<dbAsnafsherkatsanad>();

        //                var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
        //                foreach (var item2 in list3)
        //                {
        //                    double summm44 = 0;
        //                    double summmbedh44 = 0;
        //                    double summm1144 = 0;
        //                    double summmbedh1144 = 0;
        //                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad5 = new List<dbAsnafsherkatsanad>();
        //                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad6 = new List<dbAsnafsherkatsanad>();
        //                    var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

        //                    foreach (var item20 in list4)
        //                    {
        //                        var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

        //                        foreach (var item201 in list40)
        //                        {
        //                            var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
        //                            foreach (var item2011 in list401)
        //                            {
        //                                var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

        //                                var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                                var list50112 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                                dbAsnafsherkatsanad6.AddRange(list50112);

        //                                dbAsnafsherkatsanad5.AddRange(list5011);
        //                            }
        //                            var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                            var list5012 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                            dbAsnafsherkatsanad6.AddRange(list5012);

        //                            dbAsnafsherkatsanad5.AddRange(list501);
        //                        }
        //                        var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                        var list502 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                        dbAsnafsherkatsanad4.AddRange(list502);

        //                        dbAsnafsherkatsanad5.AddRange(list50);
        //                    }

        //                    var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                    var list342 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                    dbAsnafsherkatsanad6.AddRange(list342);

        //                    dbAsnafsherkatsanad5.AddRange(list34);

        //                    foreach (var it in dbAsnafsherkatsanad6)
        //                    {
        //                        summm44 += (double)it.bestankar;
        //                        summmbedh44 += (double)it.bedehkar;
        //                        summm += (double)it.bestankar;
        //                        summmbedh += (double)it.bedehkar;
        //                        summm41 += (double)it.bestankar;
        //                        summmbedh41 += (double)it.bedehkar;
        //                    }
        //                    foreach (var it in dbAsnafsherkatsanad5)
        //                    {
        //                        summm1144 += (double)it.bestankar;
        //                        summmbedh1144 += (double)it.bedehkar;
        //                        summm11 += (double)it.bestankar;
        //                        summmbedh11 += (double)it.bedehkar;
        //                        summm4 += (double)it.bestankar;
        //                        summmbedh4 += (double)it.bedehkar;
        //                    }
        //                    sumkol = (double)(summm44 - summmbedh44 - summmbedh1144 + summm1144);
        //                    var gozareshmally2 = new gozareshmally
        //                    {
        //                        SUM = sumkol,
        //                        SUMBEDEHKAR = summmbedh,
        //                        SUMBESTENKAR = summm,
        //                        SUMBEDEHKAR2 = summmbedh11,
        //                        SUMBESTENKAR2 = summm11,

        //                        coding = item1.codmahsol,
        //                        namecoding = item1.name,
        //                        FK_ID = FK_ID2
        //                    };
        //                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally2);




        //                }
        //                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                var list2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                dbAsnafsherkatsanad4.AddRange(list2);

        //                dbAsnafsherkatsanad2.AddRange(list);
        //                foreach (var it in dbAsnafsherkatsanad4)
        //                {
        //                    summm += (double)it.bestankar;
        //                    summmbedh += (double)it.bedehkar;
        //                    summm41 += (double)it.bestankar;
        //                    summmbedh41 += (double)it.bedehkar;
        //                }
        //                foreach (var it in dbAsnafsherkatsanad2)
        //                {
        //                    summm11 += (double)it.bestankar;
        //                    summmbedh11 += (double)it.bedehkar;
        //                    summm4 += (double)it.bestankar;
        //                    summmbedh4 += (double)it.bedehkar;
        //                }
        //                sumkol = (double)(summm - summmbedh - summmbedh11 + summm11);
        //                var gozareshmally1 = new gozareshmally
        //                {
        //                    SUM = sumkol,
        //                    SUMBEDEHKAR = summmbedh,
        //                    SUMBESTENKAR = summm,
        //                    SUMBEDEHKAR2 = summmbedh11,
        //                    SUMBESTENKAR2 = summm11,

        //                    coding = item1.codmahsol,
        //                    namecoding = item1.name,
        //                    FK_ID = FK_ID,
        //                    ID = FK_ID2

        //                };
        //                lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
        //            }
        //            var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //            var list52 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //            dbAsnafsherkatsanad3.AddRange(list52);

        //            dbAsnafsherkatsanad.AddRange(list5);
        //            double sumx = summm4;
        //            double bedhy = summmbedh4;
        //            foreach (var it in dbAsnafsherkatsanad3)
        //            {


        //                summm41 += (double)it.bestankar;
        //                summmbedh41 += (double)it.bedehkar;
        //            }
        //            foreach (var it in dbAsnafsherkatsanad)
        //            {

        //                summm4 += (double)it.bestankar;
        //                summmbedh4 += (double)it.bedehkar;

        //            }
        //            sumkol = (double)(summm4 - summmbedh4 + summm41 - summmbedh41);
        //            var gozareshmally = new gozareshmally
        //            {
        //                SUM = sumkol,
        //                SUMBEDEHKAR = summmbedh41,
        //                SUMBESTENKAR = summm41,
        //                SUMBEDEHKAR2 = summmbedh4,
        //                SUMBESTENKAR2 = summm4,
        //                coding = item.codmahsol,
        //                namecoding = item.name,
        //                FK_ID = 0,
        //                ID = FK_ID
        //            };
        //            lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

        //        }
        //        //if (isFirst)
        //        //{

        //        //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
        //        //    dbAsnafsherkatsanad.AddRange(list);
        //        //    isFirst = false;
        //        //}

        //    }





        //    // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



        //    //    if (findResult != null)
        //    //{
        //    //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

        //    //}


        //    return View(lstMoalefeexcel2);

        //}

        public ActionResult gozareshmaly36sotoneh()
        {
            var model = new FunctionModel { gozareshmally = new List<gozareshmally>() };
            long idCounter = 1;

            var roots = db.dbAsnadcratname.Where(x => x.IDre == null).ToList();

            foreach (var root in roots)
            {
                AddHierarchyItems2(root, 0, null, model.gozareshmally, ref idCounter);
            }

            return View(model);
        }

        private (
    double totalBestankarNon1, double totalBedehkarNon1,
    double totalBestankar1, double totalBedehkar1
) AddHierarchyItems2(
    dbAsnadcratname current,
    int level,
    long? parentId,
    List<gozareshmally> list,
    ref long idCounter)
        {
            // جمع‌های بدهکار و بستانکار برای شرط != "1"
            double ownBestankarNon1 = db.dbAsnafsherkatsanad
          .Where(x => x.FK_name == current.ID && x.dbAsnadshomrehsanad.shomarehsanad != "1")
          .Select(x => (double?)x.bestankar)
          .DefaultIfEmpty(0)
          .Sum() ?? 0;

            double ownBedehkarNon1 = db.dbAsnafsherkatsanad
                .Where(x => x.FK_name == current.ID && x.dbAsnadshomrehsanad.shomarehsanad != "1")
                .Select(x => (double?)x.bedehkar)
                .DefaultIfEmpty(0)
                .Sum() ?? 0;

            double ownBestankar1 = db.dbAsnafsherkatsanad
                .Where(x => x.FK_name == current.ID && x.dbAsnadshomrehsanad.shomarehsanad == "1")
                .Select(x => (double?)x.bestankar)
                .DefaultIfEmpty(0)
                .Sum() ?? 0;

            double ownBedehkar1 = db.dbAsnafsherkatsanad
                .Where(x => x.FK_name == current.ID && x.dbAsnadshomrehsanad.shomarehsanad == "1")
                .Select(x => (double?)x.bedehkar)
                .DefaultIfEmpty(0)
                .Sum() ?? 0;

            // مجموع کل (خود نود + فرزندان)
            double totalBestankarNon1 = ownBestankarNon1;
            double totalBedehkarNon1 = ownBedehkarNon1;

            double totalBestankar1 = ownBestankar1;
            double totalBedehkar1 = ownBedehkar1;

            var rowId = idCounter++;

            var children = db.dbAsnadcratname.Where(x => x.IDre == current.ID).ToList();

            foreach (var child in children)
            {
                var (childBestankarNon1, childBedehkarNon1, childBestankar1, childBedehkar1) =
                    AddHierarchyItems2(child, level + 1, rowId, list, ref idCounter);

                totalBestankarNon1 += childBestankarNon1;
                totalBedehkarNon1 += childBedehkarNon1;
                totalBestankar1 += childBestankar1;
                totalBedehkar1 += childBedehkar1;
            }

           
                list.Add(new gozareshmally
                {
                    ID = rowId,
                    FK_ID = parentId ?? 0,
                    coding = current.codmahsol,
                    namecoding = current.name,
                    SUMBESTENKAR = totalBestankarNon1,
                    SUMBEDEHKAR = totalBedehkarNon1,
                    SUMBESTENKAR2 = totalBestankar1,
                    SUMBEDEHKAR2 = totalBedehkar1,
                    SUM = (totalBestankarNon1 + totalBestankar1) - (totalBedehkarNon1 + totalBedehkar1),
                });
            

            return (totalBestankarNon1, totalBedehkarNon1, totalBestankar1, totalBedehkar1);
        }


        public ActionResult gozareshmaly28sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود
            long FK_ID = 0;

            foreach (var currentItem in findResult)
            {
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر


                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    FK_ID += 1;
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad3 = new List<dbAsnafsherkatsanad>();

                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm4 = 0;
                    double summmbedh4 = 0;
                    double summm41 = 0;
                    double summmbedh41 = 0;
                    double sumkol = 0;

                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        double summm = 0;
                        double summmbedh = 0;
                        double summm11 = 0;
                        double summmbedh11 = 0;
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
                        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad4 = new List<dbAsnafsherkatsanad>();

                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                        var list50112 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                        dbAsnafsherkatsanad4.AddRange(list50112);

                                        dbAsnafsherkatsanad2.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                    var list5012 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                    dbAsnafsherkatsanad4.AddRange(list5012);

                                    dbAsnafsherkatsanad2.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                var list502 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                dbAsnafsherkatsanad4.AddRange(list502);

                                dbAsnafsherkatsanad2.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                            var list342 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                            dbAsnafsherkatsanad4.AddRange(list342);

                            dbAsnafsherkatsanad2.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                        var list2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                        dbAsnafsherkatsanad4.AddRange(list2);

                        dbAsnafsherkatsanad2.AddRange(list);
                        foreach (var it in dbAsnafsherkatsanad4)
                        {
                            summm += (double)it.bestankar;
                            summmbedh += (double)it.bedehkar;
                            summm41 += (double)it.bestankar;
                            summmbedh41 += (double)it.bedehkar;
                        }
                        foreach (var it in dbAsnafsherkatsanad2)
                        {
                            summm11 += (double)it.bestankar;
                            summmbedh11 += (double)it.bedehkar;
                            summm4 += (double)it.bestankar;
                            summmbedh4 += (double)it.bedehkar;
                        }
                        sumkol = (double)(summm - summmbedh - summmbedh11 + summm11);
                        var gozareshmally1 = new gozareshmally
                        {
                            SUM = sumkol,
                            SUMBEDEHKAR = summmbedh,
                            SUMBESTENKAR = summm,
                            SUMBEDEHKAR2 = summmbedh11,
                            SUMBESTENKAR2 = summm11,

                            coding = item1.codmahsol,
                            namecoding = item1.name,
                            FK_ID = FK_ID
                        };
                        lstMoalefeexcel2.gozareshmally.Add(gozareshmally1);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                    var list52 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                    dbAsnafsherkatsanad3.AddRange(list52);

                    dbAsnafsherkatsanad.AddRange(list5);
                    double sumx = summm4;
                    double bedhy = summmbedh4;
                    foreach (var it in dbAsnafsherkatsanad3)
                    {


                        summm41 += (double)it.bestankar;
                        summmbedh41 += (double)it.bedehkar;
                    }
                    foreach (var it in dbAsnafsherkatsanad)
                    {

                        summm4 += (double)it.bestankar;
                        summmbedh4 += (double)it.bedehkar;

                    }
                    sumkol = (double)(summm4 - summmbedh4 + summm41 - summmbedh41);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh41,
                        SUMBESTENKAR = summm41,
                        SUMBEDEHKAR2 = summmbedh4,
                        SUMBESTENKAR2 = summm4,
                        coding = item.codmahsol,
                        namecoding = item.name,
                        FK_ID = 0,
                        ID = FK_ID
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }
                //if (isFirst)
                //{

                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}

            }





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }

        public ActionResult gozareshmaly16sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود


            foreach (var currentItem in findResult)
            {
          
                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر
     
                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
                    double summm2 = 0;
                    double summmbedh2 = 0;
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm3 = 0;
                    double summmbedh3 = 0;
                    double summm = 0;
                    double summmbedh = 0;
                    double sumkol = 0;
                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                        var list50112 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                        dbAsnafsherkatsanad2.AddRange(list50112);

                                        dbAsnafsherkatsanad.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                    var list5012 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                    dbAsnafsherkatsanad2.AddRange(list5012);

                                    dbAsnafsherkatsanad.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                var list502 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                dbAsnafsherkatsanad2.AddRange(list502);

                                dbAsnafsherkatsanad.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                            var list342 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                            dbAsnafsherkatsanad2.AddRange(list342);

                            dbAsnafsherkatsanad.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                        var list2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                        dbAsnafsherkatsanad2.AddRange(list2);

                        dbAsnafsherkatsanad.AddRange(list2);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                    var list52 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                    dbAsnafsherkatsanad2.AddRange(list52);

                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {
                        summm += (double)it.bestankar;
                        summmbedh += (double)it.bedehkar;
                    }
                    foreach (var it in dbAsnafsherkatsanad2)
                    {
                        summm2 += (double)it.bestankar;
                        summmbedh2 += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm - summmbedh + summm2 - summmbedh2);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh,
                        SUMBESTENKAR = summm,
                        SUMBEDEHKAR2 = summmbedh2,
                        SUMBESTENKAR2 = summm2,
                        coding = item.codmahsol,
                        namecoding = item.name
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);
      
                }





             
            }








                //if (isFirst)
                //{
                //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                //    dbAsnafsherkatsanad.AddRange(list);
                //    isFirst = false;
                //}
           





            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }

        //public class GozareshNode
        //{
        //    public long ID { get; set; }
        //    public long? FK_ID { get; set; }
        //    public string Coding { get; set; }
        //    public string NameCoding { get; set; }
        //    public double SumBedehkar { get; set; }
        //    public double SumBestenkar { get; set; }
        //    public double Sum => SumBedehkar - SumBestenkar;

        //    public List<GozareshNode> Children { get; set; } = new();
        //}
        //public List<GozareshNode> BuildGozareshTree(List<gozareshmally> rawData)
        //{
        //    var lookup = rawData.ToDictionary(x => x.ID, x => new GozareshNode
        //    {
        //        ID = x.ID,
        //        FK_ID = x.FK_ID,
        //        Coding = x.coding,
        //        NameCoding = x.namecoding,
        //        SumBedehkar = x.SUMBEDEHKAR,
        //        SumBestenkar = x.SUMBESTENKAR
        //    });

        //    List<GozareshNode> roots = new();

        //    foreach (var node in lookup.Values)
        //    {
        //        if (node.FK_ID != null && lookup.ContainsKey(node.FK_ID.Value))
        //        {
        //            lookup[node.FK_ID.Value].Children.Add(node);
        //        }
        //        else
        //        {
        //            roots.Add(node);
        //        }
        //    }

        //    return roots;
        //}

        //public ActionResult ShowGozaresh()
        //{
        //    FunctionModel lstMoalefeexcel2 = new FunctionModel();
        //    lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

        //    // پر کردن دستی یا با فیلتر
        //    // مثلاً: lstMoalefeexcel2.gozareshmally = db.YourCustomFilteredList();

        //    var gozareshTree = BuildGozareshTree(lstMoalefeexcel2.gozareshmally);

        //    return View(gozareshTree);
        //}



        public ActionResult gozareshmaly18sotoneh()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

            var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            int count = 0;
            //var currentItem = findResult; // شروع از آیتم فعلی


            // ادامه جستجو تا زمانی که IDre برابر null شود


            foreach (var currentItem in findResult)
            {

                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر

                var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                foreach (var item in list11)
                {
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
                    double summm2 = 0;
                    double summmbedh2 = 0;
                    List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                    double summm3 = 0;
                    double summmbedh3 = 0;
                    double summm = 0;
                    double summmbedh = 0;
                    double sumkol = 0;
                    var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
                    foreach (var item1 in listt2)
                    {
                        var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
                        foreach (var item2 in list3)
                        {
                            var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

                            foreach (var item20 in list4)
                            {
                                var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

                                foreach (var item201 in list40)
                                {
                                    var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
                                    foreach (var item2011 in list401)
                                    {
                                        var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

                                        var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                        var list50112 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                        dbAsnafsherkatsanad2.AddRange(list50112);

                                        dbAsnafsherkatsanad.AddRange(list5011);
                                    }
                                    var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                    var list5012 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                    dbAsnafsherkatsanad2.AddRange(list5012);

                                    dbAsnafsherkatsanad.AddRange(list501);
                                }
                                var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                                var list502 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                                dbAsnafsherkatsanad2.AddRange(list502);

                                dbAsnafsherkatsanad.AddRange(list50);
                            }

                            var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                            var list342 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                            dbAsnafsherkatsanad2.AddRange(list342);

                            dbAsnafsherkatsanad.AddRange(list34);
                        }
                        var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                        var list2 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                        dbAsnafsherkatsanad2.AddRange(list2);

                        dbAsnafsherkatsanad.AddRange(list2);
                    }
                    var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
                    var list52 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
                    dbAsnafsherkatsanad2.AddRange(list52);

                    dbAsnafsherkatsanad.AddRange(list5);
                    foreach (var it in dbAsnafsherkatsanad)
                    {
                        summm += (double)it.bestankar;
                        summmbedh += (double)it.bedehkar;
                        summm3 += (double)it.bestankar;
                        summmbedh3 += (double)it.bedehkar;
                    }
                    foreach (var it in dbAsnafsherkatsanad2)
                    {
                        summm2 += (double)it.bestankar;
                        summmbedh2 += (double)it.bedehkar;
                        summm3 += (double)it.bestankar;
                        summmbedh3 += (double)it.bedehkar;
                    }
                    sumkol = (double)(summm - summmbedh + summm2 - summmbedh2);
                    var gozareshmally = new gozareshmally
                    {
                        SUM = sumkol,
                        SUMBEDEHKAR = summmbedh,
                        SUMBESTENKAR = summm,
                        SUMBEDEHKAR2 = summmbedh2,
                        SUMBESTENKAR2 = summm2,
                        SUMBEDEHKAR4 = summmbedh3,
                        SUMBESTENKAR4 = summm3,
                        coding = item.codmahsol,
                        namecoding = item.name
                    };
                    lstMoalefeexcel2.gozareshmally.Add(gozareshmally);

                }






            }








            //if (isFirst)
            //{
            //    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
            //    dbAsnafsherkatsanad.AddRange(list);
            //    isFirst = false;
            //}






            // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(lstMoalefeexcel2);

        }


        //public ActionResult gozareshmaly18sotoneh()
        //{
        //    FunctionModel lstMoalefeexcel2 = new FunctionModel();
        //    lstMoalefeexcel2.gozareshmally = new List<gozareshmally>();

        //    var findResult = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
        //    int count = 0;
        //    //var currentItem = findResult; // شروع از آیتم فعلی


        //    // ادامه جستجو تا زمانی که IDre برابر null شود


        //    foreach (var currentItem in findResult)
        //    {
        //        bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر
        //        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
        //        List<dbAsnafsherkatsanad> dbAsnafsherkatsanad2 = new List<dbAsnafsherkatsanad>();
        //        double summm2 = 0;
        //        double summmbedh2 = 0;
        //        double summm3 = 0;
        //        double summmbedh3 = 0;
        //        double summm = 0;
        //        double summmbedh = 0;
        //        double sumkol = 0;
        //        var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
        //        foreach (var item in list11)
        //        {
        //            var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
        //            foreach (var item1 in listt2)
        //            {
        //                var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
        //                foreach (var item2 in list3)
        //                {
        //                    var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

        //                    foreach (var item20 in list4)
        //                    {
        //                        var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

        //                        foreach (var item201 in list40)
        //                        {
        //                            var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
        //                            foreach (var item2011 in list401)
        //                            {
        //                                var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

        //                                var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                                dbAsnafsherkatsanad.AddRange(list5011);
        //                            }
        //                            var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                            dbAsnafsherkatsanad.AddRange(list501);
        //                        }
        //                        var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                        dbAsnafsherkatsanad.AddRange(list50);
        //                    }

        //                    var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                    dbAsnafsherkatsanad.AddRange(list34);
        //                }
        //                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //                dbAsnafsherkatsanad.AddRange(list);
        //            }
        //            var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID && s.dbAsnadshomrehsanad.shomarehsanad == "1").ToList();
        //            dbAsnafsherkatsanad.AddRange(list5);
        //        }





        //        var list1111 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
        //        foreach (var item in list1111)
        //        {
        //            var listt2 = db.dbAsnadcratname.Where(p => p.IDre == item.ID).ToList();
        //            foreach (var item1 in listt2)
        //            {
        //                var list3 = db.dbAsnadcratname.Where(p => p.IDre == item1.ID).ToList();
        //                foreach (var item2 in list3)
        //                {
        //                    var list4 = db.dbAsnadcratname.Where(p => p.IDre == item2.ID).ToList();

        //                    foreach (var item20 in list4)
        //                    {
        //                        var list40 = db.dbAsnadcratname.Where(p => p.IDre == item20.ID).ToList();

        //                        foreach (var item201 in list40)
        //                        {
        //                            var list401 = db.dbAsnadcratname.Where(p => p.IDre == item201.ID).ToList();
        //                            foreach (var item2011 in list401)
        //                            {
        //                                var list40123 = db.dbAsnadcratname.Where(p => p.IDre == item2011.ID).ToList();

        //                                var list5011 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2011.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                                dbAsnafsherkatsanad2.AddRange(list5011);
        //                            }
        //                            var list501 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item201.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                            dbAsnafsherkatsanad2.AddRange(list501);
        //                        }
        //                        var list50 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item20.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                        dbAsnafsherkatsanad2.AddRange(list50);
        //                    }

        //                    var list34 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item2.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                    dbAsnafsherkatsanad2.AddRange(list34);
        //                }
        //                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item1.ID && s.dbAsnadshomrehsanad.shomarehsanad != "1").ToList();
        //                dbAsnafsherkatsanad2.AddRange(list);
        //            }
        //            var list5 = db.dbAsnafsherkatsanad.Where(s => s.FK_name == item.ID).ToList();
        //            dbAsnafsherkatsanad2.AddRange(list5);
        //        }








        //        if (isFirst)
        //        {
        //            var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
        //            dbAsnafsherkatsanad.AddRange(list);
        //            isFirst = false;
        //        }
        //        foreach (var it in dbAsnafsherkatsanad)
        //        {
        //            summm += (double)it.bestankar;
        //            summmbedh += (double)it.bedehkar;
        //            summm3 += (double)it.bestankar;
        //            summmbedh3 += (double)it.bedehkar;
        //        }
        //        foreach (var it in dbAsnafsherkatsanad2)
        //        {
        //            summm2 += (double)it.bestankar;
        //            summmbedh2 += (double)it.bedehkar;
        //            summm3 += (double)it.bestankar;
        //            summmbedh3 += (double)it.bedehkar;
        //        }
        //        sumkol = (double)(summm - summmbedh + summm2 - summmbedh2);
        //        var gozareshmally = new gozareshmally
        //        {
        //            SUM = sumkol,
        //            SUMBEDEHKAR = summmbedh,
        //            SUMBESTENKAR = summm,
        //            SUMBEDEHKAR2 = summmbedh2,
        //            SUMBESTENKAR2 = summm2,
        //            SUMBEDEHKAR4 = summmbedh3,
        //            SUMBESTENKAR4 = summm3,
        //            coding = currentItem.codmahsol,
        //            namecoding = currentItem.name
        //        };
        //        lstMoalefeexcel2.gozareshmally.Add(gozareshmally);
        //    }





        //    // اگر رکورد پیدا شد، آن را به tolll اضافه کنید



        //    //    if (findResult != null)
        //    //{
        //    //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

        //    //}


        //    return View(lstMoalefeexcel2);

        //}

        public ActionResult EXCEL_Asnad()
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = EXCEL_Asnadupload();

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "EXCEL_ASNAD" + extension);

        }
        private Workbook EXCEL_Asnadupload()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EXCEL_Asnad.xlsx"));
            Row Row;
            //var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null).ToList();
            //var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var usr = db.dbAsnadcratname.ToList();
            //var company = db.tbCompanies.ToList();
            //var job = db.tbjob.ToList();
            int counter = 1;
            var count = 1;
            var count1 = 1;
            var count2 = 1;

            var count3 = 2;


            foreach (var item in usr)
            {


                var fin = db.dbAsnadcratname.Where(p => p.ID == item.ID).FirstOrDefault();
                string tolll = "";
                var currentItem = fin; // شروع از آیتم فعلی

                bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر

                // ادامه جستجو تا زمانی که IDre برابر null شود
                while (currentItem != null && currentItem.IDre != null)
                {
                    // پیدا کردن رکورد بعدی با استفاده از IDre فعلی
                    currentItem = db.dbAsnadcratname.Where(p => p.ID == currentItem.IDre).FirstOrDefault();

                    // اگر رکورد پیدا شد، آن را به tolll اضافه کنید
                    if (currentItem != null)
                    {
                        if (isFirst)
                        {
                            tolll += currentItem.name;

                            tolll += "_";
                            // در ابتدا بدون فاصله
                            isFirst = false; // بعد از اولین اضافه کردن، isFirst را false می‌کنیم
                        }
                        else
                        {
                            tolll += "_" + currentItem.name; // اگر اولین رکورد نیست، فاصله اضافه می‌شود
                        }
                    }
                }


                Row = new Row() { Height = 20, Index = count1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.name,
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
                        Value = item.codmahsol,
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
                        Value = tolll,
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
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            //foreach (var item in moalfe)
            //{
            //    Row = new Row() { Height = 20, Index = count2 };
            //    Row.AddCells(new List<Cell>()
            //    {
            //        new Cell()
            //        {
            //            Value = item.md_Title,
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
            //    count2++;
            //    Moalefeexcelfile.Sheets[1].AddRow(Row);
            //}
            //foreach (var item in company)
            //{
            //    Row = new Row() { Height = 20, Index = counter };
            //    Row.AddCells(new List<Cell>()
            //    {
            //        new Cell()
            //        {
            //            Value = item.CompanyName+"-"+(item.tbUsers.FullName),
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = 1
            //        },
            //             new Cell()
            //        {
            //            Value = item.ID,
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = 2
            //        },
            //    });
            //    counter++;
            //    Moalefeexcelfile.Sheets[1].AddRow(Row);
            //}
            //foreach (var item in job)
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
            //            Index = 6
            //        },
            //             new Cell()
            //        {
            //            Value = item.ID,
            //            FontFamily = "B Nazanin",
            //            Bold = false,
            //            Enable = true,
            //            Wrap = false,
            //            FontSize = 12,
            //            Italic = false,
            //            Underline = false,
            //            Index = 7
            //        },
            //    });
            //    count1++;
            //    Moalefeexcelfile.Sheets[1].AddRow(Row);
            //}
            return Moalefeexcelfile;








        }

        public async Task<ActionResult> viewdetail(int id = 0, int month = 0, int year = 0)
        {
            var tbmarahesabt411 = await db.tbmarahesabt4.ToListAsync();
            var tbmarahlsabtbastedit1 = await db.tbmarahlsabtbastedit1.ToListAsync();
            var tbSavedFunctions = await db.tbSavedFunctions.ToListAsync();
            var tbmoalfefishexcel = await db.tbmoalfefishexcel.ToListAsync();
            var tbmarahesabt4 = db.tbmarahesabt4.FirstOrDefault(p => p.ID == id);
            var tbmarahlsabtbastedit2 = db.tbmarahlsabtbastedit2.Where(p => p.FK_City == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City && p.FK_namebAST == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn).ToList();
            var tbmarahesabt41 = await db.tbmarahesabt4.ToListAsync();
            var tbEquipmentMoalefeValueReffrenceSave = await db.tbEquipmentMoalefeValueReffrenceSave.ToListAsync();

            if (tbmarahesabt4 == null)
            {
                return HttpNotFound(); // اگر tbmarahesabt4 پیدا نشد، خطا بدهد
            }
            PersianCalendar persianCalendar22 = new PersianCalendar();

            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.tbdetail = new List<tbdetail>();

            if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
                .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 1) == true)
            {

                foreach (var ity in tbmarahlsabtbastedit2)
                {





                    var findlonk = tbmarahesabt41.Where(p => p.FK_basteh == ity.ID && p.Month <= month && p.Year <= year).ToList();
                    foreach (var ytu in findlonk)
                    {
                        var x1 = await db.tbSavedFunctions
             .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.submitt == true && p.del != true && p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                       && p.svdfunc_UserSaveID == ytu.FK_usr && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year))
             .FirstOrDefaultAsync();
                        int dayasli = 0;
                        string DAY1 = "";
                        int day4 = 1;
                        DateTime date1223 = persianCalendar22.ToDateTime(year, month, day4, 0, 0, 0, 0);

                        //var findd2343 = tbmarahesabt4.Where(p => p.FK_usr == ytu.FK_usr && p.del != true).ToList();
                        if (tbmarahesabt4.sabt == true)
                        {
                            int day2 = tbmarahesabt4.day2 ?? 1;
                            date1223 = persianCalendar22.ToDateTime(year, month, day2, 0, 0, 0, 0);

                            // ساختن تاریخ کامل شمسی از مقادیر year, month, day2

                            // دریافت تاریخ فعلی به صورت شمسی
                            DateTime now = DateTime.Now;

                            if (x1 != null)
                            {
                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x1.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
.FirstOrDefault();
                                if (finddd11 != null)
                                {
                                    now = (DateTime)finddd11.svdfunc_SavedDateTime;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //else if (fd112345 != null)
                            //{
                            //    var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.del != true).FirstOrDefault();
                            //    if (tbmoalfefishexcel11 != null)
                            //    {
                            //        now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //    }
                            //}
                            //else if (fd21 != null)
                            //{
                            //    var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.del != true).FirstOrDefault();
                            //    if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //    {
                            //        now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //    }
                            //}
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);
                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);

                            // محاسبه تفاوت روزها
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                            if (dayasli < 0)
                            {
                                DAY1 = (difference.Days.ToString());

                            }
                            else
                            {
                                DAY1 = difference.Days.ToString();

                            }
                        }
                        else if (tbmarahesabt4.control == true)
                        {
                            int day3 = 0;
                            bool cheack = false;
                            int month22 = month;
                            int yer22 = year;
                            List<tbmarahesabt4> tbmarahesabt4143 = new List<tbmarahesabt4>();
                            List<tbmarahesabt4> tbmarahesabt421 = new List<tbmarahesabt4>();

                            foreach (var ty1 in tbmarahesabt4143.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true).ToList())
                            {
                                if (tbmarahesabt4143.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt4143.Add(ty1);

                                }
                            }
                            foreach (var ty1 in tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true && p.number <= tbmarahesabt4.number).ToList())
                            {
                                if (tbmarahesabt4143.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt4143.Add(ty1);

                                }
                            }

                            foreach (var ty in tbmarahesabt4143)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }
                            foreach (var ty in tbmarahesabt421)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }
                            // پیدا کردن آخرین رکورد
                            //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh&&p.taeed==true).OrderByDescending(s => s.number).FirstOrDefault();
                            //int day2 = findcontrol?.day2 ?? 1;
                            //int additionalDays = t.mohlat ?? 1;

                            //// محاسبه مجموع روزها
                            //int day3 = day2 + additionalDays;

                            //// تقویم شمسی برای بررسی تعداد روزهای ماه

                            //while (day3 > persianCalendar.GetDaysInMonth(yer22, month22))
                            //{
                            //    // تعداد روزهای ماه جاری
                            //    int daysInCurrentMonth = persianCalendar.GetDaysInMonth(yer22, month22);

                            //    // کاهش تعداد روزهای ماه جاری از مجموع روزها
                            //    day3 -= daysInCurrentMonth;

                            //    // به ماه بعدی برویم
                            //    month22++;

                            //    // اگر ماه از 12 گذشت، به سال بعد برویم
                            //    if (month22 > 12)
                            //    {
                            //        month22 = 1;
                            //        yer22++;
                            //    }
                            //}

                            // محاسبه تاریخ نهایی
                            date1223 = persianCalendar22.ToDateTime(yer22, month22, day3, 0, 0, 0, 0);
                            DateTime now = DateTime.Now;
                            if (x1 != null)
                            {
                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x1.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
.FirstOrDefault();
                                if (finddd11 != null)
                                {
                                    now = (DateTime)finddd11.svdfunc_SavedDateTime;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //else if (fd112345 != null)
                            //{
                            //    var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //    if (tbmoalfefishexcel11 != null)
                            //    {
                            //        now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //    }
                            //}
                            //else if (fd21 != null)
                            //{
                            //    var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //    if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //    {
                            //        now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //    }
                            //}
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);

                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                        }
                        else if (tbmarahesabt4.taeed == true)
                        {
                            int month22 = month;
                            int yer22 = year;
                            int day3 = 0;
                            bool cheack = false;
                            List<tbmarahesabt4> tbmarahesabt41123 = new List<tbmarahesabt4>();
                            foreach (var ty1 in tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true && p.number <= tbmarahesabt4.number).ToList())
                            {
                                if (tbmarahesabt41123.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt41123.Add(ty1);

                                }
                            }

                            foreach (var ty in tbmarahesabt41123)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }

                            date1223 = persianCalendar22.ToDateTime(yer22, month22, day3, 0, 0, 0, 0);

                            // پیدا کردن آخرین رکورد

                            DateTime now = DateTime.Now;
                            if (x1 != null)
                            {
                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x1.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
.FirstOrDefault();
                                if (finddd11 != null)
                                {
                                    now = (DateTime)finddd11.svdfunc_SavedDateTime;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //                            if (fd1111 != null)
                            //                            {
                            //                                var finddd11 = tbSavedFunctions
                            //.Where(p => p.FK_tbmarahesabt4 == t.ID && p.FK_Basteh == fd1111.ID && p.del != true && p.svdfunc_UserSaveID == t.FK_usr
                            //&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
                            //.FirstOrDefault();
                            //                                if (finddd11 != null)
                            //                                {
                            //                                    now = (DateTime)finddd11.svdfunc_SavedDateTime;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //                                }

                            //                            }
                            //                            else if (fd112345 != null)
                            //                            {
                            //                                var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //                                if (tbmoalfefishexcel11 != null)
                            //                                {
                            //                                    now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //                                }
                            //                            }
                            //                            else if (fd21 != null)
                            //                            {
                            //                                var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //                                if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //                                {
                            //                                    now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //                                }
                            //                            }
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);

                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                            if (dayasli < 0)
                            {
                                DAY1 = (difference.Days.ToString());

                            }
                            else
                            {
                                DAY1 = difference.Days.ToString();

                            }
                        }
                        var x = await db.tbSavedFunctions
                  .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true && p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                            && p.svdfunc_UserSaveID == ytu.FK_usr && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year))
                  .FirstOrDefaultAsync();

                        var x2 = await db.tbSavedFunctions
       .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del == true && p.svdfunc_pymnID == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                 && p.svdfunc_UserSaveID == ytu.FK_usr && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year))
       .FirstOrDefaultAsync();
                        if (x1 != null) // بررسی مقدار t
                        {
                            var persianCalendar = new System.Globalization.PersianCalendar();
                            var savedDate = (DateTime)x1.svdfunc_SavedDateTime;

                            // تبدیل تاریخ به شمسی
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                vazrt = 1,
                                tbmarahesabt4 = x1.tbmarahesabt4,
                                tbSavedFunctions = x1,
                                day = 1,
                                vaziat = "موفق",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                month = month,
                                year = year,
                                daymondeh = dayasli,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته
                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }

                        else if (x != null) // بررسی مقدار t
                        {
                            var persianCalendar = new System.Globalization.PersianCalendar();
                            var savedDate = (DateTime)x.svdfunc_SavedDateTime;

                            // تبدیل تاریخ به شمسی
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                daymondeh = dayasli,
                                vazrt = 1,
                                tbmarahesabt4 = x.tbmarahesabt4,
                                tbSavedFunctions = x,
                                vaziat = "در انتظار ارسال",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }
                        else if (x2 != null) // بررسی مقدار t
                        {
                            var persianCalendar = new System.Globalization.PersianCalendar();
                            var savedDate = (DateTime)x2.svdfunc_SavedDateTime;

                            // تبدیل تاریخ به شمسی
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                vazrt = 1,
                                tbmarahesabt4 = x2.tbmarahesabt4,
                                tbSavedFunctions = x2,
                                vaziat = "اصلاح برگشت",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                day = 1,
                                daymondeh = dayasli,
                                month = month,
                                year = year,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }
                        else
                        {
                            var tbdetail = new tbdetail
                            {
                                vazrt = 1,
                                tbmarahesabt4 = ytu,
                                daymondeh = dayasli,
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                vaziat = "در انتظار ثبت",
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = "" // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }

                    }
                }


            }
            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
                .Any(s => s.FK_taghiz != null) == true)
            {
                var x = await db.tbEquipmentMoalefeValueReffrenceSave
                    .Where(p => p.FK_Basteh == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.FK_pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                                && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year))
                    .ToListAsync();

                foreach (var t in x)
                {
                    if (t != null)
                    {
                        var tbdetail = new tbdetail
                        {
                            tbmarahesabt4 = t.tbmarahesabt4,
                            tbEquipmentMoalefeValueReffrenceSave = t,
                            day = 1,
                            month = month,
                            year = year
                        };
                        lstMoalefeexcel2.tbdetail.Add(tbdetail);
                    }
                }
            }
            else if (tbmarahesabt4?.tbmarahlsabtbastedit2?.tbmarahlsabtbastedit1?.tbmarahlsabtbastedit3
                .Any(s => s.FK_moalfe != null && s.tbmarahlsabtbastedit1.typemoalfeh == 2) == true)
            {


                int dayasli = 0;
                string DAY1 = "";
                int day4 = 1;
                DateTime date1223 = persianCalendar22.ToDateTime(year, month, day4, 0, 0, 0, 0);

                foreach (var ity in tbmarahlsabtbastedit2)
                {
                    var findlonk = tbmarahesabt41.Where(p => p.FK_basteh == ity.ID && p.Month <= month && p.Year <= year && p.control != true).ToList();
                    foreach (var ytu in findlonk)
                    {
                        var x = await db.tbmoalfefishexcel
                    .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true && p.submitt == true && p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                           && p.usr_sabt == ytu.FK_usr && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year))
                    .FirstOrDefaultAsync();

                        if (tbmarahesabt4.sabt == true)
                        {
                            int day2 = tbmarahesabt4.day2 ?? 1;
                            date1223 = persianCalendar22.ToDateTime(year, month, day2, 0, 0, 0, 0);

                            // ساختن تاریخ کامل شمسی از مقادیر year, month, day2

                            // دریافت تاریخ فعلی به صورت شمسی
                            DateTime now = DateTime.Now;

                            if (x != null)
                            {
                                //                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
                                //&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
                                //.FirstOrDefault();
                                if (x != null)
                                {
                                    now = (DateTime)x.Datatmie;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //else if (fd112345 != null)
                            //{
                            //    var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.del != true).FirstOrDefault();
                            //    if (tbmoalfefishexcel11 != null)
                            //    {
                            //        now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //    }
                            //}
                            //else if (fd21 != null)
                            //{
                            //    var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.del != true).FirstOrDefault();
                            //    if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //    {
                            //        now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //    }
                            //}
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);
                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);

                            // محاسبه تفاوت روزها
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                            if (dayasli < 0)
                            {
                                DAY1 = (difference.Days.ToString());

                            }
                            else
                            {
                                DAY1 = difference.Days.ToString();

                            }
                        }
                        else if (tbmarahesabt4.control == true)
                        {
                            int day3 = 0;
                            bool cheack = false;
                            int month22 = month;
                            int yer22 = year;
                            List<tbmarahesabt4> tbmarahesabt4143 = new List<tbmarahesabt4>();
                            List<tbmarahesabt4> tbmarahesabt421 = new List<tbmarahesabt4>();

                            foreach (var ty1 in tbmarahesabt4143.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true).ToList())
                            {
                                if (tbmarahesabt4143.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt4143.Add(ty1);

                                }
                            }
                            foreach (var ty1 in tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.control == true && p.number <= tbmarahesabt4.number).ToList())
                            {
                                if (tbmarahesabt4143.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt4143.Add(ty1);

                                }
                            }

                            foreach (var ty in tbmarahesabt4143)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }
                            foreach (var ty in tbmarahesabt421)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }
                            // پیدا کردن آخرین رکورد
                            //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh&&p.taeed==true).OrderByDescending(s => s.number).FirstOrDefault();
                            //int day2 = findcontrol?.day2 ?? 1;
                            //int additionalDays = t.mohlat ?? 1;

                            //// محاسبه مجموع روزها
                            //int day3 = day2 + additionalDays;

                            //// تقویم شمسی برای بررسی تعداد روزهای ماه

                            //while (day3 > persianCalendar.GetDaysInMonth(yer22, month22))
                            //{
                            //    // تعداد روزهای ماه جاری
                            //    int daysInCurrentMonth = persianCalendar.GetDaysInMonth(yer22, month22);

                            //    // کاهش تعداد روزهای ماه جاری از مجموع روزها
                            //    day3 -= daysInCurrentMonth;

                            //    // به ماه بعدی برویم
                            //    month22++;

                            //    // اگر ماه از 12 گذشت، به سال بعد برویم
                            //    if (month22 > 12)
                            //    {
                            //        month22 = 1;
                            //        yer22++;
                            //    }
                            //}

                            // محاسبه تاریخ نهایی
                            date1223 = persianCalendar22.ToDateTime(yer22, month22, day3, 0, 0, 0, 0);
                            DateTime now = DateTime.Now;
                            if (x != null)
                            {
                                //                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
                                //&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
                                //.FirstOrDefault();
                                if (x != null)
                                {
                                    now = (DateTime)x.Datatmie;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //else if (fd112345 != null)
                            //{
                            //    var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //    if (tbmoalfefishexcel11 != null)
                            //    {
                            //        now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //    }
                            //}
                            //else if (fd21 != null)
                            //{
                            //    var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //    if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //    {
                            //        now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //        // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //        // محاسبه تعداد روزهای ماه
                            //        //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //    }
                            //}
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);

                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                        }
                        else if (tbmarahesabt4.taeed == true)
                        {
                            int month22 = month;
                            int yer22 = year;
                            int day3 = 0;
                            bool cheack = false;
                            List<tbmarahesabt4> tbmarahesabt41123 = new List<tbmarahesabt4>();
                            foreach (var ty1 in tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.taeed == true && p.number <= tbmarahesabt4.number).ToList())
                            {
                                if (tbmarahesabt41123.Where(s => s.number == ty1.number).FirstOrDefault() == null)
                                {
                                    tbmarahesabt41123.Add(ty1);

                                }
                            }

                            foreach (var ty in tbmarahesabt41123)
                            {
                                if (cheack == false)
                                {
                                    var findcontrol = tbmarahesabt41.Where(p => p.FK_basteh == tbmarahesabt4.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = findcontrol?.day2 ?? 1;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }
                                else
                                {
                                    //var findcontrol = tbmarahesabt4.Where(p => p.FK_basteh == t.FK_basteh && p.sabt == true).OrderByDescending(s => s.number).FirstOrDefault();
                                    int day2 = day3;
                                    int additionalDays = ty.mohlat ?? 1;

                                    // محاسبه مجموع روزها
                                    day3 = day2 + additionalDays;

                                    // تقویم شمسی برای بررسی تعداد روزهای ماه

                                    while (day3 > persianCalendar22.GetDaysInMonth(yer22, month22))
                                    {
                                        // تعداد روزهای ماه جاری
                                        int daysInCurrentMonth = persianCalendar22.GetDaysInMonth(yer22, month22);

                                        // کاهش تعداد روزهای ماه جاری از مجموع روزها
                                        day3 -= daysInCurrentMonth;

                                        // به ماه بعدی برویم
                                        month22++;

                                        // اگر ماه از 12 گذشت، به سال بعد برویم
                                        if (month22 > 12)
                                        {
                                            month22 = 1;
                                            yer22++;
                                        }
                                    }

                                }

                                // محاسبه تاریخ نهایی
                            }

                            date1223 = persianCalendar22.ToDateTime(yer22, month22, day3, 0, 0, 0, 0);

                            // پیدا کردن آخرین رکورد

                            DateTime now = DateTime.Now;
                            if (x != null)
                            {
                                //                                var finddd11 = tbSavedFunctions.Where(p => p.FK_tbmarahesabt4 == tbmarahesabt4.ID && p.FK_Basteh == x.FK_Basteh && p.del != true && p.svdfunc_UserSaveID == tbmarahesabt4.FK_usr
                                //&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
                                //.FirstOrDefault();
                                if (x != null)
                                {
                                    now = (DateTime)x.Datatmie;


                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسیv



                                    // محاسبه تعداد روزهای ماه
                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                                }

                            }
                            //                            if (fd1111 != null)
                            //                            {
                            //                                var finddd11 = tbSavedFunctions
                            //.Where(p => p.FK_tbmarahesabt4 == t.ID && p.FK_Basteh == fd1111.ID && p.del != true && p.svdfunc_UserSaveID == t.FK_usr
                            //&& p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == month && s.MoalfeVal_Year == year) && p.submitt == true)
                            //.FirstOrDefault();
                            //                                if (finddd11 != null)
                            //                                {
                            //                                    now = (DateTime)finddd11.svdfunc_SavedDateTime;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //                                }

                            //                            }
                            //                            else if (fd112345 != null)
                            //                            {
                            //                                var tbmoalfefishexcel11 = tbmoalfefishexcel.Where(p => p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year) && p.FK_Basteh1 == fd112345.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //                                if (tbmoalfefishexcel11 != null)
                            //                                {
                            //                                    now = (DateTime)tbmoalfefishexcel11.Datatmie;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);

                            //                                }
                            //                            }
                            //                            else if (fd21 != null)
                            //                            {
                            //                                var tbEquipmentMoalefeValueReffrenceSave11 = tbEquipmentMoalefeValueReffrenceSave.Where(p => p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year) && p.FK_Basteh == fd21.ID && p.submitt == true && p.FK_tbmarahesabt4 == t.ID && p.del != true).FirstOrDefault();
                            //                                if (tbEquipmentMoalefeValueReffrenceSave11 != null)
                            //                                {
                            //                                    now = (DateTime)tbEquipmentMoalefeValueReffrenceSave11.DateTime;


                            //                                    // استفاده از PersianCalendar برای تبدیل تاریخ میلادی به شمسی



                            //                                    // محاسبه تعداد روزهای ماه
                            //                                    //int daysInMonth = GetDaysInPersianMonth(persianYear, persianMonth);
                            //                                }
                            //                            }
                            int nowYear = persianCalendar22.GetYear(now);
                            int nowMonth = persianCalendar22.GetMonth(now);
                            int nowDay = persianCalendar22.GetDayOfMonth(now);

                            // ساختن تاریخ کامل فعلی شمسی
                            DateTime nowPersianDate = persianCalendar22.ToDateTime(nowYear, nowMonth, nowDay, 0, 0, 0, 0);
                            TimeSpan difference = date1223 - nowPersianDate;

                            // ذخیره تعداد روزها در dayasli
                            dayasli = difference.Days;
                            if (dayasli < 0)
                            {
                                DAY1 = (difference.Days.ToString());

                            }
                            else
                            {
                                DAY1 = difference.Days.ToString();

                            }
                        }


                        var x1 = await db.tbmoalfefishexcel
              .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del != true && p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                     && p.usr_sabt == ytu.FK_usr && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year))
              .FirstOrDefaultAsync();
                        var x2 = await db.tbmoalfefishexcel
           .Where(p => p.FK_Basteh1 == tbmarahesabt4.tbmarahlsabtbastedit2.FK_namebAST && p.del == true && p.Fk_Pymn == tbmarahesabt4.tbmarahlsabtbastedit2.FK_pymn && p.city == tbmarahesabt4.tbmarahlsabtbastedit2.FK_City
                  && p.usr_sabt == ytu.FK_usr && p.tbMoalefeValueFish.Any(s => s.mlfvlfsh_Month == month && s.mlfvlfsh_Year == year))
           .FirstOrDefaultAsync(); var persianCalendar = new System.Globalization.PersianCalendar();


                        if (x != null)
                        {
                            var savedDate = (DateTime)x.Datatmie;

                            // تبدیل تاریخ به شمسی
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                vazrt = 2,
                                vaziat = "موفق",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                tbmarahesabt4 = x.tbmarahesabt4,
                                tbmoalfefishexcel = x,
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }

                        else if (x1 != null)
                        {
                            var savedDate = (DateTime)x1.Datatmie;
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                vazrt = 2,
                                vaziat = "در انتظار ارسال",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                tbmarahesabt4 = x1.tbmarahesabt4,
                                tbmoalfefishexcel = x1,
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }
                        else if (x2 != null)
                        {
                            var savedDate = (DateTime)x2.Datatmie;
                            string persianDate = $"{persianCalendar.GetYear(savedDate)}/{persianCalendar.GetMonth(savedDate)}/{persianCalendar.GetDayOfMonth(savedDate)}";

                            var tbdetail = new tbdetail
                            {
                                vazrt = 2,
                                vaziat = "اصلاح برگشت",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                tbmarahesabt4 = x2.tbmarahesabt4,
                                tbmoalfefishexcel = x2,
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = persianDate // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }
                        else
                        {
                            var tbdetail = new tbdetail
                            {
                                vazrt = 2,
                                vaziat = "در انتظار ثبت",
                                fullname = ytu.tbUsers.FullName,
                                personalID = ytu.tbUsers.usr_Personal_ID.ToString(),
                                tbmarahesabt4 = ytu,
                                //tbmoalfefishexcel = x,
                                day = 1,
                                month = month,
                                year = year,
                                daysabt = "" // استفاده از تاریخ شمسی به صورت رشته

                            };
                            lstMoalefeexcel2.tbdetail.Add(tbdetail);
                        }
                    }
                }

            }

            return View("~/Areas/Contracts/Views/RegistrationAndConfirmationProcedures/viewdetail.cshtml", lstMoalefeexcel2);
        }

        public ActionResult detailkolshomarehTEST(string name = "")
        {
            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            var findResult = db.dbAsnadcratname.Where(s => s.codmahsol == name).FirstOrDefault();
            int count = 0;
            var currentItem = findResult; // شروع از آیتم فعلی

            bool isFirst = true; // متغیری برای پیگیری اینکه آیا اولین رکورد است یا خیر
            if (currentItem.IDre == null)
            {
                if (isFirst)
                {
                    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list);
                    isFirst = false;
                }
                else
                {
                    var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                    foreach (var item in list11)
                    {
                        currentItem = db.dbAsnadcratname.Where(p => p.ID == item.ID).FirstOrDefault();

                        if (currentItem != null)
                        {
                            if (isFirst)
                            {
                                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list);
                                //tolll += currentItem.name;

                                //tolll += "_";
                                //// در ابتدا بدون فاصله
                                isFirst = false; // بعد از اولین اضافه کردن، isFirst را false می‌کنیم
                            }
                            else
                            {
                                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list);
                                //tolll += "_" + currentItem.name; // اگر اولین رکورد نیست، فاصله اضافه می‌شود
                            }
                        }
                    }

                }
            }
            else
            {

            }







            // ادامه جستجو تا زمانی که IDre برابر null شود
            while (currentItem != null && currentItem.IDre != null)
            {
                if (isFirst)
                {
                    var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                    dbAsnafsherkatsanad.AddRange(list);
                    isFirst = false;
                }
                else
                {
                    var list11 = db.dbAsnadcratname.Where(p => p.IDre == currentItem.ID).ToList();
                    foreach (var item in list11)
                    {
                        currentItem = db.dbAsnadcratname.Where(p => p.ID == item.ID).FirstOrDefault();

                        if (currentItem != null)
                        {
                            if (isFirst)
                            {
                                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list);
                                //tolll += currentItem.name;

                                //tolll += "_";
                                //// در ابتدا بدون فاصله
                                isFirst = false; // بعد از اولین اضافه کردن، isFirst را false می‌کنیم
                            }
                            else
                            {
                                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                                dbAsnafsherkatsanad.AddRange(list);
                                //tolll += "_" + currentItem.name; // اگر اولین رکورد نیست، فاصله اضافه می‌شود
                            }
                        }
                    }

                }


                // اگر رکورد پیدا شد، آن را به tolll اضافه کنید

            }
            if (isFirst)
            {
                var list = db.dbAsnafsherkatsanad.Where(s => s.FK_name == currentItem.ID).ToList();
                dbAsnafsherkatsanad.AddRange(list);
                isFirst = false;
            }
            //    if (findResult != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(s => s.FK_name == findResult.ID).ToList();

            //}


            return View(dbAsnafsherkatsanad);

        }



        public async Task<ActionResult> ImportExcel_AddaSNAD(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = await GetDataFromExcel_aSNAD(MyExcelStream);
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

        //public static DateTime PersianDateToDateTime(string persianDate)
        //{
        //    PersianCalendar persianCalendar = new PersianCalendar();
        //    string[] parts = persianDate.Split('/');
        //    int year = int.Parse(parts[2]);
        //    int month = int.Parse(parts[1]);
        //    int day = int.Parse(parts[0]);
        //    return persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        //}

        public async Task<string> GetDataFromExcel_aSNAD(HttpPostedFileBase MyExcelStream)
        {
            int CaranMin, countDays, sanavatt, olad, mask, shoghl, ghar, CaranMax, CaranAcceptLimit, CaranCriterion, CaranStandard, CaranAstaneShoroPadash, CaranSaranePadash, CaranAstaneShoroJarime, CaranSaraneJarime, mablgh, sarmah;
            int FK_Moalefe_ID;
            int numbersoarat = 0; string shomerh = ""; int shorehsanad = 0;
            string name = ""; string timestart = ""; string timeend = ""; string CaranAyabOZahab = ""; string month4 = ""; string month2 = ""; string year2 = ""; string year22 = ""; string day22 = "";
            string oladd, maskan, gharbar, mozdd, sanavattt, COD, SHARH;
            double BESTANKAR = 0; double BEDEKAR = 0;
            var usr = await db.tbUsers.ToListAsync();
            var moalfe = await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<string> lstMoalefe = new List<string>();
            List<tbUserContracts> Gharardad = new List<tbUserContracts>();
            int usrud = 0;
            int year = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {
                        usrud = User.usr_ID;

                        var find2 = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login != false).FirstOrDefault();
                        if (find2 != null)
                        {
                            year = (int)find2.year;
                        }
                        //var find = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login == false).FirstOrDefault();
                        //if (find2 != null)
                        //{
                        //    return View();


                        //}
                        //if (find != null)
                        //{
                        //    return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        //}
                        //else
                        //{
                        //    return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        //}
                    }
                }
            }
            List<tbCaranSettings> IsFish = new List<tbCaranSettings>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, System.IO.Path.GetExtension(MyExcelStream.FileName));
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
                                for (int i = 14; i < title.Count; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString());
                                }
                                var Monthrow = workbook.Sheets[0].Rows[0];
                                //var Month = Monthrow.Cells[1];
                                //string number = "";
                                //if (Month.Value.ToString()!=null)
                                //{
                                //    number = Month.Value.ToString();
                                //}
                                //else
                                //{
                                //    return " لطفا ستون شماره سند   سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                //}
                                //
                                //var find = db.dbAsnadshomrehsanad.Where(s => s.shomarehsanad == number && s.Year == year).FirstOrDefault();
                                //if (find != null)
                                //{
                                //    sanad = find.ID;

                                //}
                                //else
                                //{
                                //    dbAsnadshomrehsanad dbAsnadshomrehsanad = new dbAsnadshomrehsanad();
                                //    dbAsnadshomrehsanad.Year = year;
                                //    dbAsnadshomrehsanad.shomarehsanad = number;
                                //    db.dbAsnadshomrehsanad.Add(dbAsnadshomrehsanad);
                                //    db.SaveChanges();
                                //    sanad = dbAsnadshomrehsanad.ID;
                                //}
                                //dbAsnafsherkatsanad dbAsnafsherkatsanad = new dbAsnafsherkatsanad();
                                List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
                                for (int i = 1; i < count; i++)
                                {


                                    var row = workbook.Sheets[0].Rows[i];


                                    int sanad = 0;

                                    var shomarehsanad = row.Cells[0];

                                    if (shomarehsanad.Value != null)
                                    {
                                        shomerh = shomarehsanad.Value != null ? shomarehsanad.Value.ToString() : "";

                                        //countDays = Name.Value;
                                    }

                                    else
                                    {
                                        return " لطفا ستون 1  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var find22 = db.dbAsnadshomrehsanad.Where(s => s.shomarehsanad == shomerh && s.Year == year).FirstOrDefault();
                                    if (find22 != null)
                                    {
                                        sanad = find22.ID;

                                    }
                                    else
                                    {
                                        dbAsnadshomrehsanad dbAsnadshomrehsanad = new dbAsnadshomrehsanad();
                                        dbAsnadshomrehsanad.Year = year;
                                        dbAsnadshomrehsanad.shomarehsanad = shomerh;
                                        db.dbAsnadshomrehsanad.Add(dbAsnadshomrehsanad);
                                        db.SaveChanges();
                                        sanad = dbAsnadshomrehsanad.ID;
                                    }


                                    var Name = row.Cells[1];

                                    if (Name.Value != null)
                                    {
                                        COD = Name.Value != null ? Name.Value.ToString() : "";

                                        //countDays = Name.Value;
                                    }

                                    else
                                    {
                                        return " لطفا ستون 1  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }




                                    int idfindCOD = 0;
                                    var findCOD = db.dbAsnadcratname.Where(s => s.codmahsol == COD).FirstOrDefault();
                                    if (findCOD != null)
                                    {
                                        idfindCOD = (int)findCOD.ID;

                                    }
                                    //var year = row.Cells[2];//نوع قرارداد
                                    //if (int.TryParse(year.Value.ToString(), out int result))
                                    //{
                                    //    CaranMin = result;
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون2  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    var month = row.Cells[3];//کارفرما
                                    //if (int.TryParse(month.Value.ToString(), out int result2))
                                    //{
                                    //    CaranMax = result2; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 3  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    if (month.Value != null)
                                    {
                                        SHARH = month.Value != null ? month.Value.ToString() : "";

                                        //countDays = Name.Value;
                                    }

                                    else
                                    {
                                        return " لطفا ستون 1  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    var CaranAcceptLimit2 = row.Cells[4]; //کارفرما اصلی
                                    if (double.TryParse(CaranAcceptLimit2.Value.ToString(), out double result3))
                                    {
                                        BESTANKAR = result3;
                                    }
                                    else
                                    {
                                        return " لطفا ستون 4  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var CaranCriterion2 = row.Cells[5]; //گروه شغلی
                                    if (double.TryParse(CaranCriterion2.Value.ToString(), out double CaranCriterion22))
                                    {
                                        BEDEKAR = CaranCriterion22;
                                    }
                                    else
                                    {
                                        return " لطفا ستون 5  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    //var CaranStandard2 = row.Cells[6];//sanavat
                                    //if (int.TryParse(CaranStandard2.Value.ToString(), out int CaranStandard22))
                                    //{
                                    //    CaranStandard = CaranStandard22; // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 6  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}

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


                                    var CaranAstaneShoroPadash2 = row.Cells[2];//روز
                                    if (CaranAstaneShoroPadash2.Value != null)
                                    {
                                        timestart = CaranAstaneShoroPadash2.Value != null ? CaranAstaneShoroPadash2.Value.ToString() : "";
                                    }
                                    else
                                    {
                                        return " لطفا ستون38  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }


                                    //var CaranAstaneShoroPadash22 = row.Cells[8];//روز
                                    //if (CaranAstaneShoroPadash22.Value != null)
                                    //{
                                    //    timeend = CaranAstaneShoroPadash22.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 9  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}


                                    //var mozd = row.Cells[9];//مزد شغل
                                    //if (mozd.Value.ToString() != null)
                                    //{
                                    //    mozdd = mozd.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 10  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    //var sanavat = row.Cells[10];//سنوات
                                    //if (sanavat.Value.ToString() != null)
                                    //{
                                    //    sanavattt = sanavat.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 11  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    //var oladdd = row.Cells[11];//اولاد
                                    //if (oladdd.Value.ToString() != null)
                                    //{
                                    //    oladd = oladdd.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 12  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    //var maskans = row.Cells[12];//مسکن
                                    //if (maskans.Value.ToString() != null)
                                    //{
                                    //    maskan = maskans.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 12  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    //var gharr = row.Cells[13];//خواروبار
                                    //if (gharr.Value.ToString() != null)
                                    //{
                                    //    gharbar = gharr.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    //}
                                    //else
                                    //{
                                    //    return " لطفا ستون 14  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    //}
                                    dbAsnafsherkatsanad obj4 = new dbAsnafsherkatsanad();
                                    DateTime? usc_StartTime = null;

                                    string inputDate = timestart;
                                    //string inputDate2 = timeend;

                                    // تبدیل تاریخ به تاریخ میلادی
                                    DateTime persianDate = PersianDateToDateTime(inputDate);
                                    //DateTime persianDate2 = PersianDateToDateTime(inputDate2);

                                    // تبدیل تاریخ به فرمت مناسب برای SQL Server
                                    string sqlFormattedDate = persianDate.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    //var tt = usr.FirstOrDefault(p => p.usr_Personal_ID == countDays);
                                    // اختصاص تاریخ به usc_StartTime
                                    obj4.bestankar = BESTANKAR;
                                    obj4.bedehkar = BEDEKAR;
                                    obj4.date = persianDate;
                                    obj4.Titel = SHARH;
                                    obj4.FK_shmoareh = sanad;
                                    obj4.FK_name = idfindCOD;
                                    dbAsnafsherkatsanad.Add(obj4);
                                    //obj4.FK_UserID = tt.usr_ID;
                                    //if (CaranMin == 1)
                                    //{
                                    //    obj4.usc_TypeOfContract = true;

                                    //}
                                    //else
                                    //{
                                    //    obj4.usc_TypeOfContract = false;

                                    //}



                                    //obj4.FK_JobGroup = CaranCriterion;
                                    //obj4.usc_Jobtitle = null;
                                    //obj4.jobgroup_HagheMaskan = maskan;
                                    //obj4.jobgroup_HagheOlad = oladd;
                                    //obj4.jobgroup_KharoBar = gharbar;
                                    //obj4.jobgroup_ValueSanavat = sanavattt;
                                    //obj4.jobgroup_ValueMozdGroup = mozdd;
                                    //obj4.usc_TedadSalSanavat = CaranStandard;
                                    //obj4.usc_JobCode = 0;

                                    //int counter_ContractNumber = 0;

                                    //var list_contracts = db.tbUserContracts.ToList();
                                    //if (list_contracts.Count() != 0)
                                    //{
                                    //    var lastnumber = list_contracts.LastOrDefault().usc_ContractNumber;

                                    //    if (lastnumber != null)
                                    //    {
                                    //        counter_ContractNumber = System.Convert.ToInt32(lastnumber.Split('_').Last());
                                    //    }
                                    //}
                                    //counter_ContractNumber++;

                                    //var userpersonnel = tt.usr_Personal_ID;
                                    //obj4.usc_ContractNumber = SaabWebProject.Models.Utilitis.ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now) + "_" + userpersonnel + "_" + counter_ContractNumber;

                                    int shomarande = 14;
                                    List<tbUserContractsAndMoalefeGhararDadi> objectsToAdd = new List<tbUserContractsAndMoalefeGhararDadi>();

                                    //foreach (var it in lstMoalefe)
                                    //{
                                    //    tbUserContractsAndMoalefeGhararDadi obj5 = new tbUserContractsAndMoalefeGhararDadi();
                                    //    var id = moalfe.Where(p => p.md_Title.Contains(it)).FirstOrDefault();
                                    //    var Value = row.Cells[shomarande];
                                    //    var value222 = Value.Value ?? null;
                                    //    int? countDays2;
                                    //    if (int.TryParse(value222.ToString(), out int parsedValue3))
                                    //    {
                                    //        countDays2 = parsedValue3;
                                    //    }
                                    //    else
                                    //    {
                                    //        countDays2 = 0; // Or handle the case where the value cannot be parsed
                                    //    }
                                    //    obj5.FKContractID = obj4.usc_ID;
                                    //    obj5.FKMoalefeGhararDadi = id.md_ID;
                                    //    obj5.Value = countDays2;
                                    //    objectsToAdd.Add(obj5);
                                    //    shomarande++;

                                    //}

                                    //db.tbUserContractsAndMoalefeGhararDadi.AddRange(objectsToAdd);
                                    //await db.SaveChangesAsync();

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
                                db.dbAsnafsherkatsanad.AddRange(dbAsnafsherkatsanad);
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





        public ActionResult viewforgozarsh()
        {
            return View();
        }

        public class DataModeldeatail
        {
            public string Bind { get; set; }
            public string Container { get; set; }
            public string Name { get; set; }
        }
        public ActionResult viewasnadmalyy()
        {
            var usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {
                        usr = User.usr_ID;

                        var find2 = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login != false).FirstOrDefault();

                        var find = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login == false).FirstOrDefault();
                        if (find2 != null)
                        {
                            return View();


                        }
                        if (find != null)
                        {
                            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        }
                        else
                        {
                            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        }
                    }
                }
            }
            return View();
        }
        public ActionResult viewasnadmalyylistid(int id)
        {
            var find = db.dbAsnadcratasnad.Where(p => p.ID == id).ToList();
            return View(find);
        }
        public ActionResult viewasnadmalyy2()
        {
            return View();
        }
        public ActionResult viewasnadmalyy22()
        {
            return View();
        }
        public ActionResult viewmodalsanadmallyy()
        {
            return View();
        }
        public ActionResult salmahi()
        {

            return View();

        }
        public ActionResult namebast()
        {
            return View();
        }
        public ActionResult namebastmarhe2()
        {
            return View();
        }
        public ActionResult creatviewsalmahi(int year = 0)
        {
            int usrr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {
                        usrr = User.usr_ID;
                        var find22 = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID).FirstOrDefault();
                        if (find22 != null)
                        {
                            //find22.login = false;
                            //db.SaveChanges();
                        }
                    }
                }
            }
            var find = db.dbAsnadsalvorodi.Where(s => s.year != null && s.FK_usr == usrr).FirstOrDefault();
            if (find != null)
            {
                find.FK_usr = usrr;
                find.year = year;
                find.login = true;

                db.SaveChanges();
            }
            else
            {
                dbAsnadsalvorodi dbAsnadsalvorodi = new dbAsnadsalvorodi();
                dbAsnadsalvorodi.year = year;
                dbAsnadsalvorodi.FK_usr = usrr;
                dbAsnadsalvorodi.login = true;

                db.dbAsnadsalvorodi.Add(dbAsnadsalvorodi);

                db.SaveChanges();
            }
            return Content("True");
        }

        public async Task< string> DASTRESSI(List<dbsathdastressi4> modelList)
        {
            List<dbsathdastressi4> dbsathdastressi4 = new List<dbsathdastressi4>();
            using (var context = new SaabEntities())
            {
                var find34 = context.dbsathdastressi4.ToList();

                foreach (var it in modelList)
                {
                    var find = find34.Where(p => p.FK_usr == it.FK_usr && p.Active !=false&&p.vaziatt!=it.vaziatt&&p.FK_dbsathdastressi3==it.FK_dbsathdastressi3).FirstOrDefault();
                    if (find != null)
                    {
                        
                        find.vaziatt = it.vaziatt;
                    }
                    else
                    {
                        if (it.vaziatt != false)
                        {
                            context.dbsathdastressi4.Add(it);

                        }
                    }
                }
              await  context.SaveChangesAsync();


                return "true";
            }
        }


        public ActionResult sathdastresi()
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.sathdastresi = new List<sathdastresi>();
            var findusr = db.tbUsers.Where(s => s.usr_Personal_ID != null).ToList();
            var listpymn = db.Link_User_And_Peyman.Where(s => s.Status == true).ToList();
            foreach(var it in findusr)
            {
                var findpymn = listpymn.Where(s => s.FK_User_ID == it.usr_ID && s.Status == true).FirstOrDefault();

                foreach (var it1 in db.dbsathdastressi3.ToList())
                {
                    if (it.tbjob != null)
                    {
                        if (findpymn != null)
                        {
                            var sathdastresi1 = new sathdastresi
                            {
                                IDyusr = it.usr_ID,
                                IDPymn = (int)findpymn.FK_Peyman_ID,
                                namepymn = findpymn.tbPeymanContracts.pec_Title,
                                fullname = it.FullName,
                                codpersanly = (int)it.usr_Personal_ID,
                                job = it.tbjob.Name,
                                tab = it1.dbsathdastressi2.dbsathdastressi1.name,
                                zarmagmetab = it1.name,
                                tabint = (int)it1.dbsathdastressi2.FK_dbsathdastressi1,
                                zarmagmetabint = (int)it1.dbsathdastressi2.ID,
                                zarmazaergmetabint = (int)it1.ID,
                                zarzermagmetab = it1.name,

                            };
                            lstMoalefeexcel2.sathdastresi.Add(sathdastresi1);
                        }
                    }
              
                }
          
               

            }
            return View(lstMoalefeexcel2);
        }

        public ActionResult viewmodalsanadmallyybbynumber(string id = "", DateTime? tarighh = null)
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            lstMoalefeexcel2.id = new id(); // Initialize the id property
            PersianCalendar persianCalendar = new PersianCalendar();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;

                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                //var find = db.dbRelatinafrad
                //             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                //             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                //             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                //             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                //             .ToList();

            }
            // تنظیم مقدار سال پیش‌فرض
            int year = 1403;
            var findyear = db.dbAsnadsalvorodi.Where(s => s.year != null && s.FK_usr == personal && s.login == true).FirstOrDefault();
            if (findyear != null)
            {
                year = (int)findyear.year;
            }

            // ایجاد تاریخ 1-1-سال موردنظر
            DateTime firstDayOfYear = persianCalendar.ToDateTime(year, 1, 1, 0, 0, 0, 0);

            // بازیابی اطلاعات از پایگاه داده
            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            var finddd = db.dbAsnadshomrehsanad
                .Where(p => p.shomarehsanad == id && p.Year == year)
                .FirstOrDefault();

            if (finddd != null)
            {
                dbAsnafsherkatsanad = db.dbAsnafsherkatsanad
                    .Where(p => p.FK_shmoareh == finddd.ID)
                    .ToList();

                lstMoalefeexcel2.dbAsnafsherkatsanad.AddRange(dbAsnafsherkatsanad);
            }

            // مقداردهی به id
            lstMoalefeexcel2.id.id2 = id;
            lstMoalefeexcel2.id.month = firstDayOfYear; // تنظیم تاریخ به 1-1-سال
            lstMoalefeexcel2.id.MoalfeVal_Year = year;

            return View(lstMoalefeexcel2);
        }
        public ActionResult viewmodalsanadmallyybbynumberedit(int id = 0)
        {
            var findd = db.dbAsnadshomrehsanad.Where(p => p.ID == id).FirstOrDefault();

            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            lstMoalefeexcel2.id = new id(); // Initialize the id property
            PersianCalendar persianCalendar = new PersianCalendar();

            // تنظیم مقدار سال پیش‌فرض
            int year = 1403;
            var findyear = db.dbAsnadsalvorodi.Where(s => s.year != null).FirstOrDefault();
            if (findyear != null)
            {
                year = (int)findyear.year;
            }
            year = (int)findd.Year;
            // ایجاد تاریخ 1-1-سال موردنظر
            DateTime firstDayOfYear = persianCalendar.ToDateTime(year, 1, 1, 0, 0, 0, 0);

            // بازیابی اطلاعات از پایگاه داده
            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            var finddd = db.dbAsnadshomrehsanad
                .Where(p => p.shomarehsanad == findd.shomarehsanad && p.Year == findd.Year)
                .FirstOrDefault();

            if (finddd != null)
            {
                dbAsnafsherkatsanad = db.dbAsnafsherkatsanad
                    .Where(p => p.FK_shmoareh == finddd.ID)
                    .ToList();

                lstMoalefeexcel2.dbAsnafsherkatsanad.AddRange(dbAsnafsherkatsanad);
            }

            // مقداردهی به id
            lstMoalefeexcel2.id.id2 = findd.shomarehsanad;
            lstMoalefeexcel2.id.month = firstDayOfYear; // تنظیم تاریخ به 1-1-سال
            lstMoalefeexcel2.id.MoalfeVal_Year = year;
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/viewmodalsanadmallyybbynumber.cshtml", lstMoalefeexcel2);

            //return View(lstMoalefeexcel2);
        }

        public ActionResult listasnad()
        {
            int usrud = 0;
            int year = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {
                        usrud = User.usr_ID;

                        var find2 = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login != false).FirstOrDefault();
                        if (find2 != null)
                        {
                            year = (int)find2.year;
                        }
                        //var find = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login == false).FirstOrDefault();
                        //if (find2 != null)
                        //{
                        //    return View();


                        //}
                        //if (find != null)
                        //{
                        //    return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        //}
                        //else
                        //{
                        //    return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        //}
                    }
                }
            }
            return View(db.dbAsnadshomrehsanad.Where(s=>s.Year== year).ToList());
        }
        public ActionResult editnew(int id = 0, string shomarehsanad = "")
        {
            var findyr = db.dbAsnafsherkatsanad.Where(p => p.ID == id).FirstOrDefault();
            if (findyr != null)
            {
                var find = db.dbAsnadshomrehsanad.Where(s => s.shomarehsanad == shomarehsanad && s.Year == findyr.dbAsnadshomrehsanad.Year).FirstOrDefault();
                if (find != null)
                {
                    findyr.FK_shmoareh = find.ID;
                    db.SaveChanges();
                }
                else {

                    dbAsnadshomrehsanad dbAsnadshomrehsanad = new dbAsnadshomrehsanad();
                    dbAsnadshomrehsanad.shomarehsanad = shomarehsanad;
                    dbAsnadshomrehsanad.Year = findyr.dbAsnadshomrehsanad.Year;
                    db.dbAsnadshomrehsanad.Add(dbAsnadshomrehsanad);
                    db.SaveChanges();
                    findyr.FK_shmoareh = dbAsnadshomrehsanad.ID;
                    db.SaveChanges();
                }

            }
            return Content("True");

        }
        public ActionResult TestLink()
        {
            // ایجاد نمونه از FunctionModel
            FunctionModel functionModel = new FunctionModel()
            {
                id = new id(), // مقدار پیش‌فرض
                year = new year(), // مقدار پیش‌فرض
                number = new number(), // مقدار پیش‌فرض
                dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>(), // لیست خالی
                UsersMoalefe = new List<MoalefeUserInfo>()
            };

            // ایجاد 20 کاربر با مؤلفه‌های پویا
            for (int i = 1; i <= 20; i++)
            {
                var user = new MoalefeUserInfo()
                {
                    FullName = $"User {i}",
                    PersonalCode = 100 + i,
                    listMoalefe = new List<MoalefeInfo>()
                };

                // اضافه کردن 5 مؤلفه به هر کاربر
                for (int j = 1; j <= 25; j++)
                {
                    user.listMoalefe.Add(new MoalefeInfo()
                    {
                        MoalefeTitle = $"Moalefe {j}",
                        MoalefeValue = $"{j * i}%",
                        Moalefenum = j
                    });
                }

                functionModel.UsersMoalefe.Add(user);
            }

            return View(functionModel);
        }
        public ActionResult BackupDatabasetst()
        {
            try
            {
                // مسیر ذخیره‌سازی در وب سرور
                string webBackupPath = HttpContext.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/");

                // مسیر ذخیره‌سازی در درایو D
                string localBackupPath = @"D:\تست\";

                // دریافت کانکشن استرینگ صحیح
                string connectionString = ConfigurationManager.ConnectionStrings["SaabDatabaseConnection"].ConnectionString;

                // استخراج نام دیتابیس از کانکشن استرینگ
                var builder = new SqlConnectionStringBuilder(connectionString);
                string databaseName = builder.InitialCatalog; // اینجا نام دیتابیس پویا استخراج می‌شود

                // نام فایل بک‌آپ
                string backupFileName = $"{databaseName}_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.bak";

                // مسیر کامل ذخیره‌سازی
                string fullWebBackupPath = System.IO.Path.Combine(webBackupPath, backupFileName);
                string fullLocalBackupPath = System.IO.Path.Combine(localBackupPath, backupFileName);

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // اجرای دستور بک‌آپ با نام دیتابیس پویا
                    string backupQuery = $@"
                BACKUP DATABASE [{databaseName}] TO DISK = '{fullWebBackupPath}'
                BACKUP DATABASE [{databaseName}] TO DISK = '{fullLocalBackupPath}'";

                    using (SqlCommand cmd = new SqlCommand(backupQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }

                return Json(new
                {
                    success = true,
                    message = $"بک‌آپ دیتابیس {databaseName} با موفقیت در هر دو مسیر ذخیره شد!",
                    webPath = fullWebBackupPath,
                    localPath = fullLocalBackupPath
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "خطا در گرفتن بک‌آپ: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult RestoreLatestBackup()
        {
            try
            {
                // 1️⃣ پیدا کردن آخرین فایل بک‌آپ
                string backupPath = HttpContext.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/");
                var directory = new DirectoryInfo(backupPath);
                var latestBackup = directory.GetFiles("*.bak")
                                            .OrderByDescending(f => f.LastWriteTime)
                                            .FirstOrDefault();

                if (latestBackup == null)
                {
                    return Json(new { success = false, message = "هیچ فایل بک‌آپی یافت نشد!" }, JsonRequestBehavior.AllowGet);
                }

                string latestBackupPath = latestBackup.FullName;

                // 2️⃣ دریافت نام دیتابیس از کانکشن استرینگ
                string connectionString = ConfigurationManager.ConnectionStrings["SaabDatabaseConnection"].ConnectionString;
                var builder = new SqlConnectionStringBuilder(connectionString);
                string databaseName = builder.InitialCatalog;

                // 3️⃣ بررسی و ایجاد مسیر ذخیره‌سازی در درایو D
                string dataDirectory = @"D:\SQLData";
                if (!Directory.Exists(dataDirectory))
                {
                    Directory.CreateDirectory(dataDirectory);
                }

                string masterConnectionString = connectionString.Replace(databaseName, "master");

                using (SqlConnection conn = new SqlConnection(masterConnectionString))
                {
                    conn.Open();

                    // 4️⃣ دریافت نام‌های لاجیکال فایل‌ها از بک‌آپ
                    string fileListQuery = $"RESTORE FILELISTONLY FROM DISK = '{latestBackupPath}'";
                    string logicalDataFileName = "";
                    string logicalLogFileName = "";

                    using (SqlCommand cmd = new SqlCommand(fileListQuery, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string logicalName = reader["LogicalName"].ToString();
                            string type = reader["Type"].ToString();

                            if (type == "D") // Data file
                                logicalDataFileName = logicalName;
                            else if (type == "L") // Log file
                                logicalLogFileName = logicalName;
                        }
                    }

                    if (string.IsNullOrEmpty(logicalDataFileName) || string.IsNullOrEmpty(logicalLogFileName))
                    {
                        return Json(new { success = false, message = "نام‌های لاجیکال فایل‌ها در بک‌آپ یافت نشد!" }, JsonRequestBehavior.AllowGet);
                    }

                    string dataFilePath = $@"{dataDirectory}\{databaseName}.mdf";
                    string logFilePath = $@"{dataDirectory}\{databaseName}_log.ldf";

                    // 5️⃣ بستن همه اتصالات فعال به دیتابیس
                    string killConnectionsQuery = $@"
                DECLARE @killCommand NVARCHAR(MAX) = '';
                SELECT @killCommand = @killCommand + 'KILL ' + CAST(session_id AS NVARCHAR(5)) + ';'
                FROM sys.dm_exec_sessions
                WHERE database_id = DB_ID('{databaseName}') AND session_id <> @@SPID;
                EXEC(@killCommand);";

                    using (SqlCommand cmd = new SqlCommand(killConnectionsQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 6️⃣ قرار دادن دیتابیس در حالت SINGLE_USER
                    string setSingleUserQuery = $"ALTER DATABASE [{databaseName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE";
                    using (SqlCommand cmd = new SqlCommand(setSingleUserQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 7️⃣ اجرای دستور RESTORE DATABASE با نام‌های لاجیکال صحیح
                    string restoreQuery = $@"
                RESTORE DATABASE [{databaseName}]
                FROM DISK = '{latestBackupPath}'
                WITH REPLACE, RECOVERY,
                MOVE '{logicalDataFileName}' TO '{dataFilePath}',
                MOVE '{logicalLogFileName}' TO '{logFilePath}'";

                    using (SqlCommand cmd = new SqlCommand(restoreQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 8️⃣ بازگرداندن دیتابیس به حالت MULTI_USER
                    string setMultiUserQuery = $"ALTER DATABASE [{databaseName}] SET MULTI_USER";
                    using (SqlCommand cmd = new SqlCommand(setMultiUserQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }

                return Json(new { success = true, message = "بازیابی دیتابیس با موفقیت انجام شد!", backupFile = latestBackupPath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "خطا در بازیابی دیتابیس: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult viewCHEzK1(string nohesab="")
        {
            dbAsnadCheack1 dbAsnadCheack1 =new  dbAsnadCheack1();
            dbAsnadCheack1.Name = nohesab;
            db.dbAsnadCheack1.Add(dbAsnadCheack1);
            db.SaveChanges();
            return Content("True");
        }
        public ActionResult createCHEzK4(string vaziat = "")
        {
            dbAsnadCheack5 dbAsnadCheack5 = new dbAsnadCheack5();
            dbAsnadCheack5.NAME = vaziat;
            db.dbAsnadCheack5.Add(dbAsnadCheack5);
            db.SaveChanges();
            return Content("True");
        }



        public ActionResult viewCHEzK2(dbAsnadCheack2 Filters) {
            db.dbAsnadCheack2.Add(Filters);
            db.SaveChanges();
            
            //dbAsnadCheack1 dbAsnadCheack1 = new dbAsnadCheack1();
            //dbAsnadCheack1.Name = nohesab;
            //db.dbAsnadCheack1.Add(dbAsnadCheack1);
            //db.SaveChanges();
            return Content("True");
        }


        public ActionResult viewCHEzK3(dbAsnadCheack3 Filters)
        {
            db.dbAsnadCheack3.Add(Filters);
            db.SaveChanges();

            //dbAsnadCheack1 dbAsnadCheack1 = new dbAsnadCheack1();
            //dbAsnadCheack1.Name = nohesab;
            //db.dbAsnadCheack1.Add(dbAsnadCheack1);
            //db.SaveChanges();
            return Content("True");
        }

        public ActionResult viewCHEzK6(dbAsnadCheack6 Filters)
        {
            db.dbAsnadCheack6.Add(Filters);
            db.SaveChanges();

            //dbAsnadCheack1 dbAsnadCheack1 = new dbAsnadCheack1();
            //dbAsnadCheack1.Name = nohesab;
            //db.dbAsnadCheack1.Add(dbAsnadCheack1);
            //db.SaveChanges();
            return Content("True");
        }

        //public ActionResult viewmodelbank(int id=0)
        //{
        //    if (id =!= 0)
        //    {
        //        var find=db.tb3
        //    }

        //}


        public ActionResult viewmodelbank(string shomarehhesab2="",string SJOMERH_seryal2="")
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.tickmomahdodat = new List<tickmomahdodat>();
            if (SJOMERH_seryal2 != "")
            {
                var find = db.dbAsnadCheack3.Where(s => s.SJOMERH_seryal == SJOMERH_seryal2).FirstOrDefault();
                if (find != null)
                {
                    lstMoalefeexcel2.namebank =find.dbAsnadCheack2.NAME;
                    lstMoalefeexcel2.IDbank =(long) find.dbAsnadCheack2.ID;
                    lstMoalefeexcel2.hesabbank = find.dbAsnadCheack2.shomarehhesab;
                    lstMoalefeexcel2.timeroz = DateTime.Now;

                }
                return View(lstMoalefeexcel2);
            }
            else {
                var find = db.dbAsnadCheack2.Where(s => s.shomarehhesab == shomarehhesab2).FirstOrDefault();
                if (find != null)
                {
                    lstMoalefeexcel2.namebank = find.NAME;
                    lstMoalefeexcel2.IDbank = (long)find.ID;
                    lstMoalefeexcel2.hesabbank = find.shomarehhesab;
                    lstMoalefeexcel2.timeroz = DateTime.Now;

                }
                return View(lstMoalefeexcel2);

            }
        }
        public string createcheack4(int FK_PYMN = 0, long FK_upload2 = 0, int day = 0, long numberpayvas = 0, List<string> fkPymn = null, List<string> fkusr = null, List<string> fkcompain = null,
         int month = 0, int FK_usr = 0, int FK_compain = 0, string Address = "", string sharhhoshdar = "", string NAME = "", int DOreh = 0, int count_dafaat = 0, int typehoshdar = -1,
         int year = 0, string Sharh = "", string vahed = "", string shomarehnameh = "", long mablaghChecked = 0, long ID = 0,long ID_ID=0,
         IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,
         string mac = "", DateTime? datepickerFromEn_usr = null, DateTime? datepickerFromEn_usrend = null, DateTime? datepickerFromEn_usrrepet = null, DateTime? datepickerFromEn_usrtamdid = null, DateTime? DataDocument = null


)
        {
            List<int> ParseAndFilterValues(string input)
            {
                List<int> result = new List<int>(); // لیستی برای ذخیره مقادیر صحیح
                if (!string.IsNullOrEmpty(input)) // بررسی اینکه ورودی خالی نباشد
                {
                    var values = input.Split(','); // جدا کردن مقادیر با استفاده از کاما
                    foreach (var value in values)
                    {
                        if (int.TryParse(value.Trim(), out int parsedValue)) // تلاش برای تبدیل به عدد صحیح
                        {
                            if (parsedValue > 0) // بررسی اینکه مقدار مثبت باشد
                            {
                                result.Add(parsedValue); // اضافه کردن به لیست
                            }
                        }
                        else
                        {
                            // در صورت نامعتبر بودن مقدار، یک Exception تولید کنید یا به‌طور مناسب لاگ کنید
                            throw new FormatException($"Invalid integer value in input: {value}");
                        }
                    }
                }
                return result;
            }
            //var miladiFrom = datepickerFromEn_usr.HasValue ? ConvertToGregorian(datepickerFromEn_usr.Value) : (DateTime?)null;
            //var miladiRepet = datepickerFromEn_usrrepet.HasValue ? ConvertToGregorian(datepickerFromEn_usrrepet.Value) : (DateTime?)null;

            List<int> fkpy = ParseAndFilterValues(fkPymn != null ? string.Join(",", fkPymn) : "");
            List<int> fkusrr = ParseAndFilterValues(fkusr != null ? string.Join(",", fkusr) : "");
            List<int> fkcompainn = ParseAndFilterValues(fkcompain != null ? string.Join(",", fkcompain) : "");
            List<dbAsnadCheack4> dbAsnadCheack4 = new List<dbAsnadCheack4>();
            var find = db.dbAsnadCheack3.Where(s => s.FK_dbAsnadCheack2 ==ID).FirstOrDefault();
            //long count = 1;
            //if (find != null)
            //{
            //    count += (long)find.ID_ID;

            //}
            var findID_ID = db.dbAsnadCheack4.Where(s => s.ID_ID != null&&s.Active!=false).OrderByDescending(t => t.ID).FirstOrDefault();
            long count = 1;
            if (findID_ID != null)
            {
                count += (long)findID_ID.ID_ID;

            }
            if (find != null)
            {
                foreach (var it in fkusrr)
                {
                    var find2 = db.tbUsers.Where(s => s.usr_ID == it).FirstOrDefault();
                    if (find2 != null)
                    {
                        dbAsnadCheack4 dbAsnadCheack41 = new dbAsnadCheack4();
                        dbAsnadCheack41.Data_time = datepickerFromEn_usr;
                        dbAsnadCheack41.Data_time2 = datepickerFromEn_usrrepet;
                        dbAsnadCheack41.FK_dbAsnadCheack3 = find.ID;
                        dbAsnadCheack41.FK_usr = it;
                        dbAsnadCheack41.sharh = Sharh;
                        dbAsnadCheack41.ID_ID = count;
                        //tbHoshdar1.NAME = NAME;
                        //tbHoshdar1.count_dafaat = count_dafaat;
                        //tbHoshdar1.FK_usr = it;
                        //tbHoshdar1.typehoshdar = typehoshdar;
                        //tbHoshdar1.sharhhoshdar = sharhhoshdar;
                        //tbHoshdar1.DOreh = DOreh;
                        //tbHoshdar1.datepickerFromEn_usrend = datepickerFromEn_usrend;
                        //tbHoshdar1.datepickerFromEn_usr = datepickerFromEn_usr;
                        //tbHoshdar1.datepickerFromEn_usrrepet = datepickerFromEn_usrrepet;
                        //tbHoshdar1.ID_ID = count;



                        dbAsnadCheack4.Add(dbAsnadCheack41);
                    }

                  
                }
                foreach (var it in fkcompainn)
                {
                    var findcomp = db.tbCompanies.Where(s => s.ID == it).FirstOrDefault();
                    if (findcomp != null)
                    {
                        dbAsnadCheack4 dbAsnadCheack41 = new dbAsnadCheack4();
                        dbAsnadCheack41.Data_time = datepickerFromEn_usr;
                        dbAsnadCheack41.Data_time2 = datepickerFromEn_usrrepet;
                        dbAsnadCheack41.FK_dbAsnadCheack3 = find.ID;
                        dbAsnadCheack41.FK_company = it;
                        dbAsnadCheack41.sharh = Sharh;
                        dbAsnadCheack41.ID_ID = count;

                        //tbHoshdar1.NAME = NAME;
                        //tbHoshdar1.count_dafaat = count_dafaat;
                        //tbHoshdar1.FK_usr = it;
                        //tbHoshdar1.typehoshdar = typehoshdar;
                        //tbHoshdar1.sharhhoshdar = sharhhoshdar;
                        //tbHoshdar1.DOreh = DOreh;
                        //tbHoshdar1.datepickerFromEn_usrend = datepickerFromEn_usrend;
                        //tbHoshdar1.datepickerFromEn_usr = datepickerFromEn_usr;
                        //tbHoshdar1.datepickerFromEn_usrrepet = datepickerFromEn_usrrepet;
                        //tbHoshdar1.ID_ID = count;



                        dbAsnadCheack4.Add(dbAsnadCheack41);
                    }
             
                }
            }

            db.dbAsnadCheack4.AddRange(dbAsnadCheack4);
            db.SaveChanges();
            return "True";
        }
        //private DateTime ConvertToGregorian(DateTime persianDate)
        //{
        //    PersianCalendar pc = new PersianCalendar();

        //    return new DateTime(
        //        pc.GetYear(persianDate),
        //        pc.GetMonth(persianDate),
        //        pc.GetDayOfMonth(persianDate),
        //        pc.GetHour(persianDate),
        //        pc.GetMinute(persianDate),
        //        pc.GetSecond(persianDate),
        //        new GregorianCalendar()
        //    );
        //}

        public ActionResult viewmoshaghasatcheak(int id = 0)
        {
            return View();
        }
     //public ActionResult sabtmadel()
     //   {
     //       var find=db.tbMoalefeValueFish.Where(s=>s.FK_Moalefe==)
     //   }
        private DateTime ConvertToGregorian(DateTime persianDate)
        {
            PersianCalendar pc = new PersianCalendar();

            return pc.ToDateTime(
                pc.GetYear(persianDate),
                pc.GetMonth(persianDate),
                pc.GetDayOfMonth(persianDate),
                pc.GetHour(persianDate),
                pc.GetMinute(persianDate),
                pc.GetSecond(persianDate),
                0 // میلی‌ثانیه لازم نیست اینجا معمولاً
            );
        }
        public ActionResult viewedit1(int id = 0)
        {
            return View();
        }
        public ActionResult listcheack5()
        {
            return View(db.dbAsnadCheack5.Where(s=>s.NAME!=null&&s.Active!=false).ToList());
        }
        public ActionResult listcheack12()
        {
            return View(db.dbAsnadCheack2.ToList());
        }
        public ActionResult listcheack16()
        {
            return View(db.dbAsnadCheack6.Where(s=>s.Active!=false).ToList());
        }
        public ActionResult listcheack13()
        {
            return View(db.dbAsnadCheack3.ToList());
        }
        public ActionResult listcheack14()
        {
            List<dbAsnadCheack4> dbAsnadCheack4 = new List<dbAsnadCheack4>();
            foreach (var it in db.dbAsnadCheack4.GroupBy(s => s.ID_ID).ToList())
            {
                var find = db.dbAsnadCheack4.Where(s => s.ID_ID == it.Key && s.Active != false).FirstOrDefault();
                if (find != null)
                {
                    dbAsnadCheack4.Add(find);

                }
            }
            return View(dbAsnadCheack4);
        }
        public ActionResult listcheack1()
        {
            return View(db.dbAsnadCheack1.Where(s=>s.Name!=null).ToList());
        }

        public ActionResult vieweditCHEzK1(string nohesab = "",int id=0)
        {
            var find = db.dbAsnadCheack1.Where(s => s.ID == id).FirstOrDefault();
            if (find != null)
            {
                find.Name = nohesab;
                db.SaveChanges();
                return Content("True");

            }
            else
            {
                return Content("False");

            }

        }
        public ActionResult viewCHEzK()
        {
            return View();
        }
        public ActionResult BackupDatabase()
        {
            try
            {
                // مسیر ذخیره‌سازی فایل بک‌آپ
                string backupPath = HttpContext.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/");
                string backupFileName = $"SaabDatabaseBackup_{DateTime.Now:yyyyMMdd_HHmmss}.bak";
                string fullBackupPath = System.IO.Path.Combine(backupPath, backupFileName);

                // گرفتن کانکشن استرینگ از Web.config (کانکشن استرینگ مخصوص SQL Server)
                string connectionString = ConfigurationManager.ConnectionStrings["SaabDatabaseConnection"].ConnectionString;

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string backupQuery = $"BACKUP DATABASE [sistan7bahman6] TO DISK = '{fullBackupPath}'";

                    using (SqlCommand cmd = new SqlCommand(backupQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }

                return Json(new { success = true, message = "بک‌آپ با موفقیت گرفته شد!", path = fullBackupPath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "خطا در گرفتن بک‌آپ: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult TreeTab4()
        {
            return View();
        }
        public ActionResult getinformatinuolaid(string id = "")
        {
            dbAsnaduplaodfil2 dbAsnaduplaodfil2 = new dbAsnaduplaodfil2();
            var findstep1 = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == id).FirstOrDefault();
            if (findstep1 != null)
            {
                var findstep2 = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == findstep1.ID).FirstOrDefault();
                if (findstep2 != null)
                {
                    return View(findstep2);
                }

            }
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/getinformatinuolaid.cshtml", dbAsnaduplaodfil2);

        }
        public ActionResult getinfdastehbandi(string id1 = "", string bindValue = "")
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            lstMoalefeexcel2.id = new id(); // Initialize the id property

            //List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            //var finddd = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad).FirstOrDefault();
            //if (finddd != null)
            //{
            //    dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(p => p.FK_shmoareh == finddd.ID).ToList();
            //    lstMoalefeexcel2.dbAsnafsherkatsanad.AddRange(dbAsnafsherkatsanad);
            //}

            lstMoalefeexcel2.id.id2 = id1;
            lstMoalefeexcel2.id.bindValue = bindValue;

            //dbAsnaduplaodfil2 dbAsnaduplaodfil2 = new dbAsnaduplaodfil2();
            //var findstep1 = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == id).FirstOrDefault();
            //if (findstep1 != null)
            //{
            //    var findstep2 = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == findstep1.ID).FirstOrDefault();
            //    if (findstep2 != null)
            //    {
            //        return View(findstep2);
            //    }

            //}
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/getinfdastehbandi.cshtml", lstMoalefeexcel2);

        }
        public class GPSModel
        {
            public double Latitude { get; set; }
            public double Longitude { get; set; }
        }
        public class tbGPSLog
        {
            [Key]
            public int ID { get; set; }
            public int FK_usr { get; set; } // شناسه کاربر
            public double Latitude { get; set; }
            public double Longitude { get; set; }
            public DateTime Timestamp { get; set; }
        }

        [HttpPost]
        public ActionResult SaveGPS(GPSModel model)
        {
            if (model != null)
            {
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

                if (cookie_user != null)
                {
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);

                    if (user != null)
                    {
                        // ذخیره موقعیت در دیتابیس
                        var gpsRecord = new tbGPSLog
                        {
                            FK_usr = user.usr_ID,
                            Latitude = model.Latitude,
                            Longitude = model.Longitude,
                            Timestamp = DateTime.Now
                        };

                        //db.tbGPSLogs.Add(gpsRecord);
                        //db.SaveChanges();

                        return Content("موقعیت با موفقیت ذخیره شد");
                    }
                }
            }
            return Content("خطا در ذخیره موقعیت");
        }
        public ActionResult deletcoding()
        {
            var fidn = db.dbAsnadcratname.Where(s => s.IDre == null).ToList();
            foreach (var t in fidn)
            {
                var find22 = db.dbAsnadcratname.Where(t1 => t1.IDre == t.ID).ToList();
                foreach (var t2 in find22)
                {
                    var find223 = db.dbAsnadcratname.Where(t1 => t1.IDre == t2.ID).ToList();
                    foreach (var t3 in find223)
                    {
                        var find2234 = db.dbAsnadcratname.Where(t1 => t1.IDre == t3.ID).ToList();
                        foreach (var t4 in find2234)
                        {
                            var find22344 = db.dbAsnadcratname.Where(t1 => t1.IDre == t4.ID).ToList();

                            foreach (var t5 in find22344)
                            {
                                var find223444 = db.dbAsnadcratname.Where(t1 => t1.IDre == t5.ID).ToList();
                                foreach (var t51 in find223444)
                                {
                                    var find2234445 = db.dbAsnadcratname.Where(t1 => t1.IDre == t51.ID).ToList();
                                    db.dbAsnadcratname.RemoveRange(find2234445);
                                    db.SaveChanges();
                                }
                                db.dbAsnadcratname.RemoveRange(find223444);
                                db.SaveChanges();
                            }
                            db.dbAsnadcratname.RemoveRange(find22344);
                            db.SaveChanges();
                        }
                        db.dbAsnadcratname.RemoveRange(find2234);
                        db.SaveChanges();
                    }
                    db.dbAsnadcratname.RemoveRange(find223);
                    db.SaveChanges();
                }
            }
            return Content("tr");
        }
        public string ebatat_Edit(List<SaabWebProject.Models.DomainModels.tbAdamAdam> things)
        {
            try
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        List<tbAdamAdam> tbAdamAdam = new List<tbAdamAdam>();


                        List<dbrelatinperson> dbrelatinperson = new List<dbrelatinperson>();
                        int fk = 0;
                        foreach (var item in things)
                        {
                            fk = item.Name ?? 0;
                            break;
                        }
                        bool cheack = true;
                        foreach (var item in things)
                        {
                            if (item.nam == -1)
                            {
                                cheack = false; break;
                            }
                            else if (item.nam == 0)
                            {
                                cheack = false; break;
                            }
                        }
                        // گرفتن لیست اولیه با FK_person مورد نظر



                        var find = db.tbAdamAdam.Where(p => p.Name == fk).ToList();

                        // پیدا کردن آیتم‌هایی که در find هستند ولی در things نیستند

                        foreach (var item in find)
                        {
                            item.Name = null;
                            db.SaveChanges();
                        }
                        if (cheack == false)
                        {
                            foreach (var item in db.tbUsers.ToList())
                            {
                                tbAdamAdam tbAdamAdam1 = new tbAdamAdam();
                                tbAdamAdam1.Name = fk;
                                tbAdamAdam1.nam = item.usr_ID;
                                tbAdamAdam1.Value = 1;
                                tbAdamAdam.Add(tbAdamAdam1);
                            }
                            things = tbAdamAdam;
                        }

                        else
                        {
                            foreach (var item in things)
                            {
                                tbAdamAdam tbAdamAdam1 = new tbAdamAdam();
                                tbAdamAdam1.Name = fk;
                                tbAdamAdam1.nam = item.nam;
                                tbAdamAdam1.Value = 1;
                                tbAdamAdam.Add(tbAdamAdam1);
                            }
                            things = tbAdamAdam;
                        }
                        db.tbAdamAdam.AddRange(tbAdamAdam);
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

        public ActionResult viewlistasnadshrkat()
        {
            return View(db.dbAsnadcratasnad.OrderByDescending(s => s.ID).ToList());
        }
        public ActionResult drop()
        {
            return View();

        }
        public ActionResult viewdrop()
        {
            return View();
        }


        public ActionResult viewdroptst()
        {
            return View();
        }

        public ActionResult viewtest2()
        {
            return View();

        }



        public ActionResult viewtest22()
        {
            return View();

        }
        public class MergedCell
        {
            public string mergedCell { get; set; }
            public int mergedCellId { get; set; }
            public List<int> mergedIds { get; set; }
            public string mergedWith { get; set; }
            public int rowspan { get; set; }
            public int colspan { get; set; }
            public string mergedText { get; set; }
        }
        [HttpGet]

        public JsonResult GetMergedData(int formId)
        {
            try
            {
                var record = db.dbForm2.FirstOrDefault(x => x.FK_Form1 == formId);
                if (record != null)
                {
                    // تبدیل مقدار MERGEE از string به JSON و ارسال آن
                    return Json(new { success = true, mergedData = JsonConvert.DeserializeObject<List<MergedCell>>(record.MERGEE) }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = false, message = "داده‌ای یافت نشد!" });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "خطا در دریافت داده‌ها", error = ex.Message });
            }
        }


        [HttpPost]
        public JsonResult SaveMergedData( dbForm2 model)
        {
            try
            {
                
                    // چک می‌کنیم که رکورد مرتبط در جدول موجود است یا نه
                    var record = db.dbForm2.FirstOrDefault(x => x.FK_Form1 == model.FK_Form1);

                    if (record != null)
                    {
                        record.MERGEE = model.MERGEE; // مقدار مرج را ذخیره کن
                    }
                    else
                    {
                        db.dbForm2.Add(new dbForm2
                        {
                            FK_Form1 = model.FK_Form1,
                            MERGEE = model.MERGEE
                        });
                    }

                    db.SaveChanges();
                

                return Json(new { success = true, message = "داده‌ها ذخیره شدند!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "خطا در ذخیره داده‌ها", error = ex.Message });
            }
        }



        public string creatfktitlfkasnad(List<dbAsnadcratasnadrelatin> Filters, IEnumerable<HttpPostedFileBase> files)
        {
            List<dbAsnadcratasnadrelatin> Flters = new List<dbAsnadcratasnadrelatin>();
            //if (files != null)
            //{
            //    foreach (var file in files)
            //    {

            //        if (file.ContentLength > 0)
            //        {
            //            var segment = file.FileName.Split('.');
            //            string file_type = segment[segment.Length - 1];
            //            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
            //            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
            //            foreach (var it in Filters)
            //            {
            //                it.File_SystemName = filename;
            //                it.File_RealName = file.FileName;

            //                Flters.Add(it);
            //            }
            //        }
            //        else
            //        {

            //        }
            //    }

            //}




            PersianCalendar persianCalendar = new PersianCalendar();


        
            db.dbAsnadcratasnadrelatin.AddRange(Filters);
            db.SaveChanges();
            return "True";
        }
        public string creatfktitlfkasnadsanad(List<dbAsnafsherkatsanad> Filters, IEnumerable<HttpPostedFileBase> files)
        {
            
            List<dbAsnadcratasnadrelatin> Flters = new List<dbAsnadcratasnadrelatin>();
            //if (files != null)
            //{
            //    foreach (var file in files)
            //    {

            //        if (file.ContentLength > 0)
            //        {
            //            var segment = file.FileName.Split('.');
            //            string file_type = segment[segment.Length - 1];
            //            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
            //            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
            //            foreach (var it in Filters)
            //            {
            //                it.File_SystemName = filename;
            //                it.File_RealName = file.FileName;

            //                Flters.Add(it);
            //            }
            //        }
            //        else
            //        {

            //        }
            //    }

            //}




            PersianCalendar persianCalendar = new PersianCalendar();



            db.dbAsnafsherkatsanad.AddRange(Filters);
            db.SaveChanges();
            return "True";
        }


        public string creatfktitlfkasnadedit(List<dbAsnadcratasnadrelatin> Filters, IEnumerable<HttpPostedFileBase> files)
        {
            List<dbAsnadcratasnadrelatin> Flters = new List<dbAsnadcratasnadrelatin>();
            //if (files != null)
            //{
            //    foreach (var file in files)
            //    {

            //        if (file.ContentLength > 0)
            //        {
            //            var segment = file.FileName.Split('.');
            //            string file_type = segment[segment.Length - 1];
            //            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
            //            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
            //            foreach (var it in Filters)
            //            {
            //                it.File_SystemName = filename;
            //                it.File_RealName = file.FileName;

            //                Flters.Add(it);
            //            }
            //        }
            //        else
            //        {

            //        }
            //    }

            //}




            PersianCalendar persianCalendar = new PersianCalendar();

            foreach(var it in Filters)
            {
                var find = db.dbAsnadcratasnadrelatin.Where(p => p.FK_asliasnad == it.FK_asliasnad).ToList();
                db.dbAsnadcratasnadrelatin.RemoveRange(find);
                db.SaveChanges();
            }

            db.dbAsnadcratasnadrelatin.AddRange(Filters);
            db.SaveChanges();
            return "True";
        }

        public string creatfkupload3(int FK_PYMN = 0, long FK_upload2 = 0, int day = 0, long numberpayvas = 0, List<string> fkPymn = null, List<string> fkusr = null, List<string> fkcompain = null,
                int month = 0, int FK_usr = 0, int FK_compain = 0, string Address = "",
                int year = 0, string Sharh = "", string vahed = "",string shomarehnameh="",long mablaghChecked=0,
                IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,
                string mac = "", DateTime? datepickerFromEn_usr = null, DateTime? datepickerFromEn_usrend = null, DateTime? datepickerFromEn_usrrepet = null, DateTime? datepickerFromEn_usrtamdid = null, DateTime? DataDocument = null


)
        {
            try
            {
               
                List<int> ParseAndFilterValues(string input)
                {
                    List<int> result = new List<int>(); // لیستی برای ذخیره مقادیر صحیح
                    if (!string.IsNullOrEmpty(input)) // بررسی اینکه ورودی خالی نباشد
                    {
                        var values = input.Split(','); // جدا کردن مقادیر با استفاده از کاما
                        foreach (var value in values)
                        {
                            if (int.TryParse(value.Trim(), out int parsedValue)) // تلاش برای تبدیل به عدد صحیح
                            {
                                if (parsedValue > 0) // بررسی اینکه مقدار مثبت باشد
                                {
                                    result.Add(parsedValue); // اضافه کردن به لیست
                                }
                            }
                            else
                            {
                                // در صورت نامعتبر بودن مقدار، یک Exception تولید کنید یا به‌طور مناسب لاگ کنید
                                throw new FormatException($"Invalid integer value in input: {value}");
                            }
                        }
                    }
                    return result;
                }

                List<int> fkpy = ParseAndFilterValues(fkPymn != null ? string.Join(",", fkPymn) : "");
                List<int> fkusrr = ParseAndFilterValues(fkusr != null ? string.Join(",", fkusr) : "");
                List<int> fkcompainn = ParseAndFilterValues(fkcompain != null ? string.Join(",", fkcompain) : "");

                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            foreach(var it in fkcompainn)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                if (!string.IsNullOrEmpty(Address))
                                {
                                    // بررسی اینکه مسیر وجود دارد یا نه
                                    if (!Directory.Exists(Address))
                                    {
                                        Directory.CreateDirectory(Address); // اگر وجود نداشت، مسیر را ایجاد کن
                                    }

                                    // ذخیره فایل در مسیر مشخص‌شده
                                    var savePath = System.IO.Path.Combine(Address, filename);
                                    file.SaveAs(savePath);
                                }
                                dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                                obj.vahed = vahed;
                                obj.Sharh = Sharh;
                                obj.year = year; obj.DAyChecked = day;
                                obj.month = month;
                                obj.Address = Address;

                                obj.svdfunc_FileNameExcel = file.FileName;
                                obj.svdfunc_FileSystemNameExcel = filename;
                                obj.timesabt = DateTime.Now;
                                obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                                obj.timerepetChecked = datepickerFromEn_usrrepet;
                                obj.timeendChecked = datepickerFromEn_usrend;
                                obj.timestartChecked = datepickerFromEn_usr;
                                obj.shomarehnameh = shomarehnameh;
                                obj.mablaghChecked = mablaghChecked;
                                obj.FK_compain = it;

                                if (FK_upload2 <= 0)
                                {
                                    obj.FK_upload2 = null;

                                }
                                else
                                {
                                    obj.FK_upload2 = FK_upload2;


                                }

                                // Add obj to the context and save changes
                                db.dbAsnaduplaodfil3.Add(obj);
                                db.SaveChanges();
                            }

                            foreach (var it in fkusrr)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                if (!string.IsNullOrEmpty(Address))
                                {
                                    // بررسی اینکه مسیر وجود دارد یا نه
                                    if (!Directory.Exists(Address))
                                    {
                                        Directory.CreateDirectory(Address); // اگر وجود نداشت، مسیر را ایجاد کن
                                    }

                                    // ذخیره فایل در مسیر مشخص‌شده
                                    var savePath = System.IO.Path.Combine(Address, filename);
                                    file.SaveAs(savePath);
                                }
                                dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                                obj.vahed = vahed;
                                obj.Sharh = Sharh;
                                obj.year = year;
                                obj.Address = Address;

                                obj.svdfunc_FileNameExcel = file.FileName;
                                obj.svdfunc_FileSystemNameExcel = filename;
                                obj.timesabt = DateTime.Now;
                                obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                                obj.timerepetChecked = datepickerFromEn_usrrepet;
                                obj.timeendChecked = datepickerFromEn_usrend;
                                obj.timestartChecked = datepickerFromEn_usr;
                                obj.shomarehnameh = shomarehnameh;
                                obj.mablaghChecked = mablaghChecked;
                                obj.FK_usr = it;
                                obj.DAyChecked = day;
                                obj.month = month;

                                if (FK_upload2 <= 0)
                                {
                                    obj.FK_upload2 = null;

                                }
                                else
                                {
                                    obj.FK_upload2 = FK_upload2;


                                }

                                // Add obj to the context and save changes
                                db.dbAsnaduplaodfil3.Add(obj);
                                db.SaveChanges();
                            }
                            foreach (var it in fkpy)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                if (!string.IsNullOrEmpty(Address))
                                {
                                    // بررسی اینکه مسیر وجود دارد یا نه
                                    if (!Directory.Exists(Address))
                                    {
                                        Directory.CreateDirectory(Address); // اگر وجود نداشت، مسیر را ایجاد کن
                                    }

                                    // ذخیره فایل در مسیر مشخص‌شده
                                    var savePath = System.IO.Path.Combine(Address, filename);
                                    file.SaveAs(savePath);
                                }
                                dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                                obj.vahed = vahed;
                                obj.Sharh = Sharh;
                                obj.year = year;
                                obj.svdfunc_FileNameExcel = file.FileName;
                                obj.svdfunc_FileSystemNameExcel = filename;
                                obj.timesabt = DateTime.Now;
                                obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                                obj.timerepetChecked = datepickerFromEn_usrrepet;
                                obj.timeendChecked = datepickerFromEn_usrend;
                                obj.timestartChecked = datepickerFromEn_usr;
                                obj.shomarehnameh = shomarehnameh;
                                obj.mablaghChecked = mablaghChecked;
                                obj.FK_PYMN = it; obj.DAyChecked = day;
                                obj.month = month;
                                obj.Address = Address;

                                if (FK_upload2 <= 0)
                                {
                                    obj.FK_upload2 = null;

                                }
                                else
                                {
                                    obj.FK_upload2 = FK_upload2;


                                }

                                // Add obj to the context and save changes
                                db.dbAsnaduplaodfil3.Add(obj);
                                db.SaveChanges();
                            }
                            if (fkpy.Count == 0)
                            {
                                if (fkusrr.Count == 0)
                                {
                                    if (fkcompainn.Count == 0)
                                    {
                                        var segment = file.FileName.Split('.');
                                        string file_type = segment[segment.Length - 1];
                                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                        if (!string.IsNullOrEmpty(Address))
                                        {
                                            // بررسی اینکه مسیر وجود دارد یا نه
                                            if (!Directory.Exists(Address))
                                            {
                                                Directory.CreateDirectory(Address); // اگر وجود نداشت، مسیر را ایجاد کن
                                            }

                                            // ذخیره فایل در مسیر مشخص‌شده
                                            var savePath = System.IO.Path.Combine(Address, filename);
                                            file.SaveAs(savePath);
                                        }
                                        dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                                        obj.vahed = vahed;
                                        obj.Sharh = Sharh;
                                        obj.year = year;
                                        obj.svdfunc_FileNameExcel = file.FileName;
                                        obj.svdfunc_FileSystemNameExcel = filename;
                                        obj.timesabt = DateTime.Now;
                                        obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                                        obj.timerepetChecked = datepickerFromEn_usrrepet;
                                        obj.timeendChecked = datepickerFromEn_usrend;
                                        obj.timestartChecked = datepickerFromEn_usr;
                                        obj.shomarehnameh = shomarehnameh;
                                        obj.mablaghChecked = mablaghChecked; obj.month = month;
                                        obj.Address = Address;

                                        if (FK_upload2 <= 0)
                                        {
                                            obj.FK_upload2 = null;

                                        }
                                        else
                                        {
                                            obj.FK_upload2 = FK_upload2;


                                        }
                                        db.dbAsnaduplaodfil3.Add(obj);
                                        db.SaveChanges();
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (var it in fkcompainn)
                    {

                        dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                        obj.vahed = vahed;
                        obj.Sharh = Sharh;
                        obj.year = year;
                        obj.month = month;
                        obj.Address = Address;

                        obj.timesabt = DateTime.Now;
                        obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                        obj.timerepetChecked = datepickerFromEn_usrrepet;
                        obj.timeendChecked = datepickerFromEn_usrend;
                        obj.timestartChecked = datepickerFromEn_usr;
                        obj.shomarehnameh = shomarehnameh;
                        obj.mablaghChecked = mablaghChecked;
                        obj.FK_compain = it;
                        obj.DAyChecked = day;



                        // Add obj to the context and save changes
                        db.dbAsnaduplaodfil3.Add(obj);
                        db.SaveChanges();
                    }

                    foreach (var it in fkusrr)
                    {

                        dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                        obj.vahed = vahed;
                        obj.Sharh = Sharh;
                        obj.year = year;
                        obj.DAyChecked = day;
                        obj.month = month;
                        obj.Address = Address;

                        obj.timesabt = DateTime.Now;
                        obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                        obj.timerepetChecked = datepickerFromEn_usrrepet;
                        obj.timeendChecked = datepickerFromEn_usrend;
                        obj.timestartChecked = datepickerFromEn_usr;
                        obj.shomarehnameh = shomarehnameh;
                        obj.mablaghChecked = mablaghChecked;
                        obj.FK_usr = it;

                        if (FK_upload2 <= 0)
                        {
                            obj.FK_upload2 = null;

                        }
                        else
                        {
                            obj.FK_upload2 = FK_upload2;


                        }

                        // Add obj to the context and save changes
                        db.dbAsnaduplaodfil3.Add(obj);
                        db.SaveChanges();
                    }
                    foreach (var it in fkpy)
                    {
                   

                        dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                        obj.vahed = vahed;
                        obj.Sharh = Sharh;
                        obj.year = year;
                        obj.DAyChecked = day;
                        obj.month = month;
                        obj.Address = Address;

                        obj.timesabt = DateTime.Now;
                        obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                        obj.timerepetChecked = datepickerFromEn_usrrepet;
                        obj.timeendChecked = datepickerFromEn_usrend;
                        obj.timestartChecked = datepickerFromEn_usr;
                        obj.shomarehnameh = shomarehnameh;
                        obj.mablaghChecked = mablaghChecked;
                        obj.FK_PYMN = it;
                        if (FK_upload2 <= 0)
                        {
                            obj.FK_upload2 = null;

                        }
                        else
                        {
                            obj.FK_upload2 = FK_upload2;


                        }

                        // Add obj to the context and save changes
                        db.dbAsnaduplaodfil3.Add(obj);
                        db.SaveChanges();
                    }
                    if (fkpy.Count == 0)
                    {
                        if (fkusrr.Count == 0)
                        {
                            if (fkcompainn.Count == 0)
                            {

                                dbAsnaduplaodfil3 obj = new dbAsnaduplaodfil3();
                                obj.vahed = vahed;
                                obj.Sharh = Sharh;
                                obj.year = year;
                                obj.DAyChecked = day;
                                obj.month = month;
                                obj.Address = Address;

                                obj.timesabt = DateTime.Now;
                                obj.timetamdidChecked = datepickerFromEn_usrtamdid;
                                obj.timerepetChecked = datepickerFromEn_usrrepet;
                                obj.timeendChecked = datepickerFromEn_usrend;
                                obj.timestartChecked = datepickerFromEn_usr;
                                obj.shomarehnameh = shomarehnameh;
                                obj.mablaghChecked = mablaghChecked;
                                if (FK_upload2 <= 0)
                                {
                                    obj.FK_upload2 = null;

                                }
                                else
                                {
                                    obj.FK_upload2 = FK_upload2;


                                }
                                db.dbAsnaduplaodfil3.Add(obj);
                                db.SaveChanges();
                            }
                        }
                    }
                  
                   
                    

                    // Add obj to the context and save changes
                 
                    return "True";

                }

                return "True";
            }
            catch (Exception ex)
            {
                return "0";
            }
        }

        public ActionResult start_dastehbandi_coding()
        {

            return View();
        }
        public ActionResult Link_dastehbandi_coding()
        {

            return View();
        }


        public ActionResult dastehbandi_coding()
        {

            return View();
        }
        public string creatasnadastandizer(List<dbAsnadDastehband2> Filters, IEnumerable<HttpPostedFileBase> files)
        {

            List<dbAsnadDastehband2> Flters = new List<dbAsnadDastehband2>();
            //if (files != null)
            //{
            //    foreach (var file in files)
            //    {

            //        if (file.ContentLength > 0)
            //        {
            //            var segment = file.FileName.Split('.');
            //            string file_type = segment[segment.Length - 1];
            //            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
            //            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
            //            foreach (var it in Filters)
            //            {
            //                it.File_SystemName = filename;
            //                it.File_RealName = file.FileName;

            //                Flters.Add(it);
            //            }
            //        }
            //        else
            //        {

            //        }
            //    }

            //}




            PersianCalendar persianCalendar = new PersianCalendar();
            var LIST2 = db.dbAsnadDastehband2.ToList();

            foreach (var it in Filters)
            {
                var LIST = LIST2.Where(s=>s.FK_dbAsnadDastehband1==it.FK_dbAsnadDastehband1).ToList();
                var find = LIST.Where(t => t.COD == it.COD).FirstOrDefault();
                if (find != null)
                {
                    return $"این کد {it.COD}تکراری است  و متعلیق به زیر محموعه ( {find.Titrl})است.";
                }

            }

            db.dbAsnadDastehband2.AddRange(Filters);
            db.SaveChanges();
            return "True";
        }
        public ActionResult adambeadamEdit(int id = 0)
        {

            var find = db.tbAdamAdam.Where(p => p.Name == id).FirstOrDefault();
            return View(find);
        }
        public ActionResult SabtZirmajmoeAdambeAdam(List<int> Fkpym = null, int usr = 0)
        {
            List<tbAdamAdam> tbAdamAdam1 = new List<tbAdamAdam>();


            foreach (var item1 in Fkpym)
            {
                if (item1 == 0 || item1 == -1)
                {
                    Fkpym = db.tbUsers.Select(s => s.usr_ID).ToList();
                    break;
                }
            }



            foreach (var it in Fkpym)
            {
                tbAdamAdam tbAdamAdam2 = new tbAdamAdam();
                tbAdamAdam2.nam = it;
                tbAdamAdam2.Name = usr;
                tbAdamAdam2.Value = 1;

                tbAdamAdam1.Add(tbAdamAdam2);

            }
            db.tbAdamAdam.AddRange(tbAdamAdam1);
            db.SaveChanges();
            return Content("True");

        }
        [HttpPost]
        public string creatrasnadcodingr(List<long> ID, string cod = "", string bindValue = "")
        {
            if (ID == null || !ID.Any())
            {
                return "هیچ آی‌دی انتخاب نشده است.";
            }

            List<dbAsnadDastehband2> Filters = new List<dbAsnadDastehband2>();
            List<dbAsnadcratname> dbAsnadcratname = new List<dbAsnadcratname>();

            var findcod = db.dbAsnadcratname.FirstOrDefault(s => s.codmahsol == bindValue);

            // پیدا کردن اولین رقم غیر صفر از سمت راست
            char[] codArray = cod.ToCharArray();
            int index = cod.Length - 1;

            while (index >= 0 && codArray[index] == '0')
            {
                index--; // حرکت به سمت چپ تا رسیدن به اولین رقم غیر صفر
            }

            if (index < 0)
            {
                return "کد معتبر نیست.";
            }

            int counter = 0;
            bool t = false; // متغیر برای کنترل شرط t

            foreach (var id in ID)
            {
                var findd = db.dbAsnadDastehband2.Where(s => s.FK_dbAsnadDastehband1 == id).ToList();

                foreach (var it in findd)
                {
                    if (findcod.ID != null)
                    {
                        var findcoding = db.dbAsnadcratname.Where(s => s.IDre == findcod.ID && s.FK_PK_dbAsnadDastehband2 == it.ID).FirstOrDefault();
                        if (findcoding == null)
                        {
                            dbAsnadcratname dbAsnadcratname1 = new dbAsnadcratname
                            {
                                name = it.Titrl,
                                IDre = findcod != null ? findcod.ID : 0,
                                FK_PK_dbAsnadDastehband2 = it.ID
                            };

                            if (!t)
                            {
                                dbAsnadcratname1.codmahsol = cod;
                                dbAsnadcratname.Add(dbAsnadcratname1);
                                t = true;
                            }
                            else
                            {
                                // **🔹 افزایش مقدار عددی بدون ایجاد مشکل در مقدار جدید**
                                int tempIndex = index;
                                int carry = 1; // همیشه مقدار جدید را یک واحد افزایش می‌دهیم

                                while (tempIndex >= 0 && carry > 0)
                                {
                                    if (char.IsDigit(codArray[tempIndex]))
                                    {
                                        int digit = (codArray[tempIndex] - '0') + carry;

                                        if (digit > 9)
                                        {
                                            codArray[tempIndex] = '0'; // اگر از 9 گذشت، صفر شود
                                            carry = 1; // انتقال به رقم بعدی
                                        }
                                        else
                                        {
                                            codArray[tempIndex] = (char)(digit + '0');
                                            carry = 0; // اگر نیاز به انتقال نبود، متوقف شود
                                        }
                                    }
                                    tempIndex--; // حرکت به سمت چپ برای بررسی رقم بعدی
                                }

                                dbAsnadcratname1.codmahsol = new string(codArray);
                                dbAsnadcratname.Add(dbAsnadcratname1);
                            }
                        }

                   
                    }
                 
                }
            }

            db.dbAsnadcratname.AddRange(dbAsnadcratname);
            db.SaveChanges();

            return "True";
        }


        public ActionResult Listdasteh(int id = 0)
        {
            var find = db.dbAsnadDastehband2.Where(p => p.FK_dbAsnadDastehband1 == id).ToList();
            return View(find);
        }




        public ActionResult creatDATABASE(int namedatabase = 0,long name=0)
        {
            var LIST = db.dbAsnadDastehband2.ToList();
            List<dbAsnadDastehband2> dbAsnadDastehband2 = new List<dbAsnadDastehband2>();
            if (namedatabase == 1)
            {
                foreach(var it in db.tbUsers.ToList())
                {
                    if (LIST.Where(S => S.FK_usr == it.usr_ID).FirstOrDefault() == null)
                    {
                        dbAsnadDastehband2 dbAsnadDastehband22 = new dbAsnadDastehband2();

                        dbAsnadDastehband22.COD = it.usr_Personal_ID;
                        dbAsnadDastehband22.FK_usr = it.usr_ID;
                        dbAsnadDastehband22.Titrl = it.FullName;
                        dbAsnadDastehband22.FK_dbAsnadDastehband1 = name;

                        dbAsnadDastehband2.Add(dbAsnadDastehband22);
                    }
             
                }
            }
            else if (namedatabase == 2)
            {
                foreach (var it in db.tbCompanies.ToList())
                {
                    if (LIST.Where(S => S.FK_HOGHOGHI == it.ID).FirstOrDefault() == null)
                    {
                        dbAsnadDastehband2 dbAsnadDastehband22 = new dbAsnadDastehband2();

                        dbAsnadDastehband22.COD = it.ExclusiveCode;
                        dbAsnadDastehband22.FK_HOGHOGHI = it.ID;
                        dbAsnadDastehband22.Titrl = it.CompanyName;
                        dbAsnadDastehband22.FK_dbAsnadDastehband1 = name;

                        dbAsnadDastehband2.Add(dbAsnadDastehband22);
                    }
                   
                }
            }
            
            db.dbAsnadDastehband2.AddRange(dbAsnadDastehband2);
            db.SaveChanges();
            return Content("True");
        }
        public ActionResult createdastehasnad(string name="")
        {
            dbAsnadDastehband1 dbAsnadDastehband1 = new dbAsnadDastehband1();
            dbAsnadDastehband1.NAME = name;
            db.dbAsnadDastehband1.Add(dbAsnadDastehband1);
            db.SaveChanges();
            return Content( "True");
        }
        public ActionResult viewmozohghasupolaod()
        {
            return View();
        }
        public string createmozoh(string titel = "")
        {
            dbAsnaduplaodfilmozoh dbAsnaduplaodfilmozoh = new dbAsnaduplaodfilmozoh();
            dbAsnaduplaodfilmozoh.TItel = titel;
            db.dbAsnaduplaodfilmozoh.Add(dbAsnaduplaodfilmozoh);
            db.SaveChanges();
            return "True";
        }
        public ActionResult listmozoh()
        {

            return View(db.dbAsnaduplaodfilmozoh.ToList());
        }
        public ActionResult listdocum()
        {
            var finddd = db.tbfkfinancial.ToList();
            try
            {
                return PartialView("~/Areas/ManagementAccounting/Views/FinancialDocuments/listdocum.cshtml", finddd);

                //return Content("OK");

            }
            catch (Exception ex)
            {
                return Content(ex.Message );
            }
        }
        public ActionResult listdocum1()
        {
            var finddd = db.tbfkfinancial.ToList();
            try
            {
                return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/listdocum.cshtml", finddd);

                //return Content("OK");

            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
        }

        public ActionResult datailfk(int id)
        {
            var tt = db.FinancialDocuments.Where(p => p.FK_final == id).ToList();
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/datailfk.cshtml", tt);
        }

        public ActionResult editfkview(int id)
        {
            var tt = db.FinancialDocuments.Where(p => p.ID == id).FirstOrDefault();
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/editfkview.cshtml", tt);
        }
        public ActionResult Deletfkfin(int ID)
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
                db.FinancialDocuments.Remove(db.FinancialDocuments.Find(ID));
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }

        }
        public string upload(int ID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        var peyman = db.tbfkfinancial.Where(p=>p.ID==ID).FirstOrDefault();
                        peyman.File_SystemName = filename;
                        peyman.File_RealName = file.FileName;
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

        public ActionResult editview(int id)
        {
            var y=db.tbfkfinancial.Where(p=>p.ID==id).FirstOrDefault();
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/editview.cshtml", y);
        }
        public string edittt(List<tbfkfinancial> tbfk)
        {
            try
            {
                foreach (var tbfk2 in tbfk)
                {
                    var t = db.tbfkfinancial.Where(p => p.ID == tbfk2.ID).FirstOrDefault();
                    t.Title = tbfk2.Title;
                   
                    t.numbershomar = tbfk2.numbershomar;
                    t.DataDocument = tbfk2.DataDocument;
                    db.SaveChanges();
                }

                return "True";
            }
            catch (Exception)
            {
                return "false";

            }
        }


        public string editt(List<FinancialDocuments> tbfk)
        {
            try {
                foreach(var tbfk2 in tbfk)
                {
                    var t = db.FinancialDocuments.Where(p => p.ID == tbfk2.ID).FirstOrDefault();
                    t.Title = tbfk2.Title;
                    if (tbfk2.Peyman_ID == null)
                    {
                        t.User_ID = tbfk2.User_ID;
                    }
                    else
                    {
                        t.Peyman_ID = tbfk2.Peyman_ID;
                    }
                    t.Creditor = tbfk2.Creditor;
                    t.Debtore = tbfk2.Debtore;
                    t.Creditlamdicatros_ID = tbfk2.Creditlamdicatros_ID;
                    db.SaveChanges();
                }
        
                return "true";
            }
            catch (Exception)
            {
                return "false";

            }
        }
        public ActionResult editfk(int id)
        {
            var tt = db.FinancialDocuments.Where(p => p.FK_final == id).ToList();
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/editfk.cshtml", tt);
        }
        public int creatfk(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0,
                        int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0,
                        int Creditor = 0, string Title = "", string DA = "",
                        IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,
                        string mac = "", DateTime? DataDocument = null


)
        {
            try
            {
                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            tbfkfinancial obj = new tbfkfinancial();
                            obj.DataDocument = DataDocument;
                            obj.Title = Title;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.numbershomar = (int)numberpayvas;
                            // Add obj to the context and save changes
                            db.tbfkfinancial.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    tbfkfinancial obj = new tbfkfinancial();
                    obj.DataDocument = DataDocument;
                    obj.Title = Title;

                    // Add obj to the context and save changes
                    db.tbfkfinancial.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public int creatforamaliatpyam(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
            int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0,
            int Creditor = 0, string Title = "", string DA = "",
            IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, int endfalse = 0,int fkrlation=0,
            string mac = "", DateTime? DataDocument = null, int import = 0


)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;

                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                //var find = db.dbRelatinafrad
                //             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                //             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                //             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                //             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                //             .ToList();

            }
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                var findd = db.dbRelatinafrad.Where(p => p.ID == fkrlation).FirstOrDefault();
                DateTime today2 = DateTime.Now;

                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            tbsherkatpyam obj = new tbsherkatpyam();
                            obj.DataDocument = today2;
                            obj.Title = Title;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.FK_isr2 = (int)personal;
                            if (personal == findd.dbrelatinperson.tbperson.FK_usr)
                            {
                                obj.FK_usr1 = findd.dbrelatinperson.FK_usr;

                            }
                            else
                            {
                                obj.FK_usr1 = findd.dbrelatinperson.tbperson.FK_usr;

                            }
                            obj.FK_RELATIONPERSON = findd.dbrelatinperson.FK_person;
                            obj.Fk_afrad = fkrlation;obj.notseee = 1;
                            // Add obj to the context and save changes
                            db.tbsherkatpyam.Add(obj);
                            db.SaveChanges();
                            return 1;

                        }
                    }
                }
                else
                {
                    tbsherkatpyam obj = new tbsherkatpyam();
                    obj.DataDocument = today2;
                    obj.Title = Title;
                    obj.FK_isr2 = (int)personal;
                    if (personal == findd.dbrelatinperson.tbperson.FK_usr)
                    {
                        obj.FK_usr1 = findd.dbrelatinperson.FK_usr;

                    }
                    else
                    {
                        obj.FK_usr1 = findd.dbrelatinperson.tbperson.FK_usr;

                    }
                    obj.FK_RELATIONPERSON = findd.dbrelatinperson.FK_person;
                    obj.Fk_afrad = fkrlation;
                    obj.notseee = 1;
                    // Add obj to the context and save changes
                    db.tbsherkatpyam.Add(obj);
                    db.SaveChanges();
                    return 1;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
          
        }




        public int creatforamaliat(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
                    int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0,
                    int Creditor = 0, string Title = "", string DA = "",
                    IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,int endfalse=0,
                    string mac = "", DateTime? DataDocument = null, int import=0


)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);

                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            tbcreatfaal obj = new tbcreatfaal();
                            obj.DataDocument = targetDate;
                            obj.Title = Title;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.numberpayvas = (int)numberpayvas;
                            obj.endfalse=endfalse;
                            obj.sharhasli = sharhasli;
                            obj.imporatant = import;
                            // Add obj to the context and save changes
                            db.tbcreatfaal.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    tbcreatfaal obj = new tbcreatfaal();
                    obj.DataDocument = targetDate;
                    obj.Title = Title;
                    obj.imporatant = import;
                    obj.endfalse = endfalse;
                    obj.sharhasli = sharhasli;
                    obj.numberpayvas = (int)numberpayvas;

                    // Add obj to the context and save changes
                    db.tbcreatfaal.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public string creatvaziat(dbvaZiat dbvaZiat)
        {
            var fi = db.dbvaZiat.Where(p => p.fk_realtinafrad == dbvaZiat.fk_realtinafrad).FirstOrDefault();
            if (fi != null)
            {
                fi.position = dbvaZiat.position;
                fi.title= dbvaZiat.title;
                db.SaveChanges();

            }
            else
            {
                db.dbvaZiat.Add(dbvaZiat);
                db.SaveChanges();
            }
          
            return "true";
        }
        public string creatvazvaselh(dbAsnadcratname dbAsnadcratname)
        {
            var fi = db.dbAsnadcratname.Where(p => p.name == dbAsnadcratname.name).FirstOrDefault();
            if (fi != null)
            {
                fi.nergh = dbAsnadcratname.nergh;
                fi.name = dbAsnadcratname.name;
                db.SaveChanges();

            }
            else
            {
                db.dbAsnadcratname.Add(dbAsnadcratname);
                db.SaveChanges();
            }

            return "True";
        }

        public string creatvazvaselhedit(dbAsnadcratname dbAsnadcratname)
        {
            var fi = db.dbAsnadcratname.Where(p => p.ID == dbAsnadcratname.ID).FirstOrDefault();
            if (fi != null)
            {
                fi.nerghnintlink = dbAsnadcratname.nerghnintlink;
                fi.name = dbAsnadcratname.name;
                fi.codmahsol = dbAsnadcratname.codmahsol;

                db.SaveChanges();

            }
            else
            {
                db.dbAsnadcratname.Add(dbAsnadcratname);
                db.SaveChanges();
            }

            return "True";
        }
        public string creatnameplace(dbAsnadcratsher dbAsnadcratsher)
        {
            var fi = db.dbAsnadcratsher.Where(p => p.nameshr == dbAsnadcratsher.nameshr).FirstOrDefault();
            if (fi != null)
            {
                fi.codeghtesadi = dbAsnadcratsher.codeghtesadi;
                fi.addres = dbAsnadcratsher.addres;
                db.SaveChanges();

            }
            else
            {
                db.dbAsnadcratsher.Add(dbAsnadcratsher);
                db.SaveChanges();
            }

            return "True";
        }


        public ActionResult viewfortarif()
        {
            var list = db.dbAsnadcratname.ToList();
            return View(list);
        }
        public ActionResult viewfortanamesher()
        {
            var list = db.dbAsnadcratsher.ToList();
            return View(list);
        }
        public ActionResult viewfortarifasnad()
        {
            return View();
        }
        public ActionResult viewfortariasnalviewname()
        {
            var list = db.dbAsnadcratsher.ToList();
            return View(list);
        }
        public ActionResult viewfortariasnalvasil()
        {
            var list = db.dbAsnadcratname.ToList();
            return View(list);
        }
        public ActionResult viewfortariasnalvasilupload()
        {
            var list = db.dbAsnaduplaodfil1.ToList();
            return View(list);
        }
        public ActionResult tbprson()
        {
            return View(db.tbperson.ToList());
        }
        public ActionResult viewforsanadnumber(DateTime? DataDocument = null, string sanad = "")
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();

            // Initialize the year property
            lstMoalefeexcel2.year = new year();

            // Persian calendar instance
            PersianCalendar persianCalendar = new PersianCalendar();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;

                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                //var find = db.dbRelatinafrad
                //             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                //             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                //             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                //             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                //             .ToList();

            }
            // Extract the Persian year from DataDocument if it has a value
            //int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
            int? persianYear = 1403;
            int year = 1403;
            var findyear = db.dbAsnadsalvorodi.Where(s => s.year != null&&s.FK_usr==personal&&s.login==true).FirstOrDefault();
            if (findyear != null)
            {
                year = (int)findyear.year;
            }
            //year = (int)findd.Year;
            // Check if persianYear and sanad are not null or empty
            if (persianYear.HasValue && !string.IsNullOrEmpty(sanad))
            {
                // Convert sanad to integer
                if (int.TryParse(sanad, out int sanadNumber))
                {
                    // Find the latest sanadmalyy for the specific year
                    var find = db.dbAsnadcratasnad
                        .Where(p => p.year == year)
                        .OrderByDescending(s => s.ID)
                        .Select(s => s.sanadmalyy)
                        .FirstOrDefault();

                    // Convert find to an integer if not null
                    if (int.TryParse(find, out int latestSanad))
                    {
                        List<string> foundNumbers = new List<string>();

                        // Search for values in the range between latestSanad + 1 and sanadNumber
                        for (int i = latestSanad + 1; i <= sanadNumber; i++)
                        {
                            // Check if the value exists in the database
                            var match = db.dbAsnadcratasnad
                                .Where(p => p.year == year && p.sanadmalyy == i.ToString())
                                .Select(s => s.sanadmalyy)
                                .FirstOrDefault();

                            // If a match is found, add it to the list
                            if (match == null)
                            {
                                foundNumbers.Add(i.ToString());
                            }
                        }

                        // Join the numbers with commas and format as "{number1, number2, ...}"
                        if (foundNumbers.Any())
                        {
                            if (lstMoalefeexcel2.number == null)
                            {
                                lstMoalefeexcel2.number = new number();
                            }

                            lstMoalefeexcel2.number.numberstr = "{" + string.Join(",", foundNumbers) + "}";
                        }
                    }




                }
            }

            // If no match was found, lstMoalefeexcel2.number will remain an empty string
            return View(lstMoalefeexcel2);
        }

        public ActionResult viewsanad(/*DateTime? DataDocument = null,*/ string sanad = "")
        {
            List<dbAsnadcratasnadrelatin> dbAsnadcratasnadrelatin = new List<dbAsnadcratasnadrelatin>();
            PersianCalendar persianCalendar = new PersianCalendar();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int personal = 0;

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                personal = user.usr_ID;

                // دریافت فعالیت‌های کاربر بر اساس شرط داده شده
                //var find = db.dbRelatinafrad
                //             .Where(p => p.dbrelatinperson.FK_usr == personal && p.FK_tbcreatfaal != null)
                //             // مرتب‌سازی: ابتدا تاریخ‌های امروز و سپس به ترتیب سایر تاریخ‌ها
                //             .OrderBy(p => p.DataDocument == DateTime.Today ? 0 : 1) // تاریخ‌های امروز اولویت 0 دارند
                //             .ThenBy(p => p.DataDocument) // سپس مرتب‌سازی به ترتیب تاریخ
                //             .ToList();

            }
            int year = 1403;
            var findyear = db.dbAsnadsalvorodi.Where(s => s.year != null && s.FK_usr == personal && s.login == true).FirstOrDefault();

            if (findyear != null)
            {
                year = (int)findyear.year;
            }
            //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
            //int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
            var find = db.dbAsnadcratasnad.Where(p => p.sanadmalyy == sanad &&p.year== year).FirstOrDefault();
            if (find != null)
            {
                dbAsnadcratasnadrelatin = db.dbAsnadcratasnadrelatin.Where(p => p.FK_asliasnad == find.ID).ToList();
            }
            return View(dbAsnadcratasnadrelatin);
        }

        public ActionResult viewsanadedit(/*DateTime? DataDocument = null,*/ int id=0)
        {
            List<dbAsnadcratasnadrelatin> dbAsnadcratasnadrelatin = new List<dbAsnadcratasnadrelatin>();
            PersianCalendar persianCalendar = new PersianCalendar();
            //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
            //int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
            var find = db.dbAsnadcratasnad.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                dbAsnadcratasnadrelatin = db.dbAsnadcratasnadrelatin.Where(p => p.FK_asliasnad == find.ID).ToList();
            }
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/viewsanadedit.cshtml", dbAsnadcratasnadrelatin);

            //return View(dbAsnadcratasnadrelatin);
        }


        public ActionResult viewsanadeditid(/*DateTime? DataDocument = null,*/ int id = 0)
        {
            List<dbAsnadcratasnadrelatin> dbAsnadcratasnadrelatin = new List<dbAsnadcratasnadrelatin>();
            PersianCalendar persianCalendar = new PersianCalendar();
            //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
            //int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
            var find = db.dbAsnadcratasnad.Where(p => p.ID == id).FirstOrDefault();
            if (find != null)
            {
                dbAsnadcratasnadrelatin = db.dbAsnadcratasnadrelatin.Where(p => p.FK_asliasnad == find.ID).ToList();
            }
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/viewsanadeditid.cshtml", dbAsnadcratasnadrelatin);

            //return View(dbAsnadcratasnadrelatin);
        }


        public ActionResult viewsanadeditidstringid(/*DateTime? DataDocument = null,*/ string sanad = "")
        {
            FunctionModel lstMoalefeexcel2 = new FunctionModel();
            lstMoalefeexcel2.dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            lstMoalefeexcel2.id = new id(); // Initialize the id property

            List<dbAsnafsherkatsanad> dbAsnafsherkatsanad = new List<dbAsnafsherkatsanad>();
            var finddd = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad).FirstOrDefault();
            if (finddd != null)
            {
                dbAsnafsherkatsanad = db.dbAsnafsherkatsanad.Where(p => p.FK_shmoareh == finddd.ID).ToList();
                lstMoalefeexcel2.dbAsnafsherkatsanad.AddRange(dbAsnafsherkatsanad);
            }

            lstMoalefeexcel2.id.id2 = sanad;
            //List<dbAsnadcratasnadrelatin> dbAsnadcratasnadrelatin = new List<dbAsnadcratasnadrelatin>();

            //var find = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad).FirstOrDefault();
            //if (find != null)
            //{
            //    dbAsnadcratasnadrelatin = db.dbAsnadcratasnadrelatin.Where(p => p.FK_asliasnad == find.ID).ToList();
            //}
            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/viewsanadeditidstringid.cshtml", lstMoalefeexcel2);

            //return View(dbAsnadcratasnadrelatin);
        }


        public long creatforamaasnadsherkatsanad(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
          int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
          int Creditor = 0, string Title = "", string DA = "", int sabt = 0,
          IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
          string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                var finddd = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad).FirstOrDefault();
                if (finddd == null)
                {
                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                dbAsnadshomrehsanad obj2 = new dbAsnadshomrehsanad();
                                //PersianCalendar persianCalendar = new PersianCalendar();

                                //dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                //obj.data = DataDocument;
                                obj2.File_SystemName = filename;
                                obj2.File_RealName = file.FileName;
                                obj2.shomarehsanad = sanad;
                                obj2.Year = persianYear;
                                //obj.FK_name = endfalse;
                                //obj.titel = sharhasli;
                                //obj.sum = 0;
                                //obj.sumstring = summm;
                                //obj.codsamanehmoadyan = codsamanehmoadyan;
                                //obj.year = persianYear;
                                //if (FK_usr == 0)
                                //{
                                //    obj.FK_usr = null;
                                //}
                                //else
                                //{
                                //    obj.FK_usr = FK_usr;

                                //}
                                //obj.sanadmalyy = sanad;
                                //obj.pardaght = pardaght2;
                                //obj.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                db.dbAsnadshomrehsanad.Add(obj2);
                                db.SaveChanges();
                                return obj2.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadshomrehsanad obj2 = new dbAsnadshomrehsanad();

                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        //obj.data = DataDocument;
                        //obj.FK_name = endfalse;
                        //obj.titel = sharhasli;
                        //obj.sum = 0;
                        //obj.sumstring = summm;
                        //if (FK_usr == 0)
                        //{
                        //    obj.FK_usr = null;
                        //}
                        //else
                        //{
                        //    obj.FK_usr = FK_usr;

                        //}
                        //obj.year = persianYear;

                        //obj.codsamanehmoadyan = codsamanehmoadyan;

                        //obj.sanadmalyy = sanad;
                        //obj.pardaght = pardaght2;
                        //obj.datausr = DataDocument2;
                        obj2.shomarehsanad = sanad;
                        obj2.Year = persianYear;

                        // Add obj to the context and save changes
                        db.dbAsnadshomrehsanad.Add(obj2);
                        db.SaveChanges();
                        return obj2.ID;

                    }
                }
                else
                {
                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                                dbAsnadshomrehsanad obj2 = new dbAsnadshomrehsanad();
                               var delete= db.dbAsnafsherkatsanad.Where(p => p.FK_shmoareh == finddd.ID).ToList();
                                db.dbAsnafsherkatsanad.RemoveRange(delete);
                                db.SaveChanges();

                                //dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                //obj.data = DataDocument;
                                finddd.File_SystemName = filename;
                                finddd.File_RealName = file.FileName;
                                finddd.shomarehsanad = sanad;
                                finddd.Year = persianYear;

                                //obj.FK_name = endfalse;
                                //obj.titel = sharhasli;
                                //obj.sum = 0;
                                //obj.sumstring = summm;
                                //obj.codsamanehmoadyan = codsamanehmoadyan;
                                //obj.year = persianYear;
                                //if (FK_usr == 0)
                                //{
                                //    obj.FK_usr = null;
                                //}
                                //else
                                //{
                                //    obj.FK_usr = FK_usr;

                                //}
                                //obj.sanadmalyy = sanad;
                                //obj.pardaght = pardaght2;
                                //obj.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                //db.dbAsnadshomrehsanad.Add(obj2);
                                db.SaveChanges();
                                return finddd.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadshomrehsanad obj2 = new dbAsnadshomrehsanad();

                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        var delete = db.dbAsnafsherkatsanad.Where(p => p.FK_shmoareh == finddd.ID).ToList();
                        db.dbAsnafsherkatsanad.RemoveRange(delete);
                        db.SaveChanges();
                        //obj.data = DataDocument;
                        //obj.FK_name = endfalse;
                        //obj.titel = sharhasli;
                        //obj.sum = 0;
                        //obj.sumstring = summm;
                        //if (FK_usr == 0)
                        //{
                        //    obj.FK_usr = null;
                        //}
                        //else
                        //{
                        //    obj.FK_usr = FK_usr;

                        //}
                        //obj.year = persianYear;

                        //obj.codsamanehmoadyan = codsamanehmoadyan;

                        //obj.sanadmalyy = sanad;
                        //obj.pardaght = pardaght2;
                        //obj.datausr = DataDocument2;
                        finddd.shomarehsanad = sanad;
                        finddd.Year = persianYear;

                        // Add obj to the context and save changes
                        //db.dbAsnadshomrehsanad.Add(obj2);
                        db.SaveChanges();
                        return finddd.ID;

                    }
                }
                //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                //int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                //if (sabt == 1)
                //{
                //    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                //    if (findd != null)
                //    {
                //        return -7;

                //    }
                //    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear &&p.sanadmalyy!=null ).OrderByDescending(s=>s.ID).FirstOrDefault();
                //    if (findd2 != null)
                //    {
                //        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                //        int numericValue;

                //        // Try to parse the string into an integer
                //        if (int.TryParse(sanadmalyy, out numericValue))
                //        {
                //            // Increment the value by 1
                //            numericValue += 1;

                //            // Convert it back to a string
                //            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                //            var sna = findd2.sanadmalyy;
                //            var add = sna + 1;
                //            if (sanad != newSanadmalyy)
                //            {
                //                return -6;

                //            }
                //            // If you need to assign it back to the original object
                //            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                //        }
                //        else
                //        {
                //            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                //            // You can log an error or set a default value as needed
                //            throw new FormatException("sanadmalyy is not a valid integer string.");
                //        }


                //    }
                //}
            

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }



        public long creatforamaasnadsherkat(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
             int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
             int Creditor = 0, string Title = "", string DA = "", int sabt = 0,
             IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
             string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                if (sabt == 1)
                {
                    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                    if (findd != null)
                    {
                        return -7;

                    }
                    var find2 = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad && p.Year == persianYear).FirstOrDefault();
                    if (find2 != null)
                    {
                        return -7;

                    }
                    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy != null).OrderByDescending(s => s.ID).FirstOrDefault();
                    if (findd2 != null)
                    {
                        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                        int numericValue;

                        // Try to parse the string into an integer
                        if (int.TryParse(sanadmalyy, out numericValue))
                        {
                            // Increment the value by 1
                            numericValue += 1;

                            // Convert it back to a string
                            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                            var sna = findd2.sanadmalyy;
                            var add = sna + 1;
                            if (sanad != newSanadmalyy)
                            {
                                return -6;

                            }
                            // If you need to assign it back to the original object
                            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                        }
                        else
                        {
                            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                            // You can log an error or set a default value as needed
                            throw new FormatException("sanadmalyy is not a valid integer string.");
                        }


                    }
                    var findd23 = db.dbAsnadshomrehsanad.Where(p => p.Year == persianYear && p.shomarehsanad != null).OrderByDescending(s => s.ID).FirstOrDefault();

                    if (findd23 != null)
                    {
                        string sanadmalyy = findd23.shomarehsanad; // This is your original string
                        int numericValue;

                        // Try to parse the string into an integer
                        if (int.TryParse(sanadmalyy, out numericValue))
                        {
                            // Increment the value by 1
                            numericValue += 1;

                            // Convert it back to a string
                            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                            var sna = findd23.shomarehsanad;
                            var add = sna + 1;
                            if (sanad != newSanadmalyy)
                            {
                                return -6;

                            }
                            // If you need to assign it back to the original object
                            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                        }
                        else
                        {
                            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                            // You can log an error or set a default value as needed
                            throw new FormatException("sanadmalyy is not a valid integer string.");
                        }


                    }


                }

                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                          
                            dbAsnadcratasnad obj = new dbAsnadcratasnad();
                            obj.data = DataDocument;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.FK_name = endfalse;
                            obj.titel = sharhasli;
                            obj.sum= 0;
                            obj.sumstring = summm;
                            obj.codsamanehmoadyan = codsamanehmoadyan;
                            obj.year = persianYear;
                            if (FK_usr == 0)
                            {
                                obj.FK_usr = null;
                            }
                            else
                            {
                                obj.FK_usr = FK_usr;

                            }
                            obj.sanadmalyy = sanad;
                            obj.pardaght = pardaght2;
                            obj.datausr = DataDocument2;
                            // Add obj to the context and save changes
                            db.dbAsnadcratasnad.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    dbAsnadcratasnad obj = new dbAsnadcratasnad();
                    obj.data = DataDocument;
                    obj.FK_name = endfalse;
                    obj.titel = sharhasli;
                    obj.sum = 0;
                    obj.sumstring = summm;
                    if (FK_usr == 0)
                    {
                        obj.FK_usr = null;
                    }
                    else
                    {
                        obj.FK_usr = FK_usr;

                    }
                    obj.year = persianYear;

                    obj.codsamanehmoadyan = codsamanehmoadyan;

                    obj.sanadmalyy = sanad;
                    obj.pardaght = pardaght2;
                    obj.datausr = DataDocument2;

                    // Add obj to the context and save changes
                    db.dbAsnadcratasnad.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public long creatforamaasnadsherkatnahei(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
             int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
             int Creditor = 0, string Title = "", string DA = "", int sabt = 0,
             IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
             string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                if (sabt == 1)
                {
                    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                    if (findd != null)
                    {
                        return -7;

                    }
                    var find2 = db.dbAsnadshomrehsanad.Where(p => p.shomarehsanad == sanad && p.Year == persianYear).FirstOrDefault();
                    if (find2 != null)
                    {
                        return -7;

                    }
                    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy != null).OrderByDescending(s => s.ID).FirstOrDefault();
                    if (findd2 != null)
                    {
                        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                        int numericValue;

                        // Try to parse the string into an integer
                        if (int.TryParse(sanadmalyy, out numericValue))
                        {
                            // Increment the value by 1
                            numericValue += 1;

                            // Convert it back to a string
                            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                            var sna = findd2.sanadmalyy;
                            var add = sna + 1;
                            if (sanad != newSanadmalyy)
                            {
                                return -6;

                            }
                            // If you need to assign it back to the original object
                            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                        }
                        else
                        {
                            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                            // You can log an error or set a default value as needed
                            throw new FormatException("sanadmalyy is not a valid integer string.");
                        }


                    }
                    var findd23 = db.dbAsnadshomrehsanad.Where(p => p.Year == persianYear && p.shomarehsanad != null).OrderByDescending(s => s.ID).FirstOrDefault();

                    if (findd23 != null)
                    {
                        string sanadmalyy = findd23.shomarehsanad; // This is your original string
                        int numericValue;

                        // Try to parse the string into an integer
                        if (int.TryParse(sanadmalyy, out numericValue))
                        {
                            // Increment the value by 1
                            numericValue += 1;

                            // Convert it back to a string
                            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                            var sna = findd23.shomarehsanad;
                            var add = sna + 1;
                            if (sanad != newSanadmalyy)
                            {
                                return -6;

                            }
                            // If you need to assign it back to the original object
                            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                        }
                        else
                        {
                            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                            // You can log an error or set a default value as needed
                            throw new FormatException("sanadmalyy is not a valid integer string.");
                        }


                    }


                }

                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            dbAsnadcratasnad obj = new dbAsnadcratasnad();
                            obj.data = DataDocument;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.FK_name = endfalse;
                            obj.titel = sharhasli;
                            obj.sum = 0;
                            obj.sumstring = summm;
                            obj.codsamanehmoadyan = codsamanehmoadyan;
                            obj.year = persianYear;
                            if (FK_usr == 0)
                            {
                                obj.FK_usr = null;
                            }
                            else
                            {
                                obj.FK_usr = FK_usr;

                            }
                            obj.taeed = true;
                            obj.sanadmalyy = sanad;
                            obj.pardaght = pardaght2;
                            obj.datausr = DataDocument2;
                            // Add obj to the context and save changes
                            db.dbAsnadcratasnad.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    dbAsnadcratasnad obj = new dbAsnadcratasnad();
                    obj.data = DataDocument;
                    obj.FK_name = endfalse;
                    obj.titel = sharhasli;
                    obj.sum = 0;
                    obj.sumstring = summm;
                    if (FK_usr == 0)
                    {
                        obj.FK_usr = null;
                    }
                    else
                    {
                        obj.FK_usr = FK_usr;

                    }
                    obj.year = persianYear;

                    obj.codsamanehmoadyan = codsamanehmoadyan;

                    obj.sanadmalyy = sanad;
                    obj.pardaght = pardaght2;
                    obj.datausr = DataDocument2;
                    obj.taeed = true;

                    // Add obj to the context and save changes
                    db.dbAsnadcratasnad.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        public long creatforamaasnadsherkatedit(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
      int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
      int Creditor = 0, string Title = "", string DA = "", int sabt = 0,int idd=0,
      IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
      string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                var finddd = db.dbAsnadcratasnad.Where(p => p.ID == idd).FirstOrDefault();
                if (finddd == null)
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                    int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                    //if (sabt == 1)
                    //{
                    //    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                    //    if (findd != null)
                    //    {
                    //        return -7;

                    //    }
                    //    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear &&p.sanadmalyy!=null ).OrderByDescending(s=>s.ID).FirstOrDefault();
                    //    if (findd2 != null)
                    //    {
                    //        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                    //        int numericValue;

                    //        // Try to parse the string into an integer
                    //        if (int.TryParse(sanadmalyy, out numericValue))
                    //        {
                    //            // Increment the value by 1
                    //            numericValue += 1;

                    //            // Convert it back to a string
                    //            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                    //            var sna = findd2.sanadmalyy;
                    //            var add = sna + 1;
                    //            if (sanad != newSanadmalyy)
                    //            {
                    //                return -6;

                    //            }
                    //            // If you need to assign it back to the original object
                    //            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                    //        }
                    //        else
                    //        {
                    //            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                    //            // You can log an error or set a default value as needed
                    //            throw new FormatException("sanadmalyy is not a valid integer string.");
                    //        }


                    //    }
                    //}
                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                                dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                obj.data = DataDocument;
                                obj.File_SystemName = filename;
                                obj.File_RealName = file.FileName;
                                obj.FK_name = endfalse;
                                obj.titel = sharhasli;
                                obj.sum = 0;
                                obj.sumstring = summm;
                                obj.codsamanehmoadyan = codsamanehmoadyan;
                                obj.year = persianYear;
                                if (FK_usr == 0)
                                {
                                    obj.FK_usr = null;
                                }
                                else
                                {
                                    obj.FK_usr = FK_usr;

                                }
                                obj.sanadmalyy = sanad;
                                obj.pardaght = pardaght2;
                                obj.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                db.dbAsnadcratasnad.Add(obj);
                                db.SaveChanges();
                                return obj.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        obj.data = DataDocument;
                        obj.FK_name = endfalse;
                        obj.titel = sharhasli;
                        obj.sum = 0;
                        obj.sumstring = summm;
                        if (FK_usr == 0)
                        {
                            obj.FK_usr = null;
                        }
                        else
                        {
                            obj.FK_usr = FK_usr;

                        }
                        obj.year = persianYear;

                        obj.codsamanehmoadyan = codsamanehmoadyan;

                        obj.sanadmalyy = sanad;
                        obj.pardaght = pardaght2;
                        obj.datausr = DataDocument2;

                        // Add obj to the context and save changes
                        db.dbAsnadcratasnad.Add(obj);
                        db.SaveChanges();
                        return obj.ID;

                    }

                    return 1;
                }
                else
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                    int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;

                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                                dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                finddd.data = DataDocument;
                                finddd.File_SystemName = filename;
                                finddd.File_RealName = file.FileName;
                                finddd.FK_name = endfalse;
                                finddd.titel = sharhasli;
                                finddd.sum = 0;
                                finddd.sumstring = summm;
                                finddd.codsamanehmoadyan = codsamanehmoadyan;
                                finddd.year = persianYear;
                                if (FK_usr == 0)
                                {
                                    finddd.FK_usr = null;
                                }
                                else
                                {
                                    finddd.FK_usr = FK_usr;

                                }
                                finddd.sanadmalyy = sanad;
                                finddd.pardaght = pardaght2;
                                finddd.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                //db.dbAsnadcratasnad.Add(obj);
                                db.SaveChanges();
                                return finddd.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        finddd.data = DataDocument;
                        finddd.FK_name = endfalse;
                        finddd.titel = sharhasli;
                        finddd.sum = 0;
                        finddd.sumstring = summm;
                        if (FK_usr == 0)
                        {
                            finddd.FK_usr = null;
                        }
                        else
                        {
                            finddd.FK_usr = FK_usr;

                        }
                        finddd.year = persianYear;

                        finddd.codsamanehmoadyan = codsamanehmoadyan;

                        finddd.sanadmalyy = sanad;
                        finddd.pardaght = pardaght2;
                        finddd.datausr = DataDocument2;

                        // Add obj to the context and save changes
                        //db.dbAsnadcratasnad.Add(obj);
                        db.SaveChanges();
                        return finddd.ID;

                    }
                    return 1;

                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public long creatforamaasnadsherkateditnahei(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
 int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
 int Creditor = 0, string Title = "", string DA = "", int sabt = 0, int idd = 0,
 IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
 string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                var finddd = db.dbAsnadcratasnad.Where(p => p.ID == idd).FirstOrDefault();
                if (finddd == null)
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                    int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                    //if (sabt == 1)
                    //{
                    //    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                    //    if (findd != null)
                    //    {
                    //        return -7;

                    //    }
                    //    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear &&p.sanadmalyy!=null ).OrderByDescending(s=>s.ID).FirstOrDefault();
                    //    if (findd2 != null)
                    //    {
                    //        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                    //        int numericValue;

                    //        // Try to parse the string into an integer
                    //        if (int.TryParse(sanadmalyy, out numericValue))
                    //        {
                    //            // Increment the value by 1
                    //            numericValue += 1;

                    //            // Convert it back to a string
                    //            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                    //            var sna = findd2.sanadmalyy;
                    //            var add = sna + 1;
                    //            if (sanad != newSanadmalyy)
                    //            {
                    //                return -6;

                    //            }
                    //            // If you need to assign it back to the original object
                    //            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                    //        }
                    //        else
                    //        {
                    //            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                    //            // You can log an error or set a default value as needed
                    //            throw new FormatException("sanadmalyy is not a valid integer string.");
                    //        }


                    //    }
                    //}
                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                                dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                obj.data = DataDocument;
                                obj.File_SystemName = filename;
                                obj.File_RealName = file.FileName;
                                obj.FK_name = endfalse;
                                obj.titel = sharhasli;
                                obj.sum = 0;
                                obj.sumstring = summm;
                                obj.codsamanehmoadyan = codsamanehmoadyan;
                                obj.year = persianYear;
                                obj.taeed = true;
                                if (FK_usr == 0)
                                {
                                    obj.FK_usr = null;
                                }
                                else
                                {
                                    obj.FK_usr = FK_usr;

                                }
                                obj.sanadmalyy = sanad;
                                obj.pardaght = pardaght2;
                                obj.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                db.dbAsnadcratasnad.Add(obj);
                                db.SaveChanges();
                                return obj.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        obj.data = DataDocument;
                        obj.FK_name = endfalse;
                        obj.titel = sharhasli;
                        obj.sum = 0;
                        obj.taeed = true;

                        obj.sumstring = summm;
                        if (FK_usr == 0)
                        {
                            obj.FK_usr = null;
                        }
                        else
                        {
                            obj.FK_usr = FK_usr;

                        }
                        obj.year = persianYear;

                        obj.codsamanehmoadyan = codsamanehmoadyan;

                        obj.sanadmalyy = sanad;
                        obj.pardaght = pardaght2;
                        obj.datausr = DataDocument2;

                        // Add obj to the context and save changes
                        db.dbAsnadcratasnad.Add(obj);
                        db.SaveChanges();
                        return obj.ID;

                    }

                    return 1;
                }
                else
                {
                    PersianCalendar persianCalendar = new PersianCalendar();
                    //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                    int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;

                    if (files != null && files.Any())
                    {
                        foreach (var file in files)
                        {
                            if (file.ContentLength > 0)
                            {
                                var segment = file.FileName.Split('.');
                                string file_type = segment[segment.Length - 1];
                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                                dbAsnadcratasnad obj = new dbAsnadcratasnad();
                                finddd.data = DataDocument;
                                finddd.File_SystemName = filename;
                                finddd.File_RealName = file.FileName;
                                finddd.FK_name = endfalse;
                                finddd.titel = sharhasli;
                                finddd.sum = 0;
                                finddd.sumstring = summm;
                                finddd.codsamanehmoadyan = codsamanehmoadyan;
                                finddd.year = persianYear;
                                finddd.taeed = true;

                                if (FK_usr == 0)
                                {
                                    finddd.FK_usr = null;
                                }
                                else
                                {
                                    finddd.FK_usr = FK_usr;

                                }
                                finddd.sanadmalyy = sanad;
                                finddd.pardaght = pardaght2;
                                finddd.datausr = DataDocument2;
                                // Add obj to the context and save changes
                                //db.dbAsnadcratasnad.Add(obj);
                                db.SaveChanges();
                                return finddd.ID;

                            }
                        }
                    }
                    else
                    {
                        dbAsnadcratasnad obj = new dbAsnadcratasnad();
                        finddd.data = DataDocument;
                        finddd.FK_name = endfalse;
                        finddd.titel = sharhasli;
                        finddd.sum = 0;
                        finddd.sumstring = summm;
                        finddd.taeed = true;

                        if (FK_usr == 0)
                        {
                            finddd.FK_usr = null;
                        }
                        else
                        {
                            finddd.FK_usr = FK_usr;

                        }
                        finddd.year = persianYear;

                        finddd.codsamanehmoadyan = codsamanehmoadyan;

                        finddd.sanadmalyy = sanad;
                        finddd.pardaght = pardaght2;
                        finddd.datausr = DataDocument2;

                        // Add obj to the context and save changes
                        //db.dbAsnadcratasnad.Add(obj);
                        db.SaveChanges();
                        return finddd.ID;

                    }
                    return 1;

                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }



        public long creatforamaasnadsherkat22(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
              int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
              int Creditor = 0, string Title = "", string DA = "", int sabt = 0,
              IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
              string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


 )
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                //if (sabt == 1)
                //{
                //    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                //    if (findd != null)
                //    {
                //        return -7;

                //    }
                //    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy != null).OrderByDescending(s => s.ID).FirstOrDefault();
                //    if (findd2 != null)
                //    {
                //        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                //        int numericValue;

                //        // Try to parse the string into an integer
                //        if (int.TryParse(sanadmalyy, out numericValue))
                //        {
                //            // Increment the value by 1
                //            numericValue += 1;

                //            // Convert it back to a string
                //            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                //            var sna = findd2.sanadmalyy;
                //            var add = sna + 1;
                //            if (sanad != newSanadmalyy)
                //            {
                //                return -6;

                //            }
                //            // If you need to assign it back to the original object
                //            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                //        }
                //        else
                //        {
                //            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                //            // You can log an error or set a default value as needed
                //            throw new FormatException("sanadmalyy is not a valid integer string.");
                //        }


                //    }
                //}
                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            dbAsnadcratasnad obj = new dbAsnadcratasnad();
                            obj.data = DataDocument;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.FK_name = endfalse;
                            obj.titel = sharhasli;
                            obj.sum = 0;
                            obj.sumstring = summm;
                            obj.codsamanehmoadyan = codsamanehmoadyan;
                            obj.year = persianYear;
                            if (FK_usr == 0)
                            {
                                obj.FK_usr = null;
                            }
                            else
                            {
                                obj.FK_usr = FK_usr;

                            }
                            obj.sanadmalyy = sanad;
                            obj.pardaght = pardaght2;
                            obj.datausr = DataDocument2;
                            // Add obj to the context and save changes
                            db.dbAsnadcratasnad.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    dbAsnadcratasnad obj = new dbAsnadcratasnad();
                    obj.data = DataDocument;
                    obj.FK_name = endfalse;
                    obj.titel = sharhasli;
                    obj.sum = 0;
                    obj.sumstring = summm;
                    if (FK_usr == 0)
                    {
                        obj.FK_usr = null;
                    }
                    else
                    {
                        obj.FK_usr = FK_usr;

                    }
                    obj.year = persianYear;

                    obj.codsamanehmoadyan = codsamanehmoadyan;

                    obj.sanadmalyy = sanad;
                    obj.pardaght = pardaght2;
                    obj.datausr = DataDocument2;

                    // Add obj to the context and save changes
                    db.dbAsnadcratasnad.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }



        public long creatforamaasnadsherkat22nahei(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0, string sharhasli = "",
            int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, long pardaght2 = 0, string codsamanehmoadyan = "",
            int Creditor = 0, string Title = "", string DA = "", int sabt = 0,
            IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, long endfalse = 0, int FK_usr = 0, DateTime? DataDocument2 = null, string sanad = "",
            string mac = "", DateTime? DataDocument = null, int import = 0, float sum = 0, string summm = ""


)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                //DateTime targetDate = new DateTime(DataDocument.Value.Year, DataDocument.Value.Month, DataDocument.Value.Day, persianCalendar);
                int? persianYear = DataDocument.HasValue ? persianCalendar.GetYear(DataDocument.Value) : (int?)null;
                //if (sabt == 1)
                //{
                //    var findd = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy == sanad).FirstOrDefault();
                //    if (findd != null)
                //    {
                //        return -7;

                //    }
                //    var findd2 = db.dbAsnadcratasnad.Where(p => p.year == persianYear && p.sanadmalyy != null).OrderByDescending(s => s.ID).FirstOrDefault();
                //    if (findd2 != null)
                //    {
                //        string sanadmalyy = findd2.sanadmalyy; // This is your original string
                //        int numericValue;

                //        // Try to parse the string into an integer
                //        if (int.TryParse(sanadmalyy, out numericValue))
                //        {
                //            // Increment the value by 1
                //            numericValue += 1;

                //            // Convert it back to a string
                //            string newSanadmalyy = numericValue.ToString(); // This will be "3"
                //            var sna = findd2.sanadmalyy;
                //            var add = sna + 1;
                //            if (sanad != newSanadmalyy)
                //            {
                //                return -6;

                //            }
                //            // If you need to assign it back to the original object
                //            //findd2.sanadmalyy = newSanadmalyy; // Update the property
                //        }
                //        else
                //        {
                //            // Handle the case where parsing fails (if sanadmalyy is not a valid number)
                //            // You can log an error or set a default value as needed
                //            throw new FormatException("sanadmalyy is not a valid integer string.");
                //        }


                //    }
                //}
                if (files != null && files.Any())
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));

                            dbAsnadcratasnad obj = new dbAsnadcratasnad();
                            obj.data = DataDocument;
                            obj.File_SystemName = filename;
                            obj.File_RealName = file.FileName;
                            obj.FK_name = endfalse;
                            obj.titel = sharhasli;
                            obj.sum = 0;
                            obj.sumstring = summm;
                            obj.codsamanehmoadyan = codsamanehmoadyan;
                            obj.year = persianYear;
                            if (FK_usr == 0)
                            {
                                obj.FK_usr = null;
                            }
                            else
                            {
                                obj.FK_usr = FK_usr;

                            }
                            obj.taeed = true;

                            obj.sanadmalyy = sanad;
                            obj.pardaght = pardaght2;
                            obj.datausr = DataDocument2;
                            // Add obj to the context and save changes
                            db.dbAsnadcratasnad.Add(obj);
                            db.SaveChanges();
                            return obj.ID;

                        }
                    }
                }
                else
                {
                    dbAsnadcratasnad obj = new dbAsnadcratasnad();
                    obj.data = DataDocument;
                    obj.FK_name = endfalse;
                    obj.titel = sharhasli;
                    obj.sum = 0;
                    obj.sumstring = summm;
                    if (FK_usr == 0)
                    {
                        obj.FK_usr = null;
                    }
                    else
                    {
                        obj.FK_usr = FK_usr;

                    }
                    obj.year = persianYear;

                    obj.codsamanehmoadyan = codsamanehmoadyan;

                    obj.sanadmalyy = sanad;
                    obj.pardaght = pardaght2;
                    obj.datausr = DataDocument2;
                    obj.taeed = true;

                    // Add obj to the context and save changes
                    db.dbAsnadcratasnad.Add(obj);
                    db.SaveChanges();
                    return obj.ID;

                }

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public ActionResult viewsanadmallysher()
        {
            var usr = 0;
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {
                        usr = User.usr_ID;
                        var find2 = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login != false).FirstOrDefault();

                        var find = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID && p.login == false).FirstOrDefault();
                        if (find2 != null)
                        {
                            return View();


                        }
                        if (find != null)
                        {
                            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        }
                        else
                        {
                            return View("~/Areas/ManagementAccounting/Views/FinancialDocuments/salmahi.cshtml");

                        }
                    }
                }
            }
            return View();
        }

        public ActionResult usr()
        {
            var fin = db.tbUsers.ToList();
            return View(fin);
        }

        public string EditFinantial(int Peyman_ID = 0, int Detailedtitle_ID = 0, int ID = 0,

     int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, int Creditor = 0, string Title = "", string DA = "",
     IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, string mac = "", DateTime DataDocument = default)
        {
            FinancialDocuments obj = rep_def_rsn.Find(ID);
            if (User_ID == 0)
            {
                obj.User_ID = null;
            }
            else
            {
                obj.User_ID = User_ID;
            }
            if (Detailedtitle_ID == 0)
            {
                obj.Detailedtitle_ID = null;
            }
            else
            {
                obj.Detailedtitle_ID = Detailedtitle_ID;
            }

            if (Creditlamdicatros_ID == 0)
            {
                obj.Creditlamdicatros_ID = null;
            }
            else
            {
                obj.Creditlamdicatros_ID = Creditlamdicatros_ID;
            }
            DateTime DataDocument2 = DateTime.Now;
            DA = DataDocument2.ToShamsi();
            obj.MacAdress = GetMacAddress();
            obj.DataDocument = DataDocument;
            obj.Rigstertime = DA;
            obj.Peyman_ID = Peyman_ID;
            obj.Debtore = Debtore;
            obj.Creditor = Creditor;
            obj.Title = Title;
            if (files != null)
            {
                foreach (var file in files)
                {
                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;
                        return rep_def_rsn.Update2(obj).ToString();
                    }
                }
                return "True";
            }
            else
            {
                return rep_def_rsn.Update2(obj).ToString();
            }
        }
        public string UploadAttachmentFile(int Peyman_ID = 0, int Detailedtitle_ID = 0, long numberpayvas = 0,

     int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, int Creditor = 0, string Title = "", string DA = "",
     IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, string mac = "", DateTime DataDocument = default)
        {


            FinancialDocuments obj = new FinancialDocuments();

            if (User_ID == 0)
            {
                obj.User_ID = null;
            }
            else
            {
                obj.User_ID = User_ID;
            }



            if (Detailedtitle_ID == 0)
            {
                obj.Detailedtitle_ID = null;
            }
            else
            {
                obj.Detailedtitle_ID = Detailedtitle_ID;
            }

            if (Creditlamdicatros_ID == 0)
            {
                obj.Creditlamdicatros_ID = null;
            }
            else
            {
                obj.Creditlamdicatros_ID = Creditlamdicatros_ID;
            }
            DateTime DataDocument2 = DateTime.Now;
            DA = DataDocument2.ToShamsi();

            obj.numberpayvas = numberpayvas;
            obj.MacAdress = GetMacAddress();

            obj.DataDocument = DataDocument;
            obj.Rigstertime = DA;
            obj.Peyman_ID = Peyman_ID;


            obj.Debtore = Debtore;
            obj.Creditor = Creditor;
            obj.Title = Title;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;

                        return rep_def_rsn.Create(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return rep_def_rsn.Create(obj).ToString();
            }
        }

        [AuthorizeAAA]
        [HttpPost]
        public string Edit(int ID = 0, int Peyman_ID = 0, int Detailedtitle_ID = 0,

     int User_ID = 0, int Creditlamdicatros_ID = 0, int Debtore = 0, int Creditor = 0, string Title = "", string DA = "",
     IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, string mac = "")
        {


            FinancialDocuments obj = new FinancialDocuments();
            obj.ID = ID;
            if (User_ID == 0)
            {
                obj.User_ID = null;
            }
            else
            {
                obj.User_ID = User_ID;
            }



            if (Detailedtitle_ID == 0)
            {
                obj.Detailedtitle_ID = null;
            }
            else
            {
                obj.Detailedtitle_ID = Detailedtitle_ID;
            }

            if (Creditlamdicatros_ID == 0)
            {
                obj.Creditlamdicatros_ID = null;
            }
            else
            {
                obj.Creditlamdicatros_ID = Creditlamdicatros_ID;
            }
            DateTime DataDocument2 = DateTime.Now;
            DA = DataDocument2.ToShamsi();


            obj.MacAdress = GetMacAddress();

            //obj.DataDocument = DataDocument;
            obj.Rigstertime = DA;
            obj.Peyman_ID = Peyman_ID;


            obj.Debtore = Debtore;
            obj.Creditor = Creditor;
            obj.Title = Title;



            if (files != null)
            {
                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;

                        return rep_def_rsn.Update2(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return rep_def_rsn.Update2(obj).ToString();
            }
        }
        public ActionResult sendfinancial(List<int> pyman_ID, List<int?> User_ID, DateTime ToDate = default, DateTime Fromdata = default, string TitleGhest = "", string TitleMarkazHazine = "", int FK_MarkazHazine = 0)
        {
            if (pyman_ID == null || !pyman_ID.Any() || User_ID == null || !User_ID.Any())
            {
                return View("Error");
            }

            var pyman = db.FinancialDocuments
                .Where(p => pyman_ID.Contains((int)p.Peyman_ID) && User_ID.Contains(p.User_ID) && p.DataDocument >= Fromdata && p.DataDocument <= ToDate)
                .ToList();

            return View(pyman);
        }



        /// <param name="Filters"></param>

        [AuthorizeAAA]
        public string FinancialDocuments_Add(List<FinancialDocuments> Filters)
        {
            return rep_def_rsn.Create(Filters);
        }


        public string tbDetailedtitle_Add(List<tbDetailedtitle> Filters)
        {
            return rep_def_rsn.Create2(Filters);
        }
        #endregion

        #region Report 
        #region variables
        #endregion
        #region pages

        [AuthorizeAAA]
        public ActionResult Report()
        {
            return View();
        }
        public ActionResult GetTaraz(List<int> personelCode, List<int?> PeymanCode, string TitleGhest,
            string TitleMarkazHazine, int FK_MarkazHazine, DateTime FromDate, DateTime ToDate)
        {
            //دریافت داده اسناد مالی
            sendfinancial(personelCode, PeymanCode /*, TitleGhest, TitleMarkazHazine, FK_MarkazHazine*/, FromDate/*, ToDate*/);
            //دریافت داده فیش حقوقی
            return PartialView();
        }

        #endregion
        #region Functions
        public ActionResult ExportAddMoalefeExcel2()
        {
            var OutPutFile = SetDataExcel_AddMoalefe2();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            return File(stream.ToArray(), mimeType, "asnamalypersnel" + extension);
        }


        public Workbook SetDataExcel_AddMoalefe2()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/asnadmalyusr.xlsx"));
            Row Row;
            var rc = db.tbUsers.ToList();
            //name and lastname and groupjob
            int counter = 3;
            //var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in rc)
            {
                Row = new Row() { Height = 20, Index = counter };
                //var t = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == item.installment_ID && p.Month == 2).FirstOrDefault();

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
                    }
                });

                Row.AddCells(new List<Cell>()
                {
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
            

                counter++;
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
            return Moalefeexcelfile;
        }

        public ActionResult ImportExcel_Addasnadpersonal(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_asnalpersonal(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        //public static DateTime PersianDateToDateTime(string persianDate)
        //{
        //    PersianCalendar persianCalendar = new PersianCalendar();
        //    string[] parts = persianDate.Split('/');
        //    int year = int.Parse(parts[2]);
        //    int month = int.Parse(parts[1]);
        //    int day = int.Parse(parts[0]);
        //    return persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        //}

        public string GetDataFromExcel_asnalpersonal(HttpPostedFileBase MyExcelStream)//برای نیروهای امانی
        {
            int mlfvlfsh_Year2=0;
            float mlfvlfsh_Value;
            string name = "";
            string timeend = ""; string title2 = "";

            List<FinancialDocuments> issabt = new List<FinancialDocuments>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, System.IO.Path.GetExtension(MyExcelStream.FileName));
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
                                var row2 = workbook.Sheets[0].Rows[1];
                                var Name2 = row2.Cells[0];
                                var Name3 = row2.Cells[1];
                                var Name4 = row2.Cells[2];
                                if (Name3.Value != null)
                                {
                                    int mlfvlfsh_Year;
                                    bool isNumeric = int.TryParse(Name3.Value.ToString(), out mlfvlfsh_Year);

                                    if (!isNumeric)
                                    {
                                        return "مقدار ستون نام در سطر " + " ( " + 1 + " ) " + "عددی نمی باشد";
                                    }
                                    else
                                    {
                                        mlfvlfsh_Year2 = mlfvlfsh_Year;
                                    }
                                }
                               
                                    if (Name2.Value != null)
                                    {
                                        timeend = Name2.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + 2 + " ) " + "را پر کنید ";
                                    }





                                if (Name4.Value != null)
                                {
                                    title2 = Name4.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                }
                                else
                                {
                                    return " لطفا ستون نام  سطر " + " ( " + 3 + " ) " + "را پر کنید ";
                                }
                                string inputDate2 = timeend;
                                DateTime persianDate2 = PersianDateToDateTime(inputDate2);

                                tbfkfinancial tbfina =new tbfkfinancial();
                                tbfina.numbershomar =  mlfvlfsh_Year2;
                                tbfina.DataDocument = persianDate2;
                                tbfina.Title = title2;
                                db.tbfkfinancial.Add(tbfina);
                                db.SaveChanges();


                                for (int i = 3; i < count; i++)
                                {
                                    FinancialDocuments fin =new FinancialDocuments();
                                    int id = 0;
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[1];
                                    if (Name.Value != null)
                                    {
                                        int mlfvlfsh_Year;
                                        bool isNumeric = int.TryParse(Name.Value.ToString(), out mlfvlfsh_Year);

                                        if (!isNumeric)
                                        {
                                            return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                        }
                                        else
                                        {
                                            var c2 = db.tbUsers.Where(p => p.usr_Personal_ID == mlfvlfsh_Year).FirstOrDefault();
                                            if (c2 != null)
                                            {
                                                fin.User_ID = c2.usr_ID;
                                            }
                                            else
                                            {
                                                return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";

                                            }
                                        }
                                     
                                    }


                                    var maskans = row.Cells[3];//مسکن
                                    if (double.TryParse(maskans.Value.ToString(), out double CaranStandard222))
                                    {
                                        fin.Debtore = CaranStandard222; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var maskans2 = row.Cells[2];
                                    if (double.TryParse(maskans2.Value.ToString(), out double CaranStandard2222))
                                    {
                                        fin.Creditor = CaranStandard2222; // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var maskans4 = row.Cells[4];

                                    if (maskans4.Value != null)
                                    {
                                        fin.Title = maskans4.Value.ToString(); // Assuming month.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    fin.FK_final = tbfina.ID;
                                    fin.DataDocument = persianDate2;
                                    issabt.Add(fin);
                             



                                }
                                db.FinancialDocuments.AddRange(issabt);
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
                    return "True";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;

                }
            }
        }
        [HttpPost]
        public JsonResult delete_bodjeh(int id = 0)
        {
            var find = db.FinancialDocuments.Where(s => s.FK_final == id).ToList();
            var find2 = db.tbfkfinancial.Where(s => s.ID == id).ToList();

            if (find.Any())
            {
                db.FinancialDocuments.RemoveRange(find);
                db.SaveChanges();
                db.tbfkfinancial.RemoveRange(find2);
                db.SaveChanges();
                return Json(true);

            }

            return Json(false);
        }


        #endregion


        #endregion
    }
}