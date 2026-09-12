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
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using System.Globalization;
using Microsoft.Ajax.Utilities;
using Telerik.Windows.Documents.Spreadsheet.Model;
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


namespace SaabWebProject.Areas.Setting.Controllers
{
    public class ReportingController : Controller
    {
        System.Globalization.PersianCalendar cs2 = new System.Globalization.PersianCalendar();

        SaabEntities db = new SaabEntities();
        List<Reportvalu> fishValues = new List<Reportvalu>();

        // GET: Setting/Reporting
        public ActionResult Index()
        {
            return View();
        }

        public static string ToShamsi(DateTime miladiDate)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            string persianDate = string.Format("{0}/{1}/{2}",
                persianCalendar.GetYear(miladiDate),
                persianCalendar.GetMonth(miladiDate).ToString("00"),
                persianCalendar.GetDayOfMonth(miladiDate).ToString("00"));
            return persianDate;
        }



        public static int CalculateMonthDifference(DateTime startDate, DateTime endDate)
        {
            // محاسبه تعداد ماه‌ها بین دو تاریخ
            int months = (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;

            // اگر تاریخ پایان از تاریخ شروع زودتر است، کاهش ماه‌ها
            if (endDate.Day < startDate.Day)
            {
                months--;
            }

            return months;
        }

        public async Task< ActionResult> reportmamorin(int month=3,int year=1403)
        {
            var moalf =await db.tbContractMoalefeDastmozdi.ToListAsync();
            List<tbUsers> usr = new List<tbUsers>();
            var fish = await db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month <= month && p.mlfvlfsh_Year == year).ToListAsync();
            var fish2 =await db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == month && p.mlfvlfsh_Year == year).ToListAsync();

            var eydipadash = moalf.Where(p => p.md_Title == "عیدی و پاداش ( ماهانه-ریال)").FirstOrDefault();
            var sanavat = moalf.Where(p => p.md_Title == "سنوات خدمت (ماهانه-ریال)").FirstOrDefault();
            var hogog = moalf.Where(p => p.md_Title == "خالص قابل دریافت (ریال)").FirstOrDefault(); 
            var tablet = moalf.Where(p => p.md_Title == "کمک هزینه تبلت و رایانه(ریال)").FirstOrDefault();
            var car = moalf.Where(p => p.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").FirstOrDefault();
            var beme = moalf.Where(p => p.md_Title == "حق بیمه سهم کارمند (ریال)").FirstOrDefault();
            var bemetakmil = moalf.Where(p => p.md_Title == "بیمه تکمیلی(ریال)").FirstOrDefault();
            var malyat = moalf.Where(p => p.md_Title == "مالیات سهم کارمند (ریال)").FirstOrDefault();
            long traz = 0;
            long traz2 = 0;
            long traz3 = 0;
            var us = fish2.Where(p => p.FK_Moalefe == hogog.md_ID&&p.mlfvlfsh_Value>0).ToList();
            foreach (var item in us)
            {
                usr.Add(item.tbUsers);
            }


            var motalbat = moalf.Where(p => p.md_Title == "پیشخوان-مامورین-مانده مطالبات").FirstOrDefault();
            var userAverageValues = new List<(long usr_ID, long averageValue)>();

            foreach (var it in usr)
            {
                var t =await db.FinancialDocuments.Where(p => p.User_ID == it.usr_ID).ToListAsync();
                long averageValue = 0;
                var avalfish = fish.Where(p => p.FK_Moalefe == hogog.md_ID &&
                                               p.FK_User == it.usr_ID &&
                                               p.mlfvlfsh_Month <= month &&
                                               p.mlfvlfsh_Year == year).ToList();
                if (avalfish.Any())
                {
                    averageValue = (long)avalfish.Average(p => p.mlfvlfsh_Value); // Replace 'mlfvlfsh_Value' with the actual property name
                }

                userAverageValues.Add((it.usr_ID, averageValue));
            }
            var sortedUserAverageValues = userAverageValues.OrderByDescending(x => x.averageValue).ToList();

            //foreach (var it in usr)
            //{
            //    var t = db.FinancialDocuments.Where(p => p.User_ID == it.usr_ID).ToList();
            //    long averageValue = 0;
            //    var avalfish = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID&&p.mlfvlfsh_Month<=month&&p.mlfvlfsh_Year==year).ToList();
            //    if (avalfish.Any())
            //    {
            //         averageValue =(long) avalfish.Average(p => p.mlfvlfsh_Value); // Replace 'value' with the actual property name
            //    }
            //    long eb = 0;
            //    long en = 0;

            //}
            foreach (var it in usr)
            {
                var t =await db.FinancialDocuments.Where(p => p.User_ID == it.usr_ID).ToListAsync();
                foreach (var item1 in t)
                {
                    if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                    {
                        PersianCalendar persianCalendar = new PersianCalendar();
                        DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                        int persianYear = persianCalendar.GetYear(dataDocument);
                        int persianMonth = persianCalendar.GetMonth(dataDocument);

                        if (persianYear == year && persianMonth <= month)
                        {
                            if (item1.Debtore != 0 && item1.Debtore != null)
                            {
                                traz -= (long)item1.Debtore;
                                traz3 -= (long)item1.Debtore;

                            }
                            else if (item1.Creditor != 0 && item1.Creditor != null)
                            {
                                traz += (long)item1.Creditor;
                                traz2 += (long)item1.Creditor;


                            }
                        }

                    }
                }
                double mohasb = 0;
                double saneyd = 0;

                if (month != 12)
                {
                    var eydi = fish.Where(p => p.FK_Moalefe == eydipadash.md_ID && p.FK_User == it.usr_ID).ToList();
                    foreach (var item in eydi)
                    {
                        saneyd += (double)item.mlfvlfsh_Value;

                    }
                    var sanav = fish.Where(p => p.FK_Moalefe == sanavat.md_ID && p.FK_User == it.usr_ID).ToList();
                    foreach (var item in sanav)
                    {
                        saneyd += (double)item.mlfvlfsh_Value;

                    }
                    var nahaei = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID).ToList();
                    foreach (var item in nahaei)
                    {
                        mohasb += (double)item.mlfvlfsh_Value;

                    }
                    mohasb += traz;
                }
                else
                {

                    var nahaei = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID).ToList();
                    foreach (var item in nahaei)
                    {
                        mohasb += (double)item.mlfvlfsh_Value;

                    }
                    mohasb += traz;

                }
                Reportvalu motalbaytt = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-مانده مطالبات",
                    Value = mohasb.ToString(),
                    Usr_ID=it.usr_ID,

                };
                fishValues.Add(motalbaytt);
                var taghsisi=await db.tbUser_link_BusinessSide.Where(p=>p.FK_u_ID==it.usr_ID&&p.Status==true).FirstOrDefaultAsync();
                if (taghsisi != null)
                {
                    var tag = await db.tbBusinessSide.Where(p => p.up_Id == taghsisi.FK_up_ID).FirstOrDefaultAsync();

                    Reportvalu taghsis = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تخصص اصلی",
                        Value = tag.up_name.ToString(),
                        Usr_ID = it.usr_ID,

                    };
                    fishValues.Add(taghsis);
                }
            





                Reportvalu zaghayr = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-مانده ذخایر",
                    Value = saneyd.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(zaghayr);
                var agsatt=await db.tbinstallments.Where(p=>p.User_ID==it.usr_ID&&p.tb_Subset_of_installments.Any(s=>s.Month<=month&&p.Year==year)).FirstOrDefaultAsync();
                double agsatkol = 0;
                if (agsatt != null)
                {
                    agsatkol = agsatt.Total_Amount;
                }
                //foreach (var item in agsatt)
                //{
                //    var subset =db.tb_Subset_of_installments.Where(p=>p.fk_installment_ID==it.usr_ID&&p.Year==year&&p.Month<=month).ToList();
                //    foreach (var item1 in subset)
                //    {
                        
                //    }
                //}
                double mondeagsat = 0;
               var takmoiul = fish.Where(p => p.FK_Moalefe == bemetakmil.md_ID && p.FK_User == it.usr_ID).ToList();
                foreach (var item in takmoiul)
                {
                    mondeagsat += (double)item.mlfvlfsh_Value;

                }
                agsatkol -= mondeagsat;
                Reportvalu agsat = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-مانده اقساط",
                    Value = agsatkol.ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(agsat);
                Reportvalu padash = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-مانده پاداش و جرایم",
                    Value = "0",
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(padash);
                Reportvalu kol = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-جمع مطالبات",
                    Value = (agsatkol+ mohasb+ saneyd).ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(kol);






                var summpepol = fish.Where(p => p.FK_Moalefe == hogog.md_ID).ToList();

                Reportvalu sumpepole = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-ارزیابی-جمع نفرات",
                    Value = summpepol.Count().ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(sumpepole);







                var ghar=await db.tbUserContracts.Where(p=>p.FK_UserID==it.usr_ID).ToListAsync();
                if(ghar.Count>1)
                {
                    var g=await db.tbUserContracts.Where(p=>p.FK_UserID ==it.usr_ID).FirstOrDefaultAsync();
                    var end = db.tbUserContracts.Where(p => p.FK_UserID == it.usr_ID).OrderByDescending(s => s.usc_ID).FirstOrDefault();

                    Reportvalu numberoneghardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-شماره قرارداد1",
                        Value = g.usc_ContractNumber.ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(numberoneghardad);

                    Reportvalu startgarda1 = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ شروع قرارداد1",
                        Value = ToShamsi((DateTime)g.usc_StartTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(startgarda1);

                    Reportvalu endgardad1 = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ خاتمه قرارداد1",
                        Value = ToShamsi((DateTime)g.usc_EndTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(endgardad1);

                    Reportvalu statuseoneghardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-وضعیت قرارداد1",
                        Value ="0",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(statuseoneghardad);

                    Reportvalu numbersecondgardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-شماره قرارداد2",
                        Value = end.usc_ContractNumber.ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(numbersecondgardad);

                    Reportvalu startdatesecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ شروع قرارداد2",
                        Value = ToShamsi((DateTime)end.usc_StartTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(startdatesecond);

                    Reportvalu enddatsecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ خاتمه قرارداد2",
                        Value = ToShamsi((DateTime)end.usc_EndTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(enddatsecond);



                    Reportvalu statusesecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-وضعیت قرارداد2",
                        Value = "1",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(statusesecond);
                }
               else if (ghar.Count == 1)
                {
                    var g = db.tbUserContracts.Where(p => p.FK_UserID == it.usr_ID).FirstOrDefault();
                    var end = db.tbUserContracts.Where(p => p.FK_UserID == it.usr_ID).OrderByDescending(s => s.usc_ID).FirstOrDefault();

                    Reportvalu numberoneghardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-شماره قرارداد1",
                        Value = g.usc_ContractNumber.ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(numberoneghardad);

                    Reportvalu startgarda1 = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ شروع قرارداد1",
                        Value = ToShamsi((DateTime)g.usc_StartTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(startgarda1);

                    Reportvalu endgardad1 = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ خاتمه قرارداد1",
                        Value = ToShamsi((DateTime)g.usc_EndTime).ToString(),
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(endgardad1);

                    Reportvalu statuseoneghardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-وضعیت قرارداد1",
                        Value = "1",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(statuseoneghardad);

                    Reportvalu numbersecondgardad = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-شماره قرارداد2",
                        Value = "0",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(numbersecondgardad);

                    Reportvalu startdatesecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ شروع قرارداد2",
                        Value = "0",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(startdatesecond);

                    Reportvalu enddatsecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ خاتمه قرارداد2",
                        Value = "0",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(enddatsecond);



                    Reportvalu statusesecond = new Reportvalu
                    {
                        Title = "پیشخوان-مامورین-تاریخ خاتمه قرارداد2",
                        Value = "0",
                        Usr_ID = it.usr_ID,

                    };

                    fishValues.Add(statusesecond);
                }

                var avalfish = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID).FirstOrDefault();
                var endfisj = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID).OrderByDescending(s=>s.mlfvlfsh_ID).FirstOrDefault();
                long eb = 0;
                long en = 0;
                if (avalfish != null)
                {
                    foreach (var item1 in t)
                    {
                        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                        {
                            PersianCalendar persianCalendar = new PersianCalendar();
                            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                            int persianYear = persianCalendar.GetYear(dataDocument);
                            int persianMonth = persianCalendar.GetMonth(dataDocument);

                            if (persianYear == avalfish.mlfvlfsh_Year && persianMonth == avalfish.mlfvlfsh_Month)
                            {
                                if (item1.Debtore != 0 && item1.Debtore != null)
                                {
                                    eb -= (long)item1.Debtore;

                                }
                                else if (item1.Creditor != 0 && item1.Creditor != null)
                                {
                                    eb += (long)item1.Creditor;


                                }
                            }

                        }
                    }
                    eb += (long)avalfish.mlfvlfsh_Value;

                }


                if (endfisj != null)
                {
                    foreach (var item1 in t)
                    {
                        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                        {
                            PersianCalendar persianCalendar = new PersianCalendar();
                            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                            int persianYear = persianCalendar.GetYear(dataDocument);
                            int persianMonth = persianCalendar.GetMonth(dataDocument);

                            if (persianYear == endfisj.mlfvlfsh_Year && persianMonth == endfisj.mlfvlfsh_Month)
                            {
                                if (item1.Debtore != 0 && item1.Debtore != null)
                                {
                                    en -= (long)item1.Debtore;

                                }
                                else if (item1.Creditor != 0 && item1.Creditor != null)
                                {
                                    en += (long)item1.Creditor;


                                }
                            }

                        }
                    }
                    en += (long)endfisj.mlfvlfsh_Value;

                }
           

                long n = en - eb;
                long v = n * 100;
                long moo = 0;
                if (eb != 0)
                {
                    moo = (long)(v / eb);

                }
                long mo = (long)moo;
                var rank = sortedUserAverageValues.FindIndex(p => p.usr_ID == it.usr_ID)+1;
                Reportvalu ROSHD = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-ارزیابی-میزان رشد",
                    Value = mo.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(ROSHD);



                Reportvalu ROTBBEH = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-ارزیابی-رتبه",
                    Value = rank.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(ROTBBEH);




                Reportvalu GOVAHIHAVI = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-فن ورز هوایی",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIHAVI);


                Reportvalu GOVAHIGARM = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-فن ورز گرم",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIGARM);


                Reportvalu GOVAHIBARGH = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-ایمنی در برق",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIBARGH);



                Reportvalu GOVAHIMAMORMOSHTRAK = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-مامور مشترکین",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIMAMORMOSHTRAK);





                Reportvalu GOVAHIGRAAT = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-مامور قرائت",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIGRAAT);



                Reportvalu GOVAHIEMENY = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-گواهینامه-ایمنی فردی",
                    Value = 1.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIEMENY);



                Reportvalu GOVAHITEST = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-آموزش-تست کنتور",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHITEST);





                Reportvalu GOVAHIMOGAZ = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-آموزش-شناسایی غیر مجاز",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIMOGAZ);





                Reportvalu GOVAHIMOSHTRAK = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-آموزش-برخورد با مشترکین",
                    Value = 0.ToString(),
                    Usr_ID = it.usr_ID,

                };
                fishValues.Add(GOVAHIMOSHTRAK);




                //Reportvalu GOVAHI = new Reportvalu
                //{
                //    Title = "",
                //    Value = 0.ToString(),
                //    Usr_ID = it.usr_ID,

                //};
                //fishValues.Add(RPTBEH);






                double mohasebatttt = 0;
                List<double> values = new List<double>();
                double max = 0;
                double min = 0;
                double average = 0;
               
                    
                    var nahaei22 = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == it.usr_ID&&p.mlfvlfsh_Value>0).ToList();
                    foreach (var item in nahaei22)
                    {
                        double value = (double)item.mlfvlfsh_Value;
                        mohasebatttt += value;
                        values.Add(value);
                    }
                    if (values.Count > 0)
                    {
                         max = values.Max();
                         min = values.Min();
                         average = values.Average();

                      
                    }
                

                Reportvalu maxhogho = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-حداکثر",
                    Value = max.ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(maxhogho);



                Reportvalu avragehoghogh = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-متوسط",
                    Value = average.ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(avragehoghogh);
                Reportvalu minhogho = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-حداقل",
                    Value = min.ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(minhogho);


                Reportvalu maharatghanebi = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-مهارت های جانبی",
                    Value = 5.ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(maharatghanebi);



                Reportvalu updatenew = new Reportvalu
                {
                    Title = "پیشخوان-مامورین-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Usr_ID = it.usr_ID,

                };

                fishValues.Add(updatenew);






            }
            var tt = await db.tbMoalefeValuePishkhan
           .Where(p => p.mlfval_Month == month && p.mlfval_Year == year)
           .ToListAsync();

            List<tbMoalefeValuePishkhan> piskghan = new List<tbMoalefeValuePishkhan>();

            foreach (var item in fishValues)
            {
                var find = moalf
                    .Where(p => p.md_Title == item.Title)
                    .FirstOrDefault();

                if (find != null)
                {
                    // Parse item.Value to a decimal
                    if (decimal.TryParse(item.Value, out decimal parsedValue))
                    {
                        // Truncate the decimal part
                        var truncatedValue = (long)parsedValue;

                        if (tt.Count != 0)
                        {
                            var f = tt
                                .Where(p => p.mlfval_FKUser == item.Usr_ID && p.mlfval_FKMoalafeDastmozdi == find.md_ID)
                                .FirstOrDefault();

                            if (f != null)
                            {
                                f.mlfval_Value = truncatedValue.ToString();
                                await db.SaveChangesAsync();
                            }
                            else
                            {
                                tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan
                                {
                                    mlfval_Value = truncatedValue.ToString(),
                                    mlfval_Month = month,
                                    mlfval_Year = year,
                                    mlfval_FKUser = item.Usr_ID,
                                    mlfval_FKMoalafeDastmozdi = find.md_ID
                                };
                                piskghan.Add(pis);
                            }
                        }
                        else
                        {
                            tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan
                            {
                                mlfval_Value = truncatedValue.ToString(),
                                mlfval_Month = month,
                                mlfval_Year = year,
                                mlfval_FKUser = item.Usr_ID,
                                mlfval_FKMoalafeDastmozdi = find.md_ID
                            };
                            piskghan.Add(pis);
                        }
                    }
                    else
                    {
                        // Handle the case where parsing fails, if needed
                        // For example, log an error or set a default value
                        if (tt.Count != 0)
                        {
                            var f = tt
                                .Where(p => p.mlfval_FKUser == item.Usr_ID && p.mlfval_FKMoalafeDastmozdi == find.md_ID)
                                .FirstOrDefault();

                            if (f != null)
                            {
                                f.mlfval_Value = item.Value;
                                await db.SaveChangesAsync();
                            }
                            else
                            {
                                tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan
                                {
                                    mlfval_Value = item.Value,
                                    mlfval_Month = month,
                                    mlfval_Year = year,
                                    mlfval_FKUser = item.Usr_ID,
                                    mlfval_FKMoalafeDastmozdi = find.md_ID
                                };
                                piskghan.Add(pis);
                            }
                        }
                        else
                        {
                            tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan
                            {
                                mlfval_Value = item.Value,
                                mlfval_Month = month,
                                mlfval_Year = year,
                                mlfval_FKUser = item.Usr_ID,
                                mlfval_FKMoalafeDastmozdi = find.md_ID
                            };
                            piskghan.Add(pis);
                        }
                    }
                }
            }

            db.tbMoalefeValuePishkhan.AddRange(piskghan);
            await db.SaveChangesAsync();


            return View();
        }



        //http://localhost:6061/Setting/Reporting/reportpymn

        public async Task< ActionResult> reportpymn(int month = 3, int year = 1403)
        {
            var moalf = await db.tbContractMoalefeDastmozdi.ToListAsync();
            var usr = await db.tbUsers.ToListAsync();
            var pymn = await db.tbPeymanContracts.Where(p=>p.Inactive!=true&&p.pec_ID== 1076).ToListAsync();
            var pymn2 = await db.tbPeymanContracts.Where(p =>  p.pec_ID >= 1076).ToListAsync();
            var linkk =await db.Link_User_And_Peyman.Where(p => p.tbPeymanContracts.Inactive != true).ToListAsync();

            var fish = await db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month <= month && p.mlfvlfsh_Year == year&&p.mlfvlfsh_Submit==true).ToListAsync();
            var eydipadash = moalf.Where(p => p.md_Title == "عیدی و پاداش ( ماهانه-ریال)").FirstOrDefault();
            var sanavat = moalf.Where(p => p.md_Title == "سنوات خدمت (ماهانه-ریال)").FirstOrDefault();
            var hogog = moalf.Where(p => p.md_Title == "خالص قابل دریافت (ریال)").FirstOrDefault();
            var tablet = moalf.Where(p => p.md_Title == "کمک هزینه تبلت و رایانه(ریال)").FirstOrDefault();
            var car = moalf.Where(p => p.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").FirstOrDefault();
            var beme = moalf.Where(p => p.md_Title == "حق بیمه سهم کارمند (ریال)").FirstOrDefault();
            var bemetakmil = moalf.Where(p => p.md_Title == "بیمه تکمیلی(ریال)").FirstOrDefault();
            var malyat = moalf.Where(p => p.md_Title == "مالیات سهم کارمند (ریال)").FirstOrDefault();
            long traz = 0;
            long traz2 = 0;
            long traz3 = 0;

            var motalbat = moalf.Where(p => p.md_Title == "پیشخوان-مامورین-مانده مطالبات").FirstOrDefault();


            foreach(var itt in pymn)
            {

                var countt = 0;
                var activ = 0;
                foreach(var iyyt in pymn2)
                {
                    if (iyyt.tbCompanies.ID == itt.tbCompanies.ID)
                    {
                        countt += 1;
                        if (iyyt.Inactive == true)
                        {
                            activ += 1;
                        }
                    }
                }
                Reportvalu namecompain = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-نام شرکت",
                    Value = itt.tbCompanies.CompanyName.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(namecompain);
                Reportvalu addrescompain = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-آدرس",
                    Value = itt.tbCompanies.Company_Address.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(addrescompain);
                Reportvalu phoncomain = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-تلفن",
                    Value = itt.tbCompanies.Telephone_1.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(phoncomain);
                Reportvalu managecompain = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-نام مدیرعامل",
                    Value = itt.tbCompanies.tbUsers.FullName.ToString() ??"0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(managecompain);
                Reportvalu kolpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-تعدادکل پیمان",
                    Value = countt.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(kolpymn);
                Reportvalu activpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-تعدادپیمان فعال",
                    Value = activ.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(activpymn);
                Reportvalu phonemange = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-پیمانکار-موبایل",
                    Value = itt.tbCompanies.tbUsers.usr_PhoneNumber.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(phonemange);
                Reportvalu startpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تاریخ شروع",
                    Value = ToShamsi((DateTime)itt.pec_StartTime).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(startpymn);
                Reportvalu endtime = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تاریخ خاتمه",
                    Value = ToShamsi((DateTime)itt.pec_EndTime).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(endtime);
                Reportvalu numberpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-شماره قرارداد",
                    Value = itt.pec_ContractNumber.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(numberpymn);
                var elhagh = db.tbPeymanElhaghie.Where(p => p.FK_PeymanID == itt.pec_ID).FirstOrDefault();
                if (elhagh != null)
                {
                    Reportvalu elhaghgg = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مبلغ الحاقیه",
                        Value = elhagh.Price.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(elhaghgg);
                }
                else
                {
                    Reportvalu elhaghgg = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مبلغ الحاقیه",
                        Value = "0",
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(elhaghgg);
                }

                var sabad = db.tbDetermining_creditline
.Where(p => p.Pyman_ID == itt.pec_ID)   // Step 1: Filter by Pyman_ID
.GroupBy(s => s.FK_creditline)  // Step 2: Group by FK_creditline
.Select(g => g.OrderByDescending(w => w.elhagh ?? Int32.MinValue).FirstOrDefault())  // Step 3: Select the item with the largest elhagh (handling nulls) in each group
.ToList();

                long price = 0;
                List<string> title=new List<string>();
                foreach (var item in sabad)
                {
                    price += (long)item.Creditorcredit;
                    title.Add(item.tbCreditIndicators.Title.ToString());

                }

                var sorat = db.tbSoratvaziat.Where(p=>p.month<=month&&p.Year<=year&&p.Final_accept==true&&p.Fk_pymn==itt.pec_ID).ToList();
                long pricesorat = 0;
                foreach(var itttt in sorat.Where(s=>s.Finalvaluesorat!=null).ToList())
                {
                    pricesorat += (long)itttt.Finalvaluesorat;
                   
                }


                    Reportvalu movasab  = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-اعتبار مصوب",
                    Value = price.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(movasab);




                Reportvalu masrafpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-اعتبار مصرف شده",
                    Value = pricesorat.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(masrafpymn);
                Reportvalu baghmondeh= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-اعتبار باقیمانده",
                    Value = (price-pricesorat).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(baghmondeh);
                var fines = linkk.Where(p => p.FK_Peyman_ID == itt.pec_ID&&p.Status==true).ToList();
                int karkad = 0;
                foreach (var item in fines)
                {
                    var fishf = db.tbMoalefeValueFish.Where(p => p.FK_User==item.FK_User_ID&&p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)" && p.mlfvlfsh_Month == month && p.mlfvlfsh_Year == year && p.mlfvlfsh_Submit == true).FirstOrDefault();
                    if (fishf != null)
                    {
                        var finam=db.tbMoadelPadashJarimeAyab.Where(p=>p.UserID==item.FK_User_ID&&p.Month==month&&p.Year ==year).FirstOrDefault();
                        if (finam!=null)
                        {
                            karkad += (int)finam.Moadel;
                        }
                    }
                }

                Reportvalu estandardkarkord= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-استاندارد تعداد کارکنان",
                    Value = karkad.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(estandardkarkord);

                var coun = linkk.Where(p => p.FK_Peyman_ID == itt.pec_ID).Count();

                Reportvalu karkonanmogod = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تعداد کارکنان موجود",
                    Value = coun.ToString() ??"0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(karkonanmogod);

                int monthDifference = CalculateMonthDifference((DateTime)itt.pec_StartTime, (DateTime)itt.pec_EndTime);

                Reportvalu countghardad = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مدت قرارداد",
                    Value = monthDifference.ToString() ?? "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(countghardad);

                if (itt.Inactive == true)
                {
                    Reportvalu statisepymn = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-وضعیت پیمان",
                        Value = "0",
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(statisepymn);
                }
                else
                {
                    Reportvalu statisepymn = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-وضعیت پیمان",
                        Value = "1",
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(statisepymn);
                }
                Reportvalu updatebalance= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatebalance);


                Reportvalu motalebat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-مانده مطالبات",
                    Value =  "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(motalebat);
                var numbersor = 0;
                double hosn = 0.5;
                double bem = 0.5;
                double sarmayi = 0;
                double jari = 0;
                var sr = db.tbSoratvaziat.Where(p => p.Fk_pymn == itt.pec_ID && p.month == month && p.Year == year).FirstOrDefault();
                if(sr!=null)
                {
                    numbersor = sr.number_sorat ??0;
                }
                var finhos = db.tbpeymaninformatio.Where(p => p.Fk_pymn == itt.pec_ID && p.numbersorat == numbersor&&p.accept==true).FirstOrDefault();
                if(finhos!=null)
                {
                    hosn = finhos.hosin ??0;
                    bem = finhos.hosin ?? 0;
                    sarmayi = finhos.sarmayi ?? 0;
                    jari = finhos.jari ?? 0;


                }
                Reportvalu bemehkarkonan = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-سپرده بیمه",
                    Value = bem.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(bemehkarkonan);



                Reportvalu hosiananjam = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-سپرده حسن انجام کار",
                    //Value = (2*0.05 * pricesorat).ToString(),
                    Value = hosn.ToString(),

                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(hosiananjam);


                Reportvalu sabtpymnbalans = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-وضعیت ثبت نشده",
                    Value = "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sabtpymnbalans);

                Reportvalu padashjaremhpym = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-ذخیره پاداش و جریمه",
                    Value = "0",
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(padashjaremhpym);

                Reportvalu sumtraz = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-جمع تراز",
                    //Value = ((2 * 0.05 * pricesorat)+ ( 0.05 * pricesorat)).ToString()??"0",
                    Value = ((hosn) + (bem)).ToString() ?? "0",

                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sumtraz);
                var finalValues = sorat.Where(s=>s.Finalvaluesorat!=null).Select(itttt => (long)itttt.Finalvaluesorat).ToList();

                long min = finalValues.Min();
                long max = finalValues.Max();
                double average = finalValues.Average();

                Reportvalu maxsorat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-حداکثر",
                    Value = max.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(maxsorat);

                Reportvalu minsorat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-حداقل",
                    Value = min.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(minsorat);

                Reportvalu avrangpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس پیمان-متوسط",
     Value = average.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(avrangpymn);


                Reportvalu dataupdat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(avrangpymn);



                double agsatkol = 0;

                double maxxx = 0;
                double minnn = 0;
                double avraggee = 0;
                List<double> valuesss = new List<double>();

                double mohasb = 0;
                double saneyd = 0;
                foreach (var ty in linkk.Where(p => p.FK_Peyman_ID == itt.pec_ID).ToList())
                {
                    var t = db.FinancialDocuments.Where(p => p.User_ID == ty.tbUsers.usr_ID).ToList();
                    foreach (var item1 in t)
                    {
                        if (item1.DataDocument.HasValue)  // بررسی می‌کنیم که مقدار تاریخ نالیبل نیست
                        {
                            PersianCalendar persianCalendar = new PersianCalendar();
                            DateTime dataDocument = item1.DataDocument.Value; // تبدیل نالیبل به غیر نالیبل
                            int persianYear = persianCalendar.GetYear(dataDocument);
                            int persianMonth = persianCalendar.GetMonth(dataDocument);

                            if (persianYear == year && persianMonth <= month)
                            {
                                if (item1.Debtore != 0 && item1.Debtore != null)
                                {
                                    traz -= (long)item1.Debtore;
                                    traz3 -= (long)item1.Debtore;

                                }
                                else if (item1.Creditor != 0 && item1.Creditor != null)
                                {
                                    traz += (long)item1.Creditor;
                                    traz2 += (long)item1.Creditor;


                                }
                            }

                        }
                    }




                    var agsatt = db.tbinstallments.Where(p => p.User_ID == ty.tbUsers.usr_ID && p.tb_Subset_of_installments.Any(s => s.Month <= month && p.Year == year)).FirstOrDefault();
                    if (agsatt != null)
                    {
                        agsatkol = agsatt.Total_Amount;
                    }
                    //foreach (var item in agsatt)
                    //{
                    //    var subset =db.tb_Subset_of_installments.Where(p=>p.fk_installment_ID==it.usr_ID&&p.Year==year&&p.Month<=month).ToList();
                    //    foreach (var item1 in subset)
                    //    {

                    //    }
                    //}
                    double mondeagsat = 0;
                    var takmoiul = fish.Where(p => p.FK_Moalefe == bemetakmil.md_ID && p.FK_User == ty.tbUsers.usr_ID).ToList();
                    foreach (var item in takmoiul)
                    {
                        mondeagsat += (double)item.mlfvlfsh_Value;

                    }
                    agsatkol -= mondeagsat;



                   
                        var nahaeifishh = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == ty.tbUsers.usr_ID && p.mlfvlfsh_Value > 0).ToList();
                        foreach (var item in nahaeifishh)
                        {
                            double value = (double)item.mlfvlfsh_Value;
                            //mohasebatttt += value;
                            valuesss.Add(value);
                        }
                     
                    



                    if (month != 12)
                    {
                        var eydi = fish.Where(p => p.FK_Moalefe == eydipadash.md_ID && p.FK_User == ty.tbUsers.usr_ID).ToList();
                        foreach (var item in eydi)
                        {
                            saneyd += (double)item.mlfvlfsh_Value;

                        }
                        var sanav = fish.Where(p => p.FK_Moalefe == sanavat.md_ID && p.FK_User == ty.tbUsers.usr_ID).ToList();
                        foreach (var item in sanav)
                        {
                            saneyd += (double)item.mlfvlfsh_Value;

                        }
                        var nahaei = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == ty.tbUsers.usr_ID).ToList();
                        foreach (var item in nahaei)
                        {
                            mohasb += (double)item.mlfvlfsh_Value;

                        }
                        mohasb += traz;
                    }
                    else
                    {

                        var nahaei = fish.Where(p => p.FK_Moalefe == hogog.md_ID && p.FK_User == ty.tbUsers.usr_ID).ToList();
                        foreach (var item in nahaei)
                        {
                            mohasb += (double)item.mlfvlfsh_Value;

                        }
                        mohasb += traz;

                    }
                }


                if (valuesss.Count > 0)
                {
                    maxxx = valuesss.Max();
                    minnn = valuesss.Min();
                    avraggee = valuesss.Average();


                }
                Reportvalu modemotalebat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-مانده مطالبات",
                    Value = mohasb.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(modemotalebat);

               
                Reportvalu zaghayer = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-مانده ذخایر",
                    Value = saneyd.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(zaghayer);


                Reportvalu mondeahsat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-مانده اقساط",
                    Value = agsatkol.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(mondeahsat);



                Reportvalu mondemosharkat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-مانده مشارکت",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(mondemosharkat);



                Reportvalu zaghayergarmeh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-ذخیره پاداش و جریمه",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(zaghayergarmeh);



                Reportvalu maxkarmonann = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-حداکثر",
                    Value = maxxx.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(maxkarmonann);



                Reportvalu sumtarzz = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-جمع تراز",
                    Value = (mohasb+ agsatkol+ saneyd).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sumtarzz);
               

                Reportvalu avargekarkonann = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-متوسط",
                    Value = avraggee.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(avargekarkonann);

                Reportvalu minkarkonan = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-بالانس کارکنان-حداقل",
                    Value =minnn.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(minkarkonan);


                Reportvalu updatemarkaz = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemarkaz);



                Reportvalu updatetalafat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تلفات-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatetalafat);
                var moalfe = db.tbContractMoalefeDastmozdi.ToList();

                var x = moalfe.Where(p => p.md_Title == "انرژی تحویلی(kwh)").FirstOrDefault();
                var x1 = moalfe.Where(p => p.md_Title == "انرژی توزیعی(kwh)").FirstOrDefault();
                var x2 = moalfe.Where(p => p.md_Title == "وصول (ریال)").FirstOrDefault();

                var x3 = moalfe.Where(p => p.md_Title == "فروش (ریال)").FirstOrDefault();
                var x4 = moalfe.Where(p => p.md_Title == "وصول (درصد)").FirstOrDefault();
                var x5 = moalfe.Where(p => p.md_Title == "").FirstOrDefault();
                var x6 = moalfe.Where(p => p.md_Title == "").FirstOrDefault();
                double sorathatah = 0;
                double sorathaden = 0;
                double forosh = 0;
                double vosol = 0;
                var sabtamaljard = db.tbsaveSoratBasteh.Where(p => p.fk_pymn == itt.pec_ID && p.Final_Accept == true && p.tbSoratSavefromExcelMoalfe.Any(s => s.Month == month && s.Year == year)).FirstOrDefault();
                if (sabtamaljard != null)
                {

                    var sabtamal = db.tbSoratSavefromExcelMoalfe.Where(p => p.FK_SavedFunctionsID == sabtamaljard.ID).GroupBy(d => d.FK_city).ToList();
                    //شروع حساب و کتاب های انحراف ها و عملکرد ها 
                    foreach (var item2 in sabtamal)
                    {
                        //foreach (var item3 in item2)
                        //{
                        //    if (item3.FK_MoalfeDastmozdi == x.md_ID)
                        //    {
                        //        amalkr.amalkerd_EnergyTahvili = (long)item3.Value;
                        //    }
                        //    else if (item3.FK_MoalfeDastmozdi == x1.md_ID)
                        //    {
                        //        amalkr.amalkerd_EnergyToziee = (long)item3.Value;
                        //        sab += (long)item3.Value;
                        //    }
                        //    else if (item3.FK_MoalfeDastmozdi == x2.md_ID)
                        //    {
                        //        amalkr.amalkerd_Vosool = (float)item3.Value;
                        //    }
                        //    else if (item3.FK_MoalfeDastmozdi == x3.md_ID)
                        //    {
                        //        amalkr.amalkerd_Foroosh = (long)item3.Value;
                        //    }

                        //    else if (item3.FK_MoalfeDastmozdi == x4.md_ID)
                        //    {
                        //        amalkr.amalkerd_VosoolPercent = (float)item3.Value;
                        //    }
                        //}


                        var c9 = db.tbSoratSavefromExcelMoalfe
                            .Where(p => p.FK_city == item2.Key &&
                                          p.Month == month && p.Year == year).ToList();
                        foreach (var item3 in c9)
                        {
                            if (item3.FK_MoalfeDastmozdi == x.md_ID)
                            {
                                sorathatah += (double)item3.Value;
                            }
                            else if (item3.FK_MoalfeDastmozdi == x1.md_ID)
                            {
                                sorathaden += (double)item3.Value;
                            }
                            else if (item3.FK_MoalfeDastmozdi == x2.md_ID)
                            {
                                vosol += (double)item3.Value;
                            }
                            else if (item3.FK_MoalfeDastmozdi == x3.md_ID)
                            {
                                forosh += (double)item3.Value;
                            }

                        }



                    }
                }


                Reportvalu talafatdarsad = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تلفات-درصد",
                    Value = ((sorathatah- sorathatah)/ sorathatah*100).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(talafatdarsad);






                Reportvalu energyvorodi = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تلفات-انرژی ورودی",
                    Value = sorathatah.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemarkaz);





                Reportvalu tozietalafat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-تلفات-انرژی توزیعی",
                    Value = sorathaden.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(tozietalafat);


                Reportvalu updatevosol = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-وصول-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatevosol);

                Reportvalu vosolpercent = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-وصول-درصد",
                    Value = ((vosol - forosh) / vosol * 100).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(vosolpercent);

                Reportvalu vosolfrosh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-وصول-مبلغ فروش",
                    Value = forosh.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(vosolfrosh);




                Reportvalu vosolmablagh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-وصول-مبلغ وصول",
                    Value = vosol.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(vosolmablagh);



                Reportvalu updatenasb = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-زمان نصب-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatenasb);



                Reportvalu estandardnasb = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-زمان نصب-استاندارد",
                    Value = 5.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(estandardnasb);



                Reportvalu endnasb = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-زمان نصب-نتیجه",
                    Value =0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(endnasb);



                Reportvalu updtearzyabi = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updtearzyabi);


                Reportvalu arzyabipymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-تعداد پیمان تحت ارزیابی",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(arzyabipymn);

                long mohsbatpymn = 1;
                foreach (var tttty in db.tbPeymanContracts.Where(p => p.Inactive != true).ToList())
                {
                    string cleanedPrice = tttty.pec_Price.Replace(",", ""); // حذف کاماها

                    if (long.TryParse(cleanedPrice, out long pecPrice))
                    {
                        mohsbatpymn += pecPrice;
                    }
                    else
                    {
                        // ثبت لاگ برای مقدار نامعتبر
                        Console.WriteLine($"Cannot parse value: {tttty.pec_Price}");
                    }
                }


                long mohasbt = 0;
                if (long.TryParse(itt.pec_Price, out long pecPricerr))
                {
                    mohasbt = pecPricerr;
                }

                Reportvalu shaghesarzyabi = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-امتیاز تاثیر در شاخص",
                    Value = (mohasbt*100/ mohsbatpymn).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(shaghesarzyabi);
                var sumarz = 0;
                int sumaerzyab2 = 0;
                var countarz = 0;
                var avragearz = 0;
                long sumsor = 0;
                long padash = 0;
                long jarimeh = 0;
                long sumpadashandjarimeh = 0;
                double roshdo = 0;

                foreach (var itty in db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == itt.pec_ID&&p.Status==true).ToList())
                {
                    var find = db.tbMoalefeValuePishkhan.Where(s => s.mlfval_FKUser == itty.FK_User_ID && s.mlfval_Month == month && s.mlfval_Year == year && s.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-مامورین-ارزیابی-میزان رشد").FirstOrDefault();
                    if(find!=null)
                    {
                        if (find.mlfval_Value != null)
                        {
                            // حذف تمام کاماها از مقدار
                            string cleanValue = find.mlfval_Value.Replace(",", "");

                            if (int.TryParse(cleanValue, out int parsedValue))
                            {
                                sumaerzyab2 += parsedValue;
                                countarz += 1;
                            }
                        }

                    }
                }
                if (countarz != 0)
                {
                    avragearz =(int)( sumaerzyab2 / countarz);
                }
                foreach(var sor in db.tbSoratvaziat.Where(s => s.Fk_pymn == itt.pec_ID && s.padash != null&&s.month<=month&&s.Year<=year).ToList())
                {
                    padash +=(long) sor.padash;
                }
                foreach (var sor in db.tbSoratvaziat.Where(s => s.Fk_pymn == itt.pec_ID && s.jaremeh != null && s.month <= month && s.Year <= year).ToList())
                {
                    jarimeh += (long)sor.jaremeh;
                }
                sumpadashandjarimeh = (long)(padash - jarimeh);
                roshdo = (double)(sumpadashandjarimeh / pricesorat) + 1;
                roshdo *= 100;

                Reportvalu rosdpercentt = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-درصدرشد",
                    Value = avragearz.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(rosdpercentt);


                Reportvalu rotbehrosd = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-رتبه رشد",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(rotbehrosd);

                Reportvalu  etbarpymn= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-امتیاز ارزیابی",
                    Value = roshdo.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(etbarpymn);



                Reportvalu etbarrotbeh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-رتبه ارزیابی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(etbarrotbeh);



                Reportvalu persentmosharkat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-ارزیابی-درصد مشارکت فعال",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(persentmosharkat);


                Reportvalu updatesoratvatiaztpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatesoratvatiaztpymn);


                var counttt= db.tbSoratvaziat.Where(p => p.Finalvaluesorat != null &&p.Fk_pymn==itt.pec_ID).ToList();
                Reportvalu countpymnsorat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-تعداد ثبت شده",
                    Value = counttt.Count().ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(countpymnsorat);


                var countr = sorat.Where(p => p.Finalvaluesorat != null).ToList();
                Reportvalu tcounttaeedsorat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-تعداد تایید شده",
                    Value = countr.Count().ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(tcounttaeedsorat);


                Reportvalu jarysoart = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-سهم ریال جاری",
                    Value = jari.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(jarysoart);



                Reportvalu sarmayeh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-سهم ریال سرمایه ای",
                    Value = sarmayi.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sarmayeh);





                Reportvalu saranehkarkonan = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-سرانه کارکنان (درماه)",
                    Value = (pricesorat/ coun/ monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(saranehkarkonan);




                Reportvalu saranehmoshtrak = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-سرانه مشترک (درماه)",
                    Value = (pricesorat / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(saranehmoshtrak);



                Reportvalu saranehenerjy = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-سرانه انرژی",
                    Value = (pricesorat  / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(saranehenerjy);


                Reportvalu enhrafenerjy = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-انحراف مبلغ خام",
                    Value = (pricesorat   / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(enhrafenerjy);



                Reportvalu enherafvosol = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-انحراف از وصول",
                    Value = (pricesorat / coun / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(enherafvosol);



                Reportvalu emhraftalafat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-انحراف از تلفات",
                    Value = (pricesorat / coun / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(emhraftalafat);



                Reportvalu  enherafsabad= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-صورت وضعیت-انحراف سبد هزینه",
                    Value = (pricesorat / coun / monthDifference).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(enherafsabad);



                Reportvalu govahenamehpisghan = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghan);




                Reportvalu govahenamehpisghanemny = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-ایمنی فردی",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghanemny);




                Reportvalu mamorfany = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-مامور فنی مشترکین",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(mamorfany);




                Reportvalu gherat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-مامور فنی قرائت",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(gherat);




                Reportvalu govahenamehpisghanhavei = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-فن ورز هوایی",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghanhavei);




                Reportvalu govahenamehpisghangarm = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-فن ورز خط گرم",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghangarm);



                Reportvalu govahenamehpisghanshabakeh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-ایمنی در شبکه",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghanshabakeh);


                Reportvalu govahenamehpisghantestkontor = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-اصول تست کنتور",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghantestkontor);




                Reportvalu govahenamehpisghanghermojaz = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-شناسایی انشعاب غیر مجاز",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghanghermojaz);



                Reportvalu govahenamehpisghanadabmoshtrak = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-گواهینامه-آداب برخورد با مشترک",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(govahenamehpisghanadabmoshtrak);



                Reportvalu updatemeyar = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemarkaz);



                Reportvalu etbarmosolayt = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-اعتبار بیمه مسئولیت",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(etbarmosolayt);




                Reportvalu havadeth = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد حوادث جانی",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(havadeth);




                Reportvalu havadethamalyat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد حوادث عملیاتی",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(havadethamalyat);



                Reportvalu havadehcar = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد حوادث خودرویی",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(havadehcar);

                Reportvalu sazman = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد شکایات داخل سازمان",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sazman);







                Reportvalu gareghsazman = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد شکایات خارج سازمان",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(gareghsazman);




                Reportvalu enzebati = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد آرای کمیته انضباطی",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(enzebati);


                Reportvalu kamtArestandard = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد کارکنان کارکرد کمتر از استاندارد",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(kamtArestandard);



                Reportvalu bestarkarkonan = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-معیار-تعداد کارکنان کارکرد بیشتر از استاندارد",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemarkaz);




                Reportvalu updatemostanadat = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemostanadat);




                Reportvalu sabtamlkardpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت عملکرد پیمان",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sabtamlkardpymn);


                Reportvalu taeadamkardptmn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-تایید عملکرد پیمان",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(taeadamkardptmn);


                var Abzar = db.tbReffrenceSave.Where(p => p.FK_PeymanID == itt.pec_ID && p.IsFor == 4).FirstOrDefault();
                var Abzar2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == Abzar.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                var Abzar3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == Abzar2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year)&&p.Final_Accept==true).FirstOrDefault();
                var Abzar34 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == Abzar2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year)).FirstOrDefault();

                //var Abzar4 = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == Abzar3.ID && p.CountDays != 0 && p.Fk_pymn != null).ToList();
                if (Abzar34 !=null){
                    Reportvalu sabtabzar = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت ابزار پیمان",
                        Value = 1.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(sabtabzar);
                }
                else{
                    Reportvalu sabtabzar = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت ابزار پیمان",
                        Value = 0.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(sabtabzar);
                }
                if (Abzar3 != null)
                {
                    Reportvalu taedabzar = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-تایید ابزار پیمان",
                        Value = 1.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(taedabzar);
                }
                else
                {
                    Reportvalu taedabzar = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-تایید ابزار پیمان",
                        Value = 0.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(taedabzar);
                }


                var mashin = db.tbReffrenceSave.Where(p => p.FK_PeymanID == itt.pec_ID && p.IsFor == 3).FirstOrDefault();
                var mashin2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == mashin.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                var mashin3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == mashin2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year)&&p.Final_Accept==true).FirstOrDefault();
                var mashin34 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == mashin2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == month && s.Year == year)).FirstOrDefault();

                if (mashin34 != null)
                {
                    Reportvalu sabtaMASIN = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان",
                        Value = 1.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(sabtaMASIN);
                }
                else
                {
                    Reportvalu sabtaMASIN = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان",
                        Value = 0.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(sabtaMASIN);
                }
                if (mashin3 != null)
                {
                    Reportvalu taedMASHIN = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-تایید ماشین آلات پیمان",
                        Value = 1.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(taedMASHIN);
                }
                else
                {
                    Reportvalu taedMASHIN = new Reportvalu
                    {
                        Title = "پیشخوان-پیمان-مستندات وضعیت-تایید ماشین آلات پیمان",
                        Value = 0.ToString(),
                        Pymn_ID = itt.pec_ID,

                    };
                    fishValues.Add(taedMASHIN);
                }







                //Reportvalu govahenamehpisghan = new Reportvalu
                //{
                //    Title = "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان",
                //    Value = 0.ToString(),
                //    Pymn_ID = itt.pec_ID,

                //};
                //fishValues.Add(updatemarkaz);



            



                Reportvalu  projkala= new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-وضعیت پروژه کالایی",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(projkala);

                Reportvalu taminegtemaei = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-فیش واریز لیست تامین اجتماعی",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(taminegtemaei);




                Reportvalu malytpymn = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات وضعیت-فیش واریز لیست مالیات حقوق",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(malytpymn);


                Reportvalu updatefish = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تاریخ بروزرسانی",
                    Value = ToShamsi(DateTime.Now).ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatefish);
                int tamam = 0;
                int natamam = 0;
                int tedadfal = 0;
                int fahal = 0;
                foreach(var ito in db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == itt.pec_ID && p.Status == true).ToList())
                {
                    var mablagh = db.tbMoalefeValueFish.Where(p => p.FK_User==ito.FK_User_ID && p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)" && p.mlfvlfsh_Value > 0 && p.mlfvlfsh_Submit == true && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == month).FirstOrDefault();
                    if( mablagh != null)
                    {
                        tedadfal += 1;

                    }
                    //var mo = 4;
                    //var y = 1403;

                    // تاریخ مقایسه شده - 20 مارس 2025
                    //DateTime compareDate = new DateTime(2025, 3, 20);

                    // تبدیل سال هجری شمسی به میلادی
                    PersianCalendar pc = new PersianCalendar();
                    DateTime inputDate = pc.ToDateTime(year, month, 1, 0, 0, 0, 0);
                    var gh = db.tbUserContracts.Where(p => p.FK_UserID == ito.FK_User_ID).OrderByDescending(s => s.usc_ID).FirstOrDefault();
                    if (gh != null)
                    {
                        DateTime endDateMiladi = gh.usc_EndTime.Value;
                        DateTime compareDate = new DateTime(pc.GetYear(endDateMiladi), pc.GetMonth(endDateMiladi), pc.GetDayOfMonth(endDateMiladi), pc);

                        TimeSpan timeDifference = compareDate - inputDate;

                        // بررسی فاصله زمانی کمتر از 60 روز
                        if (Math.Abs(timeDifference.TotalDays) < 60)
                        {


                            natamam += 1;


                        } 
                        else if (Math.Abs(timeDifference.TotalDays) < 0)
                        {
                            tamam += 1;
                        }
                        else
                        {
                            fahal += 1;
                        }
                    }
                    // پیدا کردن سوابقی که فاصله زمانی کمتر از دو ماه دارند
               
                }

                Reportvalu afradhoghoig = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تعداد کارکنان فعال",
                    Value = fahal.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(afradhoghoig);




                Reportvalu nearend = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تعداد قراردادا پرسنلی درحال اتمام",
                    Value = natamam.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(nearend);



                Reportvalu end = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تعداد قرارداد پرسنلی منقضی شده",
                    Value = tamam.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(end);


                Reportvalu finishpardaght = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-وضعیت پرداخت حقوق ماه آخر",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(updatemarkaz);



                Reportvalu bank = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تاییدیه واریزی بانک حقوق ماه آاخر",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(bank);



                Reportvalu bemeh = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-ثبت لیست بیمه حقوق ماه آخر",
                    Value = 1.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(bemeh);


                Reportvalu malyatendmonth = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-ثبت لیست مالیات حقوق ماه آخر",
                    Value = 0.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
             
                fishValues.Add(malyatendmonth);
                int sabt = 0;
                int taedd = 0;
                if (month == 12)
                {
                    var taed = db.tbSavedFunctions.Where(p => p.tbPeymanContracts.pec_ID == itt.pec_ID && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Year == year+1 && s.MoalfeVal_Month == 1) && p.Final_Accept == true).FirstOrDefault();
                    var sav = db.tbSavedFunctions.Where(p => p.tbPeymanContracts.pec_ID == itt.pec_ID && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Year == year+1 && s.MoalfeVal_Month == 1)).FirstOrDefault();
                    if (taed != null)
                    {
                        taedd = 1;
                    }
                    if (sav != null)
                    {
                        sabt = 1;
                    }
                }
                else
                {
                    var taed = db.tbSavedFunctions.Where(p => p.tbPeymanContracts.pec_ID == itt.pec_ID && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Year == year && s.MoalfeVal_Month == month+1) && p.Final_Accept == true).FirstOrDefault();
                    var sav = db.tbSavedFunctions.Where(p => p.tbPeymanContracts.pec_ID == itt.pec_ID && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Year == year && s.MoalfeVal_Month == month+1)).FirstOrDefault();
                    if (taed != null)
                    {
                        taedd = 1;
                    }
                    if (sav != null)
                    {
                        sabt = 1;
                    }
                }




                Reportvalu sabtkarkard = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-ثبت کارکرد کارکنان حقوق ماه آینده",
                    Value = sabt.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(sabtkarkard);

                Reportvalu taeedkarkard = new Reportvalu
                {
                    Title = "پیشخوان-پیمان-مستندات حقوق-تایید کارکرد کارکنان حقوق ماه آینده",
                    Value = taedd.ToString(),
                    Pymn_ID = itt.pec_ID,

                };
                fishValues.Add(taeedkarkard);

                //Reportvalu govahenamehpisghan = new Reportvalu
                //{
                //    Title = "",
                //    Value = 0.ToString(),
                //    Pymn_ID = itt.pec_ID,

                //};
                //fishValues.Add(updatemarkaz);


                //Reportvalu govahenamehpisghan = new Reportvalu
                //{
                //    Title = "",
                //    Value = 0.ToString(),
                //    Pymn_ID = itt.pec_ID,

                //};
                //fishValues.Add(updatemarkaz);









            }













            List<tbMoalefeValuePishkhan> piskghan = new List<tbMoalefeValuePishkhan>();
            foreach (var item in fishValues)
            {
                var find = moalf.Where(p => p.md_Title == item.Title).FirstOrDefault();
                if (find != null)
                {
                    if (item.Usr_ID == 0)
                    {
                        tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan();
                        pis.mlfval_Value = item.Value;
                        pis.mlfval_Month = month;
                        pis.mlfval_Year = year;
                        pis.mlfval_FKUser = null;
                        pis.mlfval_FKMoalafeDastmozdi = find.md_ID;
                        pis.mlfval_FKPeyman = item.Pymn_ID;

                        piskghan.Add(pis);
                    }
                    else if (item.Pymn_ID == 0)
                    {
                        tbMoalefeValuePishkhan pis = new tbMoalefeValuePishkhan();
                        pis.mlfval_Value = item.Value;
                        pis.mlfval_Month = month;
                        pis.mlfval_Year = year;
                        pis.mlfval_FKUser = item.Usr_ID;
                        pis.mlfval_FKMoalafeDastmozdi = find.md_ID;
                        pis.mlfval_FKPeyman = null;

                        piskghan.Add(pis);

                    }
                   
                }

            }
            db.tbMoalefeValuePishkhan.AddRange(piskghan);
            db.SaveChanges();

            return View();
        }



        //public ActionResult ExportAddMoalefeExcel()
        //{
        //    var OutPutFile = SetDataExcel_AddMoalefe();
        //    string extension = ".xlsx";

        //    var stream = new MemoryStream();
        //    OutPutFile.Save(stream, extension);
        //    var mimeType = MimeTypes.ByExtension[extension];
        //    return File(stream.ToArray(), mimeType, "AddMoalefe" + extension);
        //}


        //public ExcelLibrary.SpreadSheet.Workbook reportbemeh(int month = 1, int year = 1403)
        //{
        //    var moalf = db.tbContractMoalefeDastmozdi.ToList();
        //    var usr=db.tbUsers.ToList();
        //    //foreach (var item in fish)
        //    //{
                
        //    //}
        //    var Moalefeexcelfile = ExcelLibrary.SpreadSheet.Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelgzareshvam.xlsx"));
        //    Row Row;
        //    var rc = db.tbinstallments.ToList();
        //    //name and lastname and groupjob
        //    int counter = 1;
        //    //var category = categoryRepo.GetAllActiveCategory();
        //    foreach(var it in usr)
        //    {
        //        var fish = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == month && p.mlfvlfsh_Year == year&&p.FK_User==it.usr_ID&&p.FK_EXCel==null).ToList();

        //        foreach (var item in fish)
        //        {
        //            Row = new Row() { Height = 20, Index = counter };
        //            //var t = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == item.installment_ID && p.Month == 4).FirstOrDefault();
        //            var  rozan = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "مزد گروه (شغل)").FirstOrDefault();
        //            var sanavat = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "مزد سنوات (ریال)").FirstOrDefault();
        //            var karkard = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز کارکرد ").FirstOrDefault();
        //            var mashmol = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع کل مشمول بیمه (ریال)").FirstOrDefault();
        //            var naghales = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "جمع ناخالص حقوق و مزایا (ریال)").FirstOrDefault();
        //            var ghales = fish.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "خالص قابل دریافت (ریال)").FirstOrDefault();

        //            double rozdastmozd = (double)((rozan.mlfvlfsh_Value ?? 0 + sanavat.mlfvlfsh_Value ?? 0) / karkard.mlfvlfsh_Value);
        //            double razmont = (double)((rozan.mlfvlfsh_Value ?? 0 + sanavat.mlfvlfsh_Value ?? 0));
        //            double mazayamashmol = (double)(mashmol.mlfvlfsh_Value - razmont);
        //            double mazamashmolandghayr = (double)(naghales.mlfvlfsh_Value);
        //            int haghbemeh = (int)(mashmol.mlfvlfsh_Value * 0.7);
        //            long sayrkosorat = (long)(mazamashmolandghayr-ghales.mlfvlfsh_Value - haghbemeh );



        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.FullName,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 0
        //            }
        //        });

        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.usr_Personal_ID,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 1
        //            }
        //        });
        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.usr_ID,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 2
        //            }
        //        });
        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.value,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 3
        //            }
        //        });

        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.Month,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 4
        //            }
        //        });
        //            Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.Year,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 5
        //            }
        //        });
        //            counter++;
        //            Moalefeexcelfile.Worksheets[0].AddRow(Row);
        //        }
        //    }
          
        //    //int counterunit = 1;
        //    //var unit = UnitRepo.ListActiveUnit();
        //    //foreach (var item in unit)
        //    //{
        //    //    Row = new Row() { Height = 20, Index = counterunit };


        //    //    Row.AddCells(new List<Cell>()
        //    //    {
        //    //        new Cell()
        //    //        {
        //    //            Value = item.unt_Name,
        //    //            FontFamily = "B Nazanin",
        //    //            Bold = false,
        //    //            Enable = true,
        //    //            Wrap = false,
        //    //            FontSize = 12,
        //    //            Italic = false,
        //    //            Underline = false,
        //    //            Index = 3
        //    //        }
        //    //    });
        //    //    counterunit++;
        //    //    Moalefeexcelfile.Sheets[1].AddRow(Row);
        //    //}
        //    return Moalefeexcelfile;
        //}




        //public ExcelLibrary.SpreadSheet.Workbook SetDataExcel_AddMoalefe()
        //{
        //    var Moalefeexcelfile = ExcelLibrary.SpreadSheet.Workbook.Load(Server.MapPath("~/Content/ExcelFiles/excelgzareshvam.xlsx"));
        //    Row Row;
        //    var rc = db.tbinstallments.ToList();
        //    //name and lastname and groupjob
        //    int counter = 1;
        //    //var category = categoryRepo.GetAllActiveCategory();
        //    foreach (var item in rc)
        //    {
        //        Row = new Row() { Height = 20, Index = counter };
        //        var t = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == item.installment_ID && p.Month == 4).FirstOrDefault();

        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.FullName,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 0
        //            }
        //        });

        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.usr_Personal_ID,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 1
        //            }
        //        });
        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = item.tbUsers.usr_ID,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 2
        //            }
        //        });
        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.value,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 3
        //            }
        //        });

        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.Month,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 4
        //            }
        //        });
        //        Row.AddCells(new List<Cell>()
        //        {
        //            new Cell()
        //            {
        //                Value = t.Year,
        //                FontFamily = "B Nazanin",
        //                Bold = false,
        //                Enable = true,
        //                Wrap = false,
        //                FontSize = 12,
        //                Italic = false,
        //                Underline = false,
        //                Index = 5
        //            }
        //        });
        //        counter++;
        //        Moalefeexcelfile.Worksheets[0].AddRow(Row);
        //    }
        //    //int counterunit = 1;
        //    //var unit = UnitRepo.ListActiveUnit();
        //    //foreach (var item in unit)
        //    //{
        //    //    Row = new Row() { Height = 20, Index = counterunit };


        //    //    Row.AddCells(new List<Cell>()
        //    //    {
        //    //        new Cell()
        //    //        {
        //    //            Value = item.unt_Name,
        //    //            FontFamily = "B Nazanin",
        //    //            Bold = false,
        //    //            Enable = true,
        //    //            Wrap = false,
        //    //            FontSize = 12,
        //    //            Italic = false,
        //    //            Underline = false,
        //    //            Index = 3
        //    //        }
        //    //    });
        //    //    counterunit++;
        //    //    Moalefeexcelfile.Sheets[1].AddRow(Row);
        //    //}
        //    return Moalefeexcelfile;
        //}








    }
}