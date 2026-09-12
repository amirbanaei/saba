using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.ViewModels.Budget.Reporter;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Stimulsoft.Report.StiRecentConnections;

namespace SaabWebProject.Areas.Budget.Controllers
{
    public class ReportController : Controller
    {
        SaabEntities db = new SaabEntities();

        // GET: Budget/Report
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult Detail(int id = 0, DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0)
        {
            Reportradif reportradif = new Reportradif();
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Pyman_ID == id && p.tbPeymanContracts.Inactive != true && p.Starttime <= Starttime && p.Endtime >= Endtime).FirstOrDefault();
            List<long> list = new List<long>();
            List<int> list2 = new List<int>();


            var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == c.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime && p.Peyman_ID == id && p.tbPeymanContracts.Inactive != true).ToList();
            foreach (var d in z)
            {
                if (d != null && d.Creditor != null)
                {
                    list2.Add((int)d.Creditor);
                    list.Add((long)d.Creditor);
                    reportradif.less.Add((long)d.Creditor);
                    reportradif.time.Add((DateTime)d.DataDocument);
                }
                else
                {
                    // Handle the case where Creditor is null, e.g., log a warning or perform some default action.
                }
            }
            reportradif.creator.Add((long)c.Creditorcredit);



            if (reportradif.less.Count > 0)
            {
                var sumLess = reportradif.less.Sum();
                var x = reportradif.creator[0] - sumLess;
                reportradif.lesssum = (x);

            }




            return View(reportradif);
        }




        public ActionResult Detail2(int id = 0, int FK_creditline = 0)
        {
            Reportradif reportradif = new Reportradif();
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Pyman_ID == id && p.tbPeymanContracts.Inactive != true).FirstOrDefault();
            List<long> list = new List<long>();
            List<int> list2 = new List<int>();


            var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == c.FK_creditline && p.Peyman_ID == id && p.tbPeymanContracts.Inactive != true).ToList();
            foreach (var d in z)
            {
                if (d != null && d.Creditor != null)
                {
                    list2.Add((int)d.Creditor);
                    list.Add((long)d.Creditor);
                    reportradif.less.Add((long)d.Creditor);
                    reportradif.time.Add((DateTime)d.DataDocument);
                }
                else
                {
                    // Handle the case where Creditor is null, e.g., log a warning or perform some default action.
                }
            }
            reportradif.creator.Add((long)c.Creditorcredit);



            if (reportradif.less.Count > 0)
            {
                var sumLess = reportradif.less.Sum();
                var x = reportradif.creator[0] - sumLess;
                reportradif.lesssum = (x);

            }




            return View("Detail", reportradif);
        }



        public ActionResult Detail3(int id = 0)
        {
            Reportradif2 reportradif = new Reportradif2();
            List<FinancialDocuments> documents = new List<FinancialDocuments>();
            var c = db.tbDetermining_creditline.Where(p => p.Pyman_ID == id).ToList();
            List<long> list = new List<long>();
            List<int> list2 = new List<int>();

            var z = db.FinancialDocuments.Where(p => p.Peyman_ID == id && p.tbPeymanContracts.Inactive != true).ToList();



            foreach (var d in z)
            {
                if (d != null && d.Creditor != null)
                {
                    list2.Add((int)d.Creditor);
                    var x = db.tbCreditIndicators.Where(x2 => x2.ID == d.Creditlamdicatros_ID).FirstOrDefault();
                    list.Add((long)d.Creditor);
                    reportradif.AddData((long)d.Creditor, (DateTime)d.DataDocument, x.Title);
                    //reportradif.less.Add((long)d.Creditor);
                    //reportradif.time.Add(d.DataDocument);
                }
                else
                {
                    // Handle the case where Creditor is null, e.g., log a warning or perform some default action.
                }
            }
            foreach (var item in c)
            {
                reportradif.creator.Add((long)item.Creditorcredit);
            }




            if (reportradif.less.Count > 0)
            {
                var sumLess = reportradif.less.Sum();
                var x = reportradif.creator[0] - sumLess;
                reportradif.lesssum = (x);

            }




            return View(reportradif);
        }
        public ActionResult Detail4(List<int> id = null, int FK_creditline = 0)
        {
            Reportradif reportradif = new Reportradif();
            List<tbDetermining_creditline> b = new List<tbDetermining_creditline>();
            foreach (var item in id)
            {
                var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Pyman_ID == item && p.tbPeymanContracts.Inactive != true).FirstOrDefault();
                b.Add(c);
            }
            var y = b.GroupBy(p => p.Pyman_ID).ToList();

            return View(b);
        }
        public ActionResult Detail5( DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0)
        {
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline  && p.Starttime <= Starttime && p.Endtime >= Endtime).ToList();
            List<FinancialDocuments> b = new List<FinancialDocuments>();
            foreach (var item in c)
            {
                var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == item.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime ).FirstOrDefault();
                b.Add(z);
            }
            return View("detail7",c);
        }
        public ActionResult detail7(DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0)
        {
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Starttime <= Starttime && p.Endtime >= Endtime).ToList();
            List<FinancialDocuments> b = new List<FinancialDocuments>();
            int totalCreditorCredit = 0; // Initialize total variable

            foreach (var item in c)
            {
                var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == item.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime).FirstOrDefault();
                b.Add(z);
                if (z != null)
                {
                    totalCreditorCredit +=(int) z.Creditor; // Accumulate the sum
                }
            }

            ViewBag.TotalCreditorCredit = totalCreditorCredit; // Pass the sum to the view
            return View(b);
        }
        public ActionResult detail77(DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0)
        {
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Starttime <= Starttime && p.Endtime >= Endtime).ToList();
            List<FinancialDocuments> b = new List<FinancialDocuments>();
            int totalCreditor = 0; // Initialize total variable

            foreach (var item in c)
            {
                var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == item.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime).FirstOrDefault();
                b.Add(z);
                if (z != null)
                {
                    totalCreditor += (int)z.Creditor; // Accumulate the sum
                }
            }

            int totalCreditorCredit = 0; // Initialize total variable

            foreach (var item in c)
            {
                var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == item.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime).FirstOrDefault();
                b.Add(z);
                if (z != null)
                {
                    totalCreditorCredit += (int)z.Creditor; // Accumulate the sum
                }
            }

            ViewBag.TotalCreditor = totalCreditor; // Pass the sum to the view
            ViewBag.TotalCreditorCredit = totalCreditorCredit; // Pass the sum to the view
            return View(b);
        }

        public ActionResult detail87(DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0)
        {
            var c = db.tbDetermining_creditline.Where(p => p.FK_creditline == FK_creditline && p.Starttime <= Starttime && p.Endtime >= Endtime).ToList();
            List<FinancialDocuments> b = new List<FinancialDocuments>();
            int totalCreditor = 0; // Initialize total variable for Creditor
            int totalCreditorCredit = 0; // Initialize total variable for Creditorcredit

            foreach (var item in c)
            {
                var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == item.FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime).FirstOrDefault();
                b.Add(z);
                if (z != null)
                {
                    totalCreditor += (int)z.Creditor; // Accumulate the sum for Creditor
                    totalCreditorCredit += (int)z.Creditor; // Accumulate the sum for Creditorcredit
                }
            }

            ViewBag.TotalCreditor = totalCreditor; // Pass the sum of Creditor to the view
            ViewBag.TotalCreditorCredit = totalCreditorCredit; // Pass the sum of Creditorcredit to the view
            return View(b);
        }
        public ActionResult viewusr(int us=0, DateTime Starttime = default, DateTime Endtime = default)
        {
            var z = db.FinancialDocuments.Where(p => p.User_ID==us && p.DataDocument >= Starttime && p.DataDocument <= Endtime).ToList();
            return View(z);
        }

        public ActionResult viewfilter()
        {
            return View();
        }
        public ActionResult Fileter(int us = 0, DateTime Starttime = default, DateTime Endtime = default, int select_Etebar22 = 0, float moblagh = 0)
        {

            var z = db.FinancialDocuments.Where(p => p.User_ID == us && p.DataDocument >= Starttime && p.DataDocument <= Endtime&&p.Debtore>= moblagh&&p.Creditor>= moblagh).ToList();
            if(select_Etebar22 == 3){
                return View(z);

            }
            else if(select_Etebar22 == 2)
            {
                return View("~/Areas/Budget/Views/Report/Viewbedahkarus.cshtml", z);

            }
            else
            {
                return View("~/Areas/Budget/Views/Report/viewfilterbestan.cshtml", z);

            }
        }
        public ActionResult Fileter2(int us = 0, DateTime Starttime = default, DateTime Endtime = default, int FK_creditline = 0, int select_Etebar22 = 0,float moblagh=0)
        {
            var z = db.FinancialDocuments.Where(p => p.Creditlamdicatros_ID == FK_creditline && p.DataDocument >= Starttime && p.DataDocument <= Endtime && p.Peyman_ID == us &&p.Creditor>= moblagh&&p.Debtore>= moblagh).ToList();
            if (select_Etebar22 == 3)
            {
                return View(z);

            }
            else if (select_Etebar22 == 2)
            {
                return View("~/Areas/Budget/Views/Report/Viewbedhkarpy.cshtml", z);

            }
            else
            {
                return View("~/Areas/Budget/Views/Report/Viewpestanpy.cshtml", z);

            }
        }





      
    }
}