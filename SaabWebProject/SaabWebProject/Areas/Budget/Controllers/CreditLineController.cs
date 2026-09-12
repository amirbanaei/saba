using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Budget;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using static Stimulsoft.Report.StiRecentConnections;

namespace SaabWebProject.Areas.Budget.Controllers
{
    public class CreditLineController : Controller 
    {
        tbCreditIndicatorsRepository crditRepo;
        SaabEntities db;
        tbDetermining_creditlineRepository rep_def_rsn = new tbDetermining_creditlineRepository();
        public CreditLineController()
        {
            db = new SaabEntities();
            crditRepo = new tbCreditIndicatorsRepository(db);
        }
        [AuthorizeAAA]
        public ActionResult Index()
        { 
            return View();
        }
        #region عناوین
        #region pages

        public ActionResult _CreateCreditIndiactors()
        {
            return View();
        }
        public ActionResult Edit_CreateCreditIndiactors(int id)
        {
            return View("~/Areas/Budget/Views/CreditLine/Edit_CreateCreditIndiactors.cshtml", crditRepo.Find(id));
        }
        public ActionResult Dtail_CreateCreditIndiactors(int id)
        {
            return View("~/Areas/Budget/Views/CreditLine/Dtail_CreateCreditIndiactors.cshtml", crditRepo.Find(id));
        }
        [HttpPost]
        public string Edit_CreateCreditIndiactors2(List<tbCreditIndicators> Filters)
        {
            return crditRepo.Update(Filters);
        }
        /// <summary>
        /// SELECT OF titles
        /// </summary>
        /// <returns></returns>
        [AuthorizeAAA]
        public ActionResult _ListAllCreditIndicatior()
        {
            var Model = crditRepo.Update();
            if (Model != null)
            {
                return PartialView("_ListAllCreditIndicatior", Model);
            }
            else
            {
                return PartialView("_ListAllCreditIndicatior", new tbCreditIndicators());
            }
        }
        #endregion
        #region methods
        [AuthorizeAAA]
        public ActionResult CreateCreditIndicators(string Title, string Description)
        {
            try
            {
                tbCreditIndicators credit = new tbCreditIndicators
                {
                    Description = Description,
                    Title = Title
                };
                if (crditRepo.Create(credit) == "True")
                {
                    return Content("1");
                }
                else
                {
                    return Content("0");
                }
            }
            catch (Exception)
            {

                return Content("0");
            }
        }

        public ActionResult _ListAllCreateCreditIndicators()
        {
            return View("~/Areas/Budget/Views/CreditLine/_ListAllCreateCreditIndicators.cshtml", crditRepo.List());
        }

        [AuthorizeAAA]
        public ActionResult _TableListOfCreditIndicators()
        {
            var Model = crditRepo.Update();
            if (Model != null)
            {
                return PartialView("_TableListOfCreditIndicators", Model);
            }
            else
            {
                return PartialView("_TableListOfCreditIndicators", new tbCreditIndicators());
            }
        }
        #endregion
        #endregion


        #region مدیریت اعتبار ردیف ها
        #region pages
        public ActionResult ManageValidityOfCredit()
        {
            return View();
        }
        public ActionResult _CreateValidityOfCredit()
        {
            return View();
        }
        public ActionResult editdate(int id=0)
        {
            return View(rep_def_rsn.Find(id));
        }

        public ActionResult tbDetermining_creditline_list()
        {
            var x = rep_def_rsn.List();
            List<tbDetermining_creditline> ne = new List<tbDetermining_creditline>();
            List<int> ne2 = new List<int>();
            List<DateTime> ne3 = new List<DateTime>();
            List<DateTime> ne4 = new List<DateTime>();



            foreach (var item in x)
            {
                //if(!ne3.Contains(item.Starttime) || !ne4.Contains(item.Endtime)|| !ne2.Contains(item.Pyman_ID))
                if (ne3.Contains(item.Starttime) && ne4.Contains(item.Endtime))
                {
                    if (!ne2.Contains(item.Pyman_ID)) { 
                 
                        ne.Add(item);
                        ne3.Add(item.Starttime);
                        ne4.Add(item.Endtime);
                        ne2.Add(item.Pyman_ID);
                    }

                }
                else
                {
                    ne.Add(item);
                    ne3.Add(item.Starttime);
                    ne4.Add(item.Endtime);
                    ne2.Add(item.Pyman_ID);
                }

            }


            return View("~/Areas/Budget/Views/CreditLine/tbDetermining_creditline_List.cshtml", ne);
        }
        public ActionResult detail_gharadad(int id)
        {
            var c=db.tbDetermining_creditline.Where(p=>p.Determining_creditline_ID==id).FirstOrDefault();
            var z= db.tbDetermining_creditline.Where(p=>p.Pyman_ID==c.Pyman_ID&&p.FK_creditline==c.FK_creditline && p.Starttime == c.Starttime && p.Endtime == c.Endtime).ToList();
            return View(z);

        }

        #endregion
        #region methods
        [AuthorizeAAA]
        public string tbDetermining_creditline_Add(List<tbDetermining_creditline> Filters)
        {
            long sum = 0;
            long x2 = 0;

            foreach (var item in Filters)
            {
                if (item.Creditorcredit.HasValue)
                {
                    sum += item.Creditorcredit.Value;
                }
            }

            var lastItem = Filters.LastOrDefault();
            if (lastItem != null)
            {
                var x = db.tbPeymanContracts.Where(p => p.pec_ID == lastItem.Pyman_ID&&p.Inactive!=true).FirstOrDefault();

                if (x != null)
                {
                    string formattedPrice = x.pec_Price.Replace(",", "");

                    long parsedPrice;
                    if (long.TryParse(formattedPrice, out parsedPrice))
                    {
                        x2 = parsedPrice;
                    }



                }
            }
            var x23 = db.tbPeymanElhaghie.Where(p => p.FK_PeymanID == lastItem.Pyman_ID && p.tbPeymanContracts.Inactive != true).ToList();
            if (x23.Count != 0)
            {
                foreach (var item in x23)
                {
                    if (item.IsLower == false)
                    {
                        string formattedPrice = item.Price.Replace(",", "");

                        long parsedPrice;
                        if (long.TryParse(formattedPrice, out parsedPrice))
                        {
                            x2 += parsedPrice;
                        }
                    }
                    else if (item.IsLower == true)
                    {
                        string formattedPrice = item.Price.Replace(",", "");

                        long parsedPrice;
                        if (long.TryParse(formattedPrice, out parsedPrice))
                        {
                            x2 -= parsedPrice;
                        }
                    }
                }
            }
            //if (sum > x2+20)
            //{
            //    return "مبلغ های های بستانکارهای شما بیش از حد مجاز در قرار داد است ";
            //}

            return rep_def_rsn.Create(Filters);

          

        }
        public ActionResult DeleteEquipment(int ID)
        {
            try
            {
                if (rep_def_rsn.Delete(ID))
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
        public ActionResult DeleteEquipment2(int ID)
        {
            var ex = db.tbDetermining_creditline.Where(p => p.Determining_creditline_ID == ID).FirstOrDefault();

            var c = db.tbDetermining_creditline.Where(p => p.Pyman_ID == ex.Pyman_ID&&p.tbPeymanContracts.Inactive!=true && p.Starttime == ex.Starttime && p.Endtime == ex.Endtime).ToList();
            foreach (var c2 in c)
            {
                try
                {
                    rep_def_rsn.Delete(c2.Determining_creditline_ID);
                   
                }
                catch (Exception)
                {

                    return Content("False");
                }
            }
            return Content("True");
        }
        public ActionResult tbDetermining_creditline_Edit(int id)
        {
            return View("~/Areas/Budget/Views/CreditLine/tbDetermining_creditline_Edit.cshtml", rep_def_rsn.Find(id));
        }
        public ActionResult tbDetermining_creditline_Detail(int id)
        {
            List<tbDetermining_creditline> ne = new List<tbDetermining_creditline>();
            List<int> ne2 = new List<int>();
            List<DateTime> ne3 = new List<DateTime>();
            List<DateTime> ne4 = new List<DateTime>();
            var ex = db.tbDetermining_creditline.Where(p => p.Determining_creditline_ID == id).FirstOrDefault();

            var c = db.tbDetermining_creditline.Where(p => p.Pyman_ID == ex.Pyman_ID&&p.tbPeymanContracts.Inactive!=true&&p.Starttime==ex.Starttime&&p.Endtime==ex.Endtime).ToList();
            foreach(var c2 in  c) {

                if (!ne2.Contains(c2.FK_creditline))
                {
                    var ex2 = db.tbDetermining_creditline.Where(p => p.FK_creditline==c2.FK_creditline&&p.Pyman_ID==ex.Pyman_ID&& p.tbPeymanContracts.Inactive != true).OrderByDescending(p => p.Determining_creditline_ID) // Replace YourDateProperty with the actual property you want to use for ordering
    .FirstOrDefault();

                    ne.Add(ex2);
                    ne2.Add(c2.FK_creditline);
                }
                }
            return View("~/Areas/Budget/Views/CreditLine/tbDetermining_creditline_Detail.cshtml", ne);
        }
        [HttpPost]
        public string tbDetermining_creditline_Edit2(List<tbDetermining_creditline> Filters)
        {
            float sum = 0;
            float x2 = 0;

            foreach (var filter in Filters)
            {
                var v = db.tbPeymanContracts.Where(p => p.pec_ID == filter.Pyman_ID && p.Inactive != true).FirstOrDefault();
                var w = db.tbDetermining_creditline
                    .Where(p => p.Starttime == filter.Starttime && p.Pyman_ID == filter.Pyman_ID&& p.tbPeymanContracts.Inactive != true && p.Endtime == filter.Endtime)
                    .ToList();

                foreach (var item in w)
                {
                    if (item.Creditorcredit.HasValue)
                    {
                        sum += item.Creditorcredit.Value;
                    }
                }

                // Convert filter.Creditorcredit (int?) to float before adding
                sum += filter.Creditorcredit.HasValue ? (float)filter.Creditorcredit.Value : 0;
            }

            var lastItem = Filters.LastOrDefault();
            if (lastItem != null)
            {
                var x = db.tbPeymanContracts.Where(p => p.pec_ID == lastItem.Pyman_ID && p.Inactive != true).FirstOrDefault();

                if (x != null)
                {
                    float parsedPrice;
                    if (float.TryParse(x.pec_Price, out parsedPrice))
                    {
                        x2 = parsedPrice;
                    }
                }
            }
            var x23 = db.tbPeymanElhaghie.Where(p => p.FK_PeymanID == lastItem.Pyman_ID && p.tbPeymanContracts.Inactive != true).ToList();
            if (x23.Count != 0)
            {
                foreach (var item in x23)
                {
                    if (item.IsLower == false)
                    {
                        string formattedPrice = item.Price.Replace(",", "");

                        long parsedPrice;
                        if (long.TryParse(formattedPrice, out parsedPrice))
                        {
                            x2 += parsedPrice;
                        }
                    }
                    else if (item.IsLower == true)
                    {
                        string formattedPrice = item.Price.Replace(",", "");

                        long parsedPrice;
                        if (long.TryParse(formattedPrice, out parsedPrice))
                        {
                            x2 -= parsedPrice;
                        }
                    }
                }
            }

            //if (sum+20 > x2)
            //{
            //    return "مبلغ های های بستانکارهای شما بیش از حد مجاز در قرار داد است ";
            //}
            //else
            //{
                // Assuming that rep_def_rsn.Create expects a single tbDetermining_creditline, not a list
                // You might need to iterate through the Filters list and call rep_def_rsn.Create for each item
                return rep_def_rsn.Create(Filters);
            
        }


        public string tbDetermining_creditline_Edit4( int Pyman_ID = 0, int Determining_creditline_ID = 0)
        {


            return "true";
        }
        public string tbDetermining_creditline_Edit3(int Pyman_ID = 0, DateTime Starttime = default, DateTime Endtime = default, int Determining_creditline_ID = 0)
        {
            // Retrieve a single record based on Pyman_ID and Determining_creditline_ID
            var x = db.tbDetermining_creditline.Where(p => p.Pyman_ID == Pyman_ID && p.tbPeymanContracts.Inactive != true && p.Determining_creditline_ID == Determining_creditline_ID).FirstOrDefault();

            if (x != null)
            {
                // Retrieve a list of records with the same Starttime, Endtime, and Pyman_ID as the 'x' record
                var c = db.tbDetermining_creditline.Where(s => s.Starttime == x.Starttime && s.Endtime == x.Endtime && s.Pyman_ID == x.Pyman_ID).ToList();

                foreach (var item in c)
                {
                    // Update the 'Starttime' and 'Endtime' properties
                    item.Starttime = Starttime;
                    item.Endtime = Endtime;

                    // Save changes to the database
                    db.SaveChanges();

                    // If you have a specific result to return, you can return it here
                }
                return "True";


                // If you don't return anything in the loop, you might want to return a default value or handle it accordingly
                //return "No matching records found.";
            }

            // If 'x' is null, handle it accordingly (e.g., return an error message)
            return "Record not found.";
        }



        public string installments_Edit(int Peyman_ID = 0, int Determining_creditline_ID = 0, DateTime Starttime = default, DateTime Endtime = default,
int User_ID = 0, int FK_creditline = 0, bool Calculation_type = true, int Creditorcredit = 0, int Debtorcredit = 0, string Title = "", int numberOfInstallment = 0,
IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {

            tbDetermining_creditline obj = new tbDetermining_creditline();
            obj.Determining_creditline_ID = Determining_creditline_ID;

                obj.Pyman_ID = Peyman_ID;
        

          

            obj.FK_creditline = FK_creditline;
            obj.Creditorcredit = Creditorcredit;
            obj.Debtorcredit = Debtorcredit;
            obj.Starttime = Starttime;
            obj.Endtime = Endtime;



            

                return rep_def_rsn.Update(obj).ToString();
         
        }
        #endregion
        #endregion





    }
}