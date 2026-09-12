using paymentReceipt.Models.DomainModels;
using paymentReceipt.Models.GenericRepositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace paymentReceipt.Controllers
{
    public class TaxController : Controller
    {
        // GET: Tax
        paymentReceiptEntities db;
        IRepository<tbTax> tbtaxRepo;
        IUnitofWorks unitofwork;
        public TaxController()
        {
            db = new paymentReceiptEntities();
            tbtaxRepo = new Repository<tbTax>(db);
            unitofwork = new UnitofWorks(db);
        }
        public ActionResult Add(tbTax tax)
        {
            tbtaxRepo.Insert(tax);
            unitofwork.SaveChange();
            return View();
        }
        public ActionResult Edit(tbTax tax)
        {
           
            var entry = tbtaxRepo.Get(p => p.tx_Step == tax.tx_Step).FirstOrDefault();
            entry.tx_TaxRatio = tax.tx_TaxRatio;
            entry.tx_YearTaxValue = tax.tx_YearTaxValue;
            entry.tx_MonthTaxValue = tax.tx_MonthTaxValue;
            entry.tx_MaximumYearSalary = tax.tx_MaximumYearSalary;
            entry.tx_MaximumMonthSalary = tax.tx_MaximumMonthSalary;
            tbtaxRepo.Update(entry);
            unitofwork.SaveChange();
            return View("Add");
            //return messagebox
        }
        public ActionResult Delete(int Step)
        {

            var Entry = tbtaxRepo.Get(p => p.tx_Step == Step).FirstOrDefault();
            tbtaxRepo.Delete(Entry);
            unitofwork.SaveChange();
            return View("Add");
            //return messagebox
        }
    }
}