using paymentReceipt.Models.DomainModels;
using paymentReceipt.Models.GenericRepositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace paymentReceipt.Controllers
{
    public class KarKardController : Controller
    {
        // GET: KarKard
        paymentReceiptEntities db;
        IRepository<tbKarkard> karkardRepo;
        IUnitofWorks unitofwork;
        public KarKardController()
        {
            db = new paymentReceiptEntities();
            karkardRepo = new Repository<tbKarkard>(db);
            unitofwork = new UnitofWorks(db);
        }
        public ActionResult Add(tbKarkard karkard)
        {
           
            karkardRepo.Insert(karkard);
            unitofwork.SaveChange();
            return View();
        }
        public ActionResult Edit(tbKarkard karkard)
        {
            
            var entry =db.tbKarkard.Where(p => p.krkrd_Title == karkard.krkrd_Title).FirstOrDefault();
            entry.krkrd_Title = karkard.krkrd_Title;
            entry.krkrd_StandardValue = karkard.krkrd_StandardValue;
            entry.krkrd_Duration = karkard.krkrd_Duration;
            karkardRepo.Update(entry);
            unitofwork.SaveChange();
            return View("Add");
            //return messagebox
        }
        public ActionResult Delete(string title)
        {
            
            var Entry = db.tbKarkard.FirstOrDefault(p => p.krkrd_Title == title);
            karkardRepo.Delete(Entry);
            unitofwork.SaveChange();
            return View("Add");
            //return messagebox
        }
    }
}