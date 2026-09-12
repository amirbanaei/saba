using paymentReceipt.Models.DomainModels;
using paymentReceipt.Models.GenericRepositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace paymentReceipt.Controllers
{
    public class FishController : Controller
    {
        // GET: Fish
        paymentReceiptEntities db;
        IRepository<tbFish> tbfishRepo;
        IUnitofWorks unitofwork;
        public FishController()
        {
            db = new paymentReceiptEntities();
            tbfishRepo = new Repository<tbFish>(db);
            unitofwork = new UnitofWorks(db);
        }
        public ActionResult Add(tbFish fish)
        {
            tbfishRepo.Insert(fish);
            unitofwork.SaveChange();
            return View();

        }
    }
}