using Microsoft.Office.Interop.Excel;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries.Formula;
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Statments.Controllers
{
    public class StatmentController : Controller
    {
        tbDescriptionsoratRepository des = new tbDescriptionsoratRepository();
        SaabEntities db = new SaabEntities();


        #region Variable
        #endregion

        #region Constructor

        #endregion

        #region Pages
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult Report()
        {
            return View();
        }
        public ActionResult soratcondition()
        {
            
            return PartialView();
        }
        public ActionResult soratconditionBerjand()
        {
            return PartialView();
        }
    

        public ActionResult EjraMohasebatSuratVaziyat( int month = 12, int year = 1402) // اجرای قطعی محاسبات صورت وضعیت
        {
            List<tbFinalCheckSoorat> temp = new List<tbFinalCheckSoorat>();
            using (var db = new SaabEntities())
            {
                foreach (var item in  db.tbPeymanContracts.Where(p=>p.Inactive!=true).ToList())
            {
          
                    if (month == 1)
                    {
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item.pec_ID && p.Year == (year - 1) && p.Month == 12).FirstOrDefault());
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item.pec_ID && p.Year == year && p.Month == 1).FirstOrDefault());
                    }
                    else
                    {
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item.pec_ID && p.Year == year && p.Month == (month - 1)).FirstOrDefault());
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item.pec_ID && p.Year == year && p.Month == month).FirstOrDefault());
                    }
                }
                return PartialView(temp);
            }

        }
        public ActionResult EjraMohasebatSuratVaziyat_Azemayeshi(List<int> Fkpym = null, int month = 12, int year = 1402) // اجرای قطعی محاسبات صورت وضعیت
        {
            List<tbFinalCheckSoorat> temp = new List<tbFinalCheckSoorat>();

            foreach (var item in Fkpym.ToList())
            {
                using (var db = new SaabEntities())
                {
                    if (month == 1)
                    {
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item && p.Year == (year - 1) && p.Month == 12).FirstOrDefault());
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item && p.Year == year && p.Month == 1).FirstOrDefault());
                    }
                    else
                    {
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item && p.Year == year && p.Month == (month - 1)).FirstOrDefault());
                        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == item && p.Year == year && p.Month == month).FirstOrDefault());
                    }
                }
            }
               return PartialView(temp);

            //using (var db = new SaabEntities())
            //{
            //    List<tbFinalCheckSoorat> temp = new List<tbFinalCheckSoorat>();
            //    if (month == 1)
            //    {
            //        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == FK_Peyman && p.Year == (year - 1) && p.Month == 12).FirstOrDefault());
            //        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == FK_Peyman && p.Year == year && p.Month == 1).FirstOrDefault());
            //    }
            //    else
            //    {
            //        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == FK_Peyman && p.Year == year && p.Month == (month - 1)).FirstOrDefault());
            //        temp.Add(db.tbFinalCheckSoorat.Where(p => p.FK_Peyman == FK_Peyman && p.Year == year && p.Month == month).FirstOrDefault());
            //    }

            //    return PartialView(temp);
            //}

        }
        public ActionResult Deletesayr(int ID = 0)
        {
            var find=db.TbDescriptionDetail.Where(p=>p.ID == ID).FirstOrDefault();
            if (find != null)
            {
                var descrip=db.Descriptionsorat.Where(p=>p.ID == find.Fk_Description).FirstOrDefault();
                if (descrip != null)
                {
                    if (find.type == false)
                    {
                        long valu =(long) (descrip.value+find.value);
                        descrip.value=valu;
                        db.SaveChanges();

                    }
                    else if(find.type == true)
                    {
                        long valu = (long)(descrip.value - find.value);
                        descrip.value = valu;
                        db.SaveChanges();
                    }
                    db.TbDescriptionDetail.Remove(find);
                    db.SaveChanges();
                }

            }
            return Content("True");
        }
        public ActionResult EjraMohasebatSuratVaziyat_HiLo()
        {
            return PartialView();
        }
        public ActionResult show_Hialosorat(int fksabad=0,int numbersorat = 0)
        {
            var t = db.tbSoratvaziat.Where(p => p.FK_Sabad == fksabad && p.number_sorat == numbersorat).OrderBy(o=>o.value).ToList();
            return PartialView(t);
        }
        public string UploadFilDescriptiom(int ID = 0, IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Statments/Contents/TbDescriptionDetail/" + filename));
                        var peyman = Find(ID);
                        peyman.FileSystemName = filename;
                        peyman.FileName = file.FileName;
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
        public TbDescriptionDetail Find(int ID)
        {
            try
            {
                return db.TbDescriptionDetail.Find(ID);
            }
            catch
            {
                return null;
            }
        }
        public ActionResult SayerAmaliyateSuoratV()
        {
            return PartialView();
        }
        public ActionResult acceptSayerAmaliyat()
        {
            return PartialView();
        }
        public ActionResult bedkarha(int id=0)
        {
            var bedahkar = db.TbDescriptionDetail.Where(p=>p.Fk_Description == id&&p.type==false).ToList();
            return PartialView(bedahkar);
        }
        public ActionResult bestankarha(int id=0)
        {
            var bedahkar = db.TbDescriptionDetail.Where(p => p.Fk_Description == id && p.type == true).ToList();
            return PartialView(bedahkar);
        }
        public ActionResult finalaccepandedit(int id = 0)
        {
            var bedahkar = db.TbDescriptionDetail.Where(p => p.ID == id ).FirstOrDefault();

            return View(bedahkar);
        }
        public ActionResult SharheSayerAmaliyateSuoratV(int id=0,int fkpymn=0)
        {
            var ex=db.Descriptionsorat.Where(p=>p.FK_Pymn==fkpymn&&p.FK_Radif==id).FirstOrDefault();
         
                return PartialView(ex);
        }
        public ActionResult tablesorat(int id = 0, int fkpymn = 0)
        {
            var tt = db.Descriptionsorat.Where(p => p.FK_Pymn == fkpymn && p.FK_Radif == id).FirstOrDefault();
            if(tt != null)
            {
                var yu = db.TbDescriptionDetail.Where(p => p.Fk_Description == tt.ID).ToList();
                return PartialView(yu);

            }
            else
            {
                return PartialView(db.TbDescriptionDetail.Where(p => p.Descriptionsorat.FK_Pymn == fkpymn && p.Descriptionsorat.FK_Radif == id).ToList());
            }
        }

        public int createdes( Descriptionsorat Filters)
        {
            var existingDescription = db.Descriptionsorat
                .Where(p => p.FK_Radif == Filters.FK_Radif && p.FK_Pymn == Filters.FK_Pymn)
                .FirstOrDefault();

            if (existingDescription != null)
            {
                Filters.ID = existingDescription.ID; // Assign the ID from the existing description
               des.Update2(Filters);
                return Filters.ID;



                // Assuming Update2 method handles updating logic
            }
            else
            {
                return des.Create(Filters);
                // Assuming Create method handles creation logic
            }
        }

        public ActionResult pavast()
        {
            return View();
        }
        public ActionResult edittaeed(int id = 0, string titlee = "", int idd4 = 0, long value = 0) 
        {
            var fin=db.TbDescriptionDetail.Where(p=>p.ID == id).FirstOrDefault();
            if(fin != null)
            {
                var kol=db.Descriptionsorat.Where(p=>p.ID==fin.Fk_Description).FirstOrDefault();
                if (kol != null)
                {
                    if (fin.type == false)
                    {
                        kol.value +=fin.value;
                        if(idd4 == 0)
                        {
                            kol.value -= value;
                        }
                       else if (idd4 == 1)
                        {
                            kol.value += value;
                        }
                        db.SaveChanges();   
                    }
                   else if (fin.type == true)
                    {
                        kol.value -= fin.value;
                        if (idd4 == 0)
                        {
                            kol.value -= value;
                        }
                        else if (idd4 == 1)
                        {
                            kol.value += value;
                        }
                        db.SaveChanges();

                    }
                }
                if(idd4 == 0) {
                    fin.type = false;

                }
                else if(idd4 == 1)
                {
                    fin.type = true;

                }
                fin.value = value;
                fin.finalaccept = true;
                fin.Titleaccept = titlee;

                db.SaveChanges();

            }
            return Content("true");



        }
        public ActionResult accep(int id = 0)
        {
            var fin = db.TbDescriptionDetail.Where(p => p.ID == id).FirstOrDefault();
            if (fin != null)
            {
                fin.finalaccept = true;
                db.SaveChanges();

            }
            return Content("true");



        }

        public string createdetail(List<TbDescriptionDetail> Filter)
        {
        
            
            return des.Create3(Filter);
        }

        [AuthorizeAAA]
        public ActionResult Downloadsoorat(int year = 1403, int peymanID = 18, string Dore = "")
        {
            var Year = /*year;*/  DateTime.Now.GetShamsYear();
            using (var db = new SaabWebProject.Models.DomainModels.SaabEntities())
            {
                var peymanName = db.tbPeymanContracts.FirstOrDefault(p => p.pec_ID == peymanID && p.Inactive != true).pec_BriefTitle;
                if (db.tbPeymanContracts.FirstOrDefault().pec_Title == "آژانس خدمات مشترکین امور برق زاهدان 1") // پروژه سیستان
                {

                    var path = Server.MapPath("~/Areas/Statments/Views/Files/soorat_Sistan/" + Year + "/" + Dore + "/" + peymanName + ".pdf");
                    byte[] filebyyte = System.IO.File.ReadAllBytes(path);
                    return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "hilo.pdf");
                }
                else//استان
                {
                    var path = Server.MapPath("~/Areas/Statments/Views/Files/soorat_Ostan/" + Year + "/" + Dore + "/" + peymanName + ".pdf");
                    byte[] filebyyte = System.IO.File.ReadAllBytes(path);
                    return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, "soorat" + Dore + ".pdf");
                }
            }
        }
        public ActionResult _ReportSooratVaziat(int FK_Peyman , int number , int year) {
            using(var db = new SaabEntities())
            {
                var temp = db.tbFinantialStatements.Where(p => p.Number == number && p.FK_Peyman == FK_Peyman).FirstOrDefault();

            }
            return View();
        }
        #endregion

        #region Events
        #endregion


    }
}