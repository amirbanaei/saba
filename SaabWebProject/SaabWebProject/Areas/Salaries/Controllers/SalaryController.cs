using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Mvc;
using System.Web.Mvc; 
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.DomainModels; 
using System.Linq; 
using Stimulsoft.Report.Export;
using Stimulsoft.Report.Web;
using System;
using SaabWebProject.Models.Repositories.Ussers;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;
using SaabWebProject.Utility;

namespace SaabWebProject.Areas.Salaries.Controllers
{
    public class SalaryController : Controller
    { 
        public SaabEntities db { get; set; }
        public static int PersonalCode { get; set; }
        public static int Month { get; set; }
        public static int Year { get; set; }
        public static int Peyman { get; set; }
        tblinkUserAndPeymanRepository linkuserPeymanRepo;
        tbPeymanContractsRepository peymanRepo;
        tbUsersRepository UserRepo;
        public SalaryController()
        {
            db = new SaabEntities();
            linkuserPeymanRepo = new tblinkUserAndPeymanRepository(db);
            UserRepo = new tbUsersRepository(db);
            peymanRepo = new tbPeymanContractsRepository(db);
        }
        // GET: Salaries/Salary
        [AuthorizeAAA]
        public ActionResult EjraMohasebatHoghoogh()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult Report()
        {
            return View();
        }


        public ActionResult Report2()
        {
            return View();
        }
        public ActionResult Hilo()
        {
            return PartialView();
        }
        public ActionResult KosuratFish()
        {
            return PartialView();
        }
        public ActionResult SabteFish()
        {
            return PartialView();
        }


        [AuthorizeAAA]
        public ActionResult ReportPage()
        {
            return PartialView();
        }

        // [HttpPost]
        [AuthorizeAAA]
        public ActionResult DownloadFish(List<int> personelCode, int year = 1402, string month = "02", int peyman = 0)
            {
            string PeymanName = "";

            switch (month)
            {
                case "1":
                    month = "01";
                    break;
                case "2":
                    month = "02";
                    break;
                case "3":
                    month = "03";
                    break;
                case "4":
                    month = "04";
                    break;
                case "5":
                    month = "05";
                    break;
                case "6":
                    month = "06";
                    break;
                case "7":
                    month = "07";
                    break;
                case "8":
                    month = "08";
                    break;
                case "9":
                    month = "09";
                    break;
                default:
                    break;
            }
            //Ostan

            try
            {
                int? personal = null;
                List<string> result = new List<string>();
                foreach (var item in personelCode)
                {
                    if (item == -1)
                    {
                        var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                        if (cookie_user != null)
                        {
                            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                            var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                            personal = user.usr_Personal_ID;
                            PeymanName = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == user.usr_ID && p.Status == true).tbPeymanContracts.pec_BriefTitle;
                        }

                    }
                    else
                    {
                        personal = db.tbUsers.FirstOrDefault(p => p.usr_ID == item).usr_Personal_ID;
                        PeymanName = peymanRepo.Update().FirstOrDefault(p => p.pec_ID == peyman).pec_BriefTitle;

                    }
                    var path = Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + "/" + year + "/" + month);

                    DirectoryInfo directoryInfo = new DirectoryInfo(path);
                    FileInfo[] allPdFs = directoryInfo.GetFiles("*.pdf");
                    result.Add(allPdFs.Where(p => p.Name == PeymanName + '-' + personal + ".pdf").Select(p => p.FullName).FirstOrDefault());
                }
                using (PdfDocument targetDoc = new PdfDocument())
                {
                    foreach (string pdf in result)
                    {
                        using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < pdfDoc.PageCount; i++)
                            {
                                targetDoc.AddPage(pdfDoc.Pages[i]);
                            }
                        }
                    }
                    targetDoc.Save(Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + year + ".pdf"));

                }


                byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + year + ".pdf"));
                return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, year + "-" + month + "-" + "-" + personal + ".pdf");

            }
            catch (Exception ee)
            {
                return Content("0");
            }


            //Sistan


            //try
            //{
            //    int? personal = null;
            //    List<string> result = new List<string>();
            //    foreach (var item in personelCode)
            //    {
            //        if (item == -1)
            //        {
            //            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            //            if (cookie_user != null)
            //            {
            //                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
            //                personal = user.usr_Personal_ID;
            //                PeymanName = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == user.usr_ID&&p.Status==true).tbPeymanContracts.pec_BriefTitle;
            //            }

            //        }
            //        else
            //        {
            //            personal = db.tbUsers.FirstOrDefault(p => p.usr_ID == item).usr_Personal_ID;
            //            PeymanName = peymanRepo.Listt().FirstOrDefault(p => p.pec_ID == peyman).pec_BriefTitle;

            //        }
            //        var path = Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + "/" + year + "/" + month);

            //        DirectoryInfo directoryInfo = new DirectoryInfo(path);
            //        FileInfo[] allPdFs = directoryInfo.GetFiles("*.pdf");
            //        result.Add(allPdFs.Where(p => p.Name == PeymanName + '-' + personal + ".pdf").Select(p => p.FullName).FirstOrDefault());
            //    }
            //    using (PdfDocument targetDoc = new PdfDocument())
            //    {
            //        foreach (string pdf in result)
            //        {
            //            using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
            //            {
            //                for (int i = 0; i < pdfDoc.PageCount; i++)
            //                {
            //                    targetDoc.AddPage(pdfDoc.Pages[i]);
            //                }
            //            }
            //        }
            //        targetDoc.Save(Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + year + ".pdf"));

            //    }


            //    byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + year + ".pdf"));
            //    return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, year + "-" + month + "-" + "-" + PeymanName + ".pdf");

            //}
            //catch (Exception ee)
            //{
            //    return Content("0");
            //}

        }
        public ActionResult DownloadFish1(List<int> personelCode, int year = 1402, string month = "02", int usr = 0)
        {
            //var usr=db.tbUsers.Where(p=>p.usr_Personal_ID == personelCode.Count).FirstOrDefault();
            var peyman = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == usr).Select(p=>p.FK_Peyman_ID).FirstOrDefault();
            string PeymanName = "";

            switch (month)
            {
                case "1":
                    month = "01";
                    break;
                case "2":
                    month = "02";
                    break;
                case "3":
                    month = "03";
                    break;
                case "4":
                    month = "04";
                    break;
                case "5":
                    month = "05";
                    break;
                case "6":
                    month = "06";
                    break;
                case "7":
                    month = "07";
                    break;
                case "8":
                    month = "08";
                    break;
                case "9":
                    month = "09";
                    break;
                default:
                    break;
            }
            //Ostan

            try
            {
                int? personal = null;
                List<string> result = new List<string>();
                foreach (var item in personelCode)
                {
                    if (item == -1)
                    {
                        var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                        if (cookie_user != null)
                        {
                            var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                            var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                            personal = user.usr_Personal_ID;
                            PeymanName = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == user.usr_ID && p.Status == true).tbPeymanContracts.pec_BriefTitle;
                        }

                    }
                    else
                    {
                        personal = db.tbUsers.FirstOrDefault(p => p.usr_ID == item).usr_Personal_ID;
                        PeymanName = peymanRepo.Update().FirstOrDefault(p => p.pec_ID == peyman).pec_BriefTitle;

                    }
                    var path = Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + "/" + year + "/" + month);

                    DirectoryInfo directoryInfo = new DirectoryInfo(path);
                    FileInfo[] allPdFs = directoryInfo.GetFiles("*.pdf");
                    result.Add(allPdFs.Where(p => p.Name == PeymanName + '-' + personal + ".pdf").Select(p => p.FullName).FirstOrDefault());
                }
                using (PdfDocument targetDoc = new PdfDocument())
                {
                    foreach (string pdf in result)
                    {
                        using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < pdfDoc.PageCount; i++)
                            {
                                targetDoc.AddPage(pdfDoc.Pages[i]);
                            }
                        }
                    }
                    targetDoc.Save(Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + year + ".pdf"));

                }


                byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Salaries/Fish/Fish_Ostan/" + year + ".pdf"));
                return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, year + "-" + month + "-" + "-" + personal + ".pdf");

            }
            catch (Exception ee)
            {
                return Content("0");
            }


            //Sistan


            //try
            //{
            //    int? personal = null;
            //    List<string> result = new List<string>();
            //    foreach (var item in personelCode)
            //    {
            //        if (item == -1)
            //        {
            //            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            //            if (cookie_user != null)
            //            {
            //                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
            //                var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
            //                personal = user.usr_Personal_ID;
            //                PeymanName = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == user.usr_ID&&p.Status==true).tbPeymanContracts.pec_BriefTitle;
            //            }

            //        }
            //        else
            //        {
            //            personal = db.tbUsers.FirstOrDefault(p => p.usr_ID == item).usr_Personal_ID;
            //            PeymanName = peymanRepo.Listt().FirstOrDefault(p => p.pec_ID == peyman).pec_BriefTitle;

            //        }
            //        var path = Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + "/" + year + "/" + month);

            //        DirectoryInfo directoryInfo = new DirectoryInfo(path);
            //        FileInfo[] allPdFs = directoryInfo.GetFiles("*.pdf");
            //        result.Add(allPdFs.Where(p => p.Name == PeymanName + '-' + personal + ".pdf").Select(p => p.FullName).FirstOrDefault());
            //    }
            //    using (PdfDocument targetDoc = new PdfDocument())
            //    {
            //        foreach (string pdf in result)
            //        {
            //            using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
            //            {
            //                for (int i = 0; i < pdfDoc.PageCount; i++)
            //                {
            //                    targetDoc.AddPage(pdfDoc.Pages[i]);
            //                }
            //            }
            //        }
            //        targetDoc.Save(Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + year + ".pdf"));

            //    }


            //    byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Areas/Salaries/Fish/Fish_Sistan/" + year + ".pdf"));
            //    return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, year + "-" + month + "-" + "-" + PeymanName + ".pdf");

            //}
            //catch (Exception ee)
            //{
            //    return Content("0");
            //}

        }

        public ActionResult viewforsettingfish() {
            return View();
        
        }


        [AuthorizeAAA]
        public ActionResult GetReport()
        {
            var report = new StiReport();
            report.Load(Server.MapPath("~/Assets/Report_File/Fish_Hoghghi.mrt"));
            // Find the text field and change its value
            StiText firstname = report.GetComponentByName("txtFirstName") as StiText;
            StiText month = report.GetComponentByName("txtMonth") as StiText;
            StiText year = report.GetComponentByName("txtYear") as StiText;
            var User = db.tbUsers.Where(p => p.usr_Personal_ID == PersonalCode).FirstOrDefault();
            if (User != null)
            {
                firstname.Text.Value = User.FullName;
            }
            //var Company = db.tbPeymanContracts.Where(p => p.pec_ID == Peyman).FirstOrDefault();
            //if (Company != null)
            //{
            //    StiText com = report.GetComponentByName("txtPeyman") as StiText;
            //    com.Text = "<font face='B Yekan' size='14'>" + Company.pec_Title + "</font>";
            //    com.Text.

            //}
            month.Text.Value = Month.GetPersianMonth();
            year.Text.Value = Year.ToString();


            return StiMvcViewer.GetReportResult(report);
        }


        [AuthorizeAAA]
        public ActionResult ViewerEvent()

        {

            return StiMvcViewer.ViewerEventResult();

        }

        [AuthorizeAAA]
        public ActionResult _ListUsersOfPeyman(int idPeyman)
        {
            var linkuserpeyman = linkuserPeymanRepo.Update().Where(p => p.FK_Peyman_ID == idPeyman).Select(p => p.FK_User_ID).ToList();
            var Model = UserRepo.Update().Where(p => linkuserpeyman.Contains(p.usr_ID)).ToList();

            return PartialView("_ListUsersOfPeyman", Model);
        }

        //=======================================================================================
        [AuthorizeAAA]
        public ActionResult SettingFish()
        {
            return View();
        }
        public ActionResult SettingFish1()
        {
            return View();
        }
        //=======================================================================================
    }
}