
using Telerik.Web.Spreadsheet;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Utility;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using System.Diagnostics.Metrics;
using System.Web.Services.Description;

namespace SaabWebProject.Areas.Users.Controllers
{
    public class CompanyController : Controller
    {
        SaabEntities db = new SaabEntities();
        tbCompaniesRepository reo_company = new tbCompaniesRepository();
        tbUsersRepository rep_users = new tbUsersRepository();

        [AuthorizeAAA]
        public ActionResult ManageComanies()
        {
            return View("~/Areas/Users/Views/Company/ManageComanies.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult _PartialListComanies()
        {
            return PartialView("~/Areas/Users/Views/Company/_PartialListComanies.cshtml", reo_company.Update().OrderByDescending(p => p.ID));
        }

        [AuthorizeAAA]
        public ActionResult _PartialUpdateCompanies(int id)
        {
            return PartialView("~/Areas/Users/Views/Company/_PartialUpdateCompanies.cshtml", reo_company.Find(id));

        }
        public ActionResult _PartialdetailCompanies(int id)
        {
            return PartialView("~/Areas/Users/Views/Company/_PartialdetailCompanies.cshtml", reo_company.Find(id));

        }
        [AuthorizeAAA]
        public ActionResult _PartialCreateCompany()
        {
            return PartialView("~/Areas/Users/Views/Company/_PartialCreateCompany.cshtml", new tbCompanies());
        }

        [AuthorizeAAA]
        public ActionResult _PartialListCompaniesInDropDown()
        {
            var a = reo_company.Update();
            return PartialView("~/Areas/Users/Views/Company/_PartialListCompaniesInDropDown.cshtml", a);
        }
        [HttpPost]
        [AuthorizeAAA]
        public string Company_Update(tbCompanies tbCompanies, tbUsers users,int fk)
        {
            tbCompanies.FK_ManagerID = fk;
            tbCompanies.IsActive = true;
            tbCompanies.tbUsers.usr_IsUser = false;
            return reo_company.Update(tbCompanies);
        }
        [HttpPost]
        [AuthorizeAAA]
        public string Company_Create(tbCompanies companies, tbUsers tbUsers)
        {
            companies.IsActive = true;
            //companies.tbUsers.usr_IsUser = false;
            return reo_company.Create(companies);
        }

        [AuthorizeAAA]
        public bool Disable_Company(int id)
        {
            return reo_company.ActiveOrNotActiveCompany(id);
        }

        [AuthorizeAAA]
        public ActionResult _GetAllCompanies_Select()
        {
            var a = reo_company.ActiveList();
            return PartialView("_GetAllCompanies_Select", a);
        }


        [AuthorizeAAA]
        public ActionResult DownloadSampleExcel()
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            using (SaabEntities db = new SaabEntities())
            {
                var workbook = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/SampleExcelCompany.xlsx"));
                var worksheet = workbook.Sheets[1];
                var Cities = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
                var usr = db.tbUsers.Where(p=>p.usr_Personal_ID!=null).ToList();
                var Counter = 1;
                var Counter2 = 1;

                var Counter3 = 1;
                var Counter4 = 1;

                foreach (var city in Cities)
                {
                    var row = new Row();
                    row.Index = Counter;
                    var cell = new Cell();
                    cell.Index = 0;
                    cell.Value = city.cityName;
                    row.AddCell(cell);
                    worksheet.AddRow(row);
                    Counter++;
                }
                foreach (var city in usr)
                {
                    var row = new Row();
                    row.Index = Counter2;
                    var cell = new Cell();
                    cell.Index = 1;
                    cell.Value = city.FullName;
                    row.AddCell(cell);
                    worksheet.AddRow(row);
                    Counter2++;
                }
                foreach (var city in usr)
                {
                    var row = new Row();
                    row.Index = Counter3;
                    var cell = new Cell();
                    cell.Index = 2;
                    cell.Value = city.usr_Personal_ID;
                    row.AddCell(cell);
                    worksheet.AddRow(row);
                    Counter3++;
                }
                foreach (var city in usr)
                {
                    var row = new Row();
                    row.Index = Counter4;
                    var cell = new Cell();
                    cell.Index = 3;
                    cell.Value = city.usr_PhoneNumber;
                    row.AddCell(cell);
                    worksheet.AddRow(row);
                    Counter4++;
                }

                workbook.Save(stream, ".xlsx");


            }
            

            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SampleExcelCompany.xlsx");

        }
        [HttpPost]


        public ActionResult GETEXCEL(HttpPostedFileBase MyExcelStream)
        {
            List<string> lstMoalefe = new List<string>();
            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;


            ;
            string Message = "#";

            List<tbCompanies> listcom = new List<tbCompanies>();

            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();


                            if (count == 1)
                            {
                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                            }

                            for (int i = 13; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }



                            var Counter = 1;

                            for (int i = 1; i < count-10; i++)
                            {
                                tbCompanies companies = new tbCompanies();
                                tbUsers user = new tbUsers();
                                var row = workbook.Worksheets[0].Rows[i];
                                var Name = row.Cells[0];

                                var Family = row.Cells[0];
                                if (Family != null)
                                {
                                    if (int.TryParse(Family.Value.ToString(), out int result))
                                    {
                                        var t = db.tbUsers.Where(p => p.usr_Personal_ID == result).FirstOrDefault();
                                        companies.FK_ManagerID = t.usr_ID;
                                    }
                                }

                                //var PhoneNumber = row.Cells[2];
                                //if (PhoneNumber != null)
                                //{
                                //    if (!string.IsNullOrEmpty(PhoneNumber.Value))
                                //    {
                                //        user.usr_PhoneNumber = System.Convert.ToInt64(PhoneNumber.Value);
                                //    }
                                //}
                                //var Email = row.Cells[3];
                                //if (Email != null)
                                //{
                                //    if (!string.IsNullOrEmpty(Email.Value))
                                //    {
                                //        user.usr_Email = Email.Value;
                                //    }
                                //}
                                var ExclusiveCode = row.Cells[2];
                                if (ExclusiveCode != null)
                                {
                                    if (!string.IsNullOrEmpty(ExclusiveCode.Value))
                                    {
                                        companies.ExclusiveCode = System.Convert.ToInt64(ExclusiveCode.Value);
                                    }
                                    else
                                    {
                                        Message = "لطفا کد اختصاصی شرکت سطر " + i + " را وارد کنید";
                                        TempData["Message"] = Message;
                                        return Content(Message);

                                    }
                                }
                                var CompanyName = row.Cells[3];
                                if (CompanyName != null)
                                {
                                    if (!string.IsNullOrEmpty(CompanyName.Value))
                                    {
                                        companies.CompanyName = CompanyName.Value;
                                    }
                                    else
                                    {
                                        Message = "لطفا نام شرکت سطر " + i + " را وارد کنید";
                                        TempData["Message"] = Message;
                                        return Content("لطفا نام شرکت سطر " + i + " را وارد کنید");

                                    }
                                }
                                var RegisterationNumber = row.Cells[4];
                                if (RegisterationNumber != null)
                                {
                                    if (!string.IsNullOrEmpty(RegisterationNumber.Value))
                                    {
                                        companies.RegistrationNumber = System.Convert.ToInt64(RegisterationNumber.Value);
                                    }
                                }
                                var PleaceOfRegister = row.Cells[5];
                                if (PleaceOfRegister != null)
                                {
                                    if (!string.IsNullOrEmpty(PleaceOfRegister.Value))
                                    {
                                        var Cities = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
                                        var City = Cities.Where(p => p.cityName == PleaceOfRegister.Value).FirstOrDefault();
                                        if (City != null)
                                        {
                                            companies.PlaceOfRegister = City.cityId;
                                        }
                                        else
                                        {
                                            Message = "لطفا محل ثبت سطر " + i + " را وارد کنید";
                                            TempData["Message"] = Message;
                                            return Content("لطفا محل ثبت سطر " + i + " را وارد کنید");


                                        }
                                    }
                                }
                                var CompanyNational = row.Cells[6];
                                if (CompanyNational != null)
                                {
                                    if (!string.IsNullOrEmpty(CompanyNational.Value))
                                    {
                                        companies.Company_National_ID = System.Convert.ToInt32(CompanyNational.Value);
                                    }
                                }
                                var Enconomic_Code = row.Cells[7];
                                if (Enconomic_Code != null)
                                {
                                    if (!string.IsNullOrEmpty(Enconomic_Code.Value))
                                    {
                                        companies.Enconomic_Code = System.Convert.ToInt64(Enconomic_Code.Value);
                                    }
                                }
                                var Telephone1 = row.Cells[8];
                                if (Telephone1 != null)
                                {
                                    if (!string.IsNullOrEmpty(Telephone1.Value))
                                    {
                                        companies.Telephone_1 = System.Convert.ToInt64(Telephone1.Value);
                                    }
                                }
                                var Telephone2 = row.Cells[9];
                                if (Telephone2 != null)
                                {
                                    if (!string.IsNullOrEmpty(Telephone2.Value))
                                    {
                                        companies.Telephone_2 = System.Convert.ToInt64(Telephone2.Value);
                                    }
                                }
                                var Fax = row.Cells[10];
                                if (Fax != null)
                                {
                                    if (!string.IsNullOrEmpty(Fax.Value))
                                    {
                                        companies.Fax_Number = System.Convert.ToInt64(Fax.Value);
                                    }
                                }
                                var Site = row.Cells[11];
                                if (Site != null)
                                {
                                    if (!string.IsNullOrEmpty(Site.Value))
                                    {
                                        companies.Url_site = Site.Value;
                                    }
                                }
                                var CompanyEmail = row.Cells[12];
                                if (CompanyEmail != null)
                                {
                                    if (!string.IsNullOrEmpty(CompanyEmail.Value))
                                    {
                                        companies.Company_Email = CompanyEmail.Value;
                                    }
                                }
                                var Address = row.Cells[13];
                                if (Address != null)
                                {
                                    if (!string.IsNullOrEmpty(Address.Value))
                                    {
                                        companies.Company_Address = Address.Value;
                                    }
                                }

                                user.usr_IsActive = true;
                                //companies.tbUsers = user;
                                companies.IsActive = true;
                                listcom.Add(companies);

                            }
                            db.tbCompanies.AddRange(listcom);
                            db.SaveChanges();
                        }
                    }
                }
                catch (Exception ex)
                {
                    return Content(ex.Message);
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return Content("True");

            //return Content("False");
            //var Model33 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

        }













        public ActionResult UploadExcel(HttpPostedFileBase excelFile)
        {
            List<tbCompanies> listcom=new List<tbCompanies>();


            string Message = "#";
            using (SaabEntities db = new SaabEntities())
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        using (ExcelEngine xl = new ExcelEngine())
                        {
                            IApplication app = xl.Excel;
                            app.Workbooks.Open(excelFile.InputStream);
                            IWorkbook workbook = app.ActiveWorkbook;
                            IWorksheet ws = workbook.Worksheets[0];

                            var Rows = ws.Rows.ToList();

                            Rows.Remove(Rows.FirstOrDefault());
                            var Counter = 1;
                            foreach (var row1 in Rows)
                            {

                                var title = workbook.Worksheets[0].Rows[0].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return Content("فایل اکسل فاقد اطلاعات می باشد");
                                }
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];

                                    tbCompanies companies = new tbCompanies();
                                    tbUsers user = new tbUsers();
                                    //var Name = row.Cells[0];
                                    //if (Name != null)
                                    //{
                                    //    if (!string.IsNullOrEmpty(Name.Value))
                                    //    {
                                    //        user.usr_Name = Name.Value;
                                    //    }
                                    //}



                                    var Family = row.Cells[0];
                                    if (Family != null)
                                    {
                                        if (int.TryParse(Family.Value.ToString(), out int result))
                                        {
                                            var t = db.tbUsers.Where(p => p.usr_Personal_ID == result).FirstOrDefault();
                                            companies.FK_ManagerID = t.usr_ID;
                                        }
                                    }

                                    //var PhoneNumber = row.Cells[2];
                                    //if (PhoneNumber != null)
                                    //{
                                    //    if (!string.IsNullOrEmpty(PhoneNumber.Value))
                                    //    {
                                    //        user.usr_PhoneNumber = System.Convert.ToInt64(PhoneNumber.Value);
                                    //    }
                                    //}
                                    //var Email = row.Cells[3];
                                    //if (Email != null)
                                    //{
                                    //    if (!string.IsNullOrEmpty(Email.Value))
                                    //    {
                                    //        user.usr_Email = Email.Value;
                                    //    }
                                    //}
                                    var ExclusiveCode = row.Cells[2];
                                    if (ExclusiveCode != null)
                                    {
                                        if (!string.IsNullOrEmpty(ExclusiveCode.Value))
                                        {
                                            companies.ExclusiveCode = System.Convert.ToInt64(ExclusiveCode.Value);
                                        }
                                        else
                                        {
                                            Message = "لطفا کد اختصاصی شرکت سطر " + Counter + " را وارد کنید";
                                            TempData["Message"] = Message;
                                            return Content(Message);

                                        }
                                    }
                                    var CompanyName = row.Cells[3];
                                    if (CompanyName != null)
                                    {
                                        if (!string.IsNullOrEmpty(CompanyName.Value))
                                        {
                                            companies.CompanyName = CompanyName.Value;
                                        }
                                        else
                                        {
                                            Message = "لطفا نام شرکت سطر " + Counter + " را وارد کنید";
                                            TempData["Message"] = Message;
                                            return Content("لطفا نام شرکت سطر " + Counter + " را وارد کنید");

                                        }
                                    }
                                    var RegisterationNumber = row.Cells[4];
                                    if (RegisterationNumber != null)
                                    {
                                        if (!string.IsNullOrEmpty(RegisterationNumber.Value))
                                        {
                                            companies.RegistrationNumber =System.Convert.ToInt64(RegisterationNumber.Value);
                                        }
                                    }
                                    var PleaceOfRegister = row.Cells[5];
                                    if (PleaceOfRegister != null)
                                    {
                                        if (!string.IsNullOrEmpty(PleaceOfRegister.Value))
                                        {
                                            var Cities = rep_users.GetAllCity(Server.MapPath("~/Areas/Users/Data/Cities.json"));
                                            var City = Cities.Where(p => p.cityName == PleaceOfRegister.Value).FirstOrDefault();
                                            if (City != null)
                                            {
                                                companies.PlaceOfRegister = City.cityId;
                                            }
                                            else
                                            {
                                                Message = "لطفا محل ثبت سطر " + Counter + " را وارد کنید";
                                                TempData["Message"] = Message;
                                                return Content("لطفا محل ثبت سطر " + Counter + " را وارد کنید");


                                            }
                                        }
                                    }
                                    var CompanyNational = row.Cells[6];
                                    if (CompanyNational != null)
                                    {
                                        if (!string.IsNullOrEmpty(CompanyNational.Value))
                                        {
                                            companies.Company_National_ID = System.Convert.ToInt32(CompanyNational.Value);
                                        }
                                    }
                                    var Enconomic_Code = row.Cells[7];
                                    if (Enconomic_Code != null)
                                    {
                                        if (!string.IsNullOrEmpty(Enconomic_Code.Value))
                                        {
                                            companies.Enconomic_Code = System.Convert.ToInt64(Enconomic_Code.Value);
                                        }
                                    }
                                    var Telephone1 = row.Cells[8];
                                    if (Telephone1 != null)
                                    {
                                        if (!string.IsNullOrEmpty(Telephone1.Value))
                                        {
                                            companies.Telephone_1 = System.Convert.ToInt64(Telephone1.Value);
                                        }
                                    }
                                    var Telephone2 = row.Cells[9];
                                    if (Telephone2 != null)
                                    {
                                        if (!string.IsNullOrEmpty(Telephone2.Value))
                                        {
                                            companies.Telephone_2 = System.Convert.ToInt64(Telephone2.Value);
                                        }
                                    }
                                    var Fax = row.Cells[10];
                                    if (Fax != null)
                                    {
                                        if (!string.IsNullOrEmpty(Fax.Value))
                                        {
                                            companies.Fax_Number = System.Convert.ToInt64(Fax.Value);
                                        }
                                    }
                                    var Site = row.Cells[11];
                                    if (Site != null)
                                    {
                                        if (!string.IsNullOrEmpty(Site.Value))
                                        {
                                            companies.Url_site = Site.Value;
                                        }
                                    }
                                    var CompanyEmail = row.Cells[12];
                                    if (CompanyEmail != null)
                                    {
                                        if (!string.IsNullOrEmpty(CompanyEmail.Value))
                                        {
                                            companies.Company_Email = CompanyEmail.Value;
                                        }
                                    }
                                    var Address = row.Cells[13];
                                    if (Address != null)
                                    {
                                        if (!string.IsNullOrEmpty(Address.Value))
                                        {
                                            companies.Company_Address = Address.Value;
                                        }
                                    }

                                    user.usr_IsActive = true;
                                    //companies.tbUsers = user;
                                    companies.IsActive = true;
                                    listcom.Add(companies);
                                }
                                break;

                            }

                            Counter++;
                        }
                        transaction.Commit();
                    }
                    catch(Exception ex)
                    {
                        transaction.Rollback();
                        TempData["Message"] = "آپلود فایل با خطا مواجه شد";
                      RedirectToAction("ManageComanies");
                        return Content("یاخطا مواجه شده است ");
                    }
                }
                db.tbCompanies.AddRange(listcom);
                db.SaveChanges();
            }

            TempData["Message"] = Message;
             RedirectToAction("ManageComanies");
            return Content("true");

        }
    }
}