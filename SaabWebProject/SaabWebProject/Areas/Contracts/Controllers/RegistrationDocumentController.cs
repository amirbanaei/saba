
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Documents;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Web.Spreadsheet;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using SaabWebProject.Models.ViewModels.Contracts.RigesterDocument;
using ExcelLibrary.BinaryFileFormat;
using System.Security.Cryptography;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class RegistrationDocumentController : Controller
    {
        SaabEntities db = new SaabEntities();

        #region ثبت مستندات دریافتی

        #region variables
        tbDocumentsAuthorizedUsersRepository rep_auth = new tbDocumentsAuthorizedUsersRepository();
        tbDocumentsDefinitionOfDocumentTitlesRepository rep_DefinitionOfDocumentTitles = new tbDocumentsDefinitionOfDocumentTitlesRepository();
        tbDocumentsDefiOfDoctTitles_UsersRepository rep_defOfDocTitle_user = new tbDocumentsDefiOfDoctTitles_UsersRepository();
        tbDocumentsRepository rep_Document = new tbDocumentsRepository();
        tbUsers onllineuser = new tbUsers();

        #endregion
        public RegistrationDocumentController()
        {
            
        }
        public ActionResult Index()
        {
            return View();
        }

        #region کاربران مجاز

        #region pages
        [AuthorizeAAA]
        public ActionResult authorizedUsers_Manage()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult authorizedUsers_Create()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult authorizedUsers_List()
        {
            var a = rep_auth.Update();
            return PartialView(a);
        }
        [AuthorizeAAA]
        public ActionResult authorizedUsers_Update(int ID)
        {
            var a = rep_auth.Find(ID);
            return View("~/Areas/Contracts/Views/RegistrationDocument/authorizedUsers_Update.cshtml", a);
        }
        [AuthorizeAAA]
        public ActionResult GetPeymansForThisUserAth()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        return View("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", rep_auth.GePeymansOfThisUSer(User.usr_ID));
                    } 
                }
            }
            return View("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList() /*new List<tbPeymanContracts>()*/);


        }
        [AuthorizeAAA]
        public ActionResult GetauthForThisPeyman(int PeymanID = 1)
        {

            return View("~/Areas/Users/Views/Users/_PartialViewUsersListWithSearchAndMultiple.cshtml", rep_auth.GetUserInThisPeymans(PeymanID));


        }
        [AuthorizeAAA]

        #endregion

        #region Events
        [AuthorizeAAA]
        public string authorizedUsers_Edit(List<SaabWebProject.Models.DomainModels.tbDocumentsAuthorizedUsers> things)
        {
            return rep_auth.Update(things);
        }
        #endregion
        #endregion

        #region تعریف عناوین مدارک
        #region pages
        [AuthorizeAAA]
        public ActionResult DefinitionOfDocumentTitles_Manage()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult DefinitionOfDocumentTitles_Create()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult DefinitionOfDocumentTitles_List()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        //var a = rep_DefinitionOfDocumentTitles.Listt(rep_auth.GePeymansOfThisUSer(User.usr_ID));
                        var a = rep_DefinitionOfDocumentTitles.Update();
                        return PartialView("~/Areas/Contracts/Views/RegistrationDocument/DefinitionOfDocumentTitles_List.cshtml", a);
                    }
                }
            }
            return PartialView("~/Areas/Contracts/Views/RegistrationDocument/DefinitionOfDocumentTitles_List.cshtml", new List<tbDocumentsDefinitionOfDocumentTitles>());
        }
        [AuthorizeAAA]
        public ActionResult DefinitionOfDocumentTitles_Update(int ID)
        {
            var a = rep_DefinitionOfDocumentTitles.Find(ID);
            if (a == null)
                a = new tbDocumentsDefinitionOfDocumentTitles();
            return View("~/Areas/Contracts/Views/RegistrationDocument/DefinitionOfDocumentTitles_Update.cshtml", a);
        }
        [AuthorizeAAA]
        public ActionResult GetAllTitlesForThisPeyman(int PeymanID)
        {
            return View(rep_DefinitionOfDocumentTitles.Listt(PeymanID));
        }
        #endregion

        #region events
        [AuthorizeAAA]
        public string DefinitionOfDocumentTitles_Save(List<tbDocumentsDefinitionOfDocumentTitles> things)
        {
            return rep_DefinitionOfDocumentTitles.Create(things);
        }
        [AuthorizeAAA]
        public string DefinitionOfDocumentTitles_Edit(tbDocumentsDefinitionOfDocumentTitles things)
        {
            return rep_DefinitionOfDocumentTitles.Update(things);
        }


        #endregion
        #endregion

        #region مجوز ثبت و دسترسی
        #region pages
        [AuthorizeAAA]
        public ActionResult LicenseRegistrationAndApproval_Manage()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult LicenseRegistrationAndApproval_Create()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult LicenseRegistrationAndApproval_List()
        {
            var result = rep_defOfDocTitle_user.Update();
            return PartialView(result);
        }
        #endregion

        #region Events
        [AuthorizeAAA]
        public string LicenseRegistrationAndApproval_Save(List<tbDocumentsDefiOfDoctTitles_Users> things)
        {
            return rep_defOfDocTitle_user.Create(things);
        }
        [AuthorizeAAA]
        public string LicenseRegistrationAndApproval_delete(int ID)
        {
            return rep_defOfDocTitle_user.Disable(ID).ToString();
        }
        #endregion
        #endregion

        #region ثبت مدارک
        #region pages
        [AuthorizeAAA]
        public ActionResult RegistrationOfDocuments_Manage()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult RegistrationOfDocuments_Create()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult RegistrationOfDocuments_List()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {

                            return PartialView(rep_Document.List_userSaver(User.usr_ID));

                        }
                        else
                        {
                            return PartialView();
                        }
                    }
                    catch (Exception E)
                    {

                        return PartialView();
                    }

                }
            }
            else
            {
                return PartialView();
            }
        }
        [AuthorizeAAA]
        public ActionResult GetPeymansForThisUser_Saver()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {
                            return View("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", rep_defOfDocTitle_user.GetPeymansForThisUser_Saver(User.usr_ID));

                        }
                        else
                        {
                            return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p=>p.Inactive!=true).ToList()/*new List<tbPeymanContracts>()*/);
                        }
                    }
                    catch (Exception E)
                    {

                        return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList());
                    }

                }
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList());
            }
        }
        [AuthorizeAAA]
        public ActionResult RegistrationOfDocuments_ListTitles(int Peyman_ID, int TypeOfDocument, int year, int month, string soorat_number)
        {
            List<tbDocuments> documents = new List<tbDocuments>();

            using (SaabEntities db = new SaabEntities())
            {
                var titles = db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.FK_PeymanID == Peyman_ID).ToList();

                if (TypeOfDocument == 1)//soorat vaziat
                {
                    documents = db.tbDocuments.Where(p =>p.TypeOfDocument==1 && p.FK_PeymanID == Peyman_ID && p.SooratNumber == soorat_number && p.Year == year).ToList();
                    foreach (var item in titles.Where(p => p.TypeOfDocument == 1).ToList())
                    {
                        if (documents.Where(p => p.FK_DDOfTitle == item.ID && p.TypeOfDocument == 1 ).FirstOrDefault() == null)
                        {
                            tbDocuments obj = new tbDocuments();
                            obj.FK_DDOfTitle = item.ID;
                            obj.FK_PeymanID = item.FK_PeymanID;
                            obj.Month = month;
                            obj.SooratNumber = soorat_number;
                            obj.TypeOfDocument = TypeOfDocument;
                            obj.Year = year;
                            documents.Add(obj);
                        }
                    }
                }
                else if (TypeOfDocument == 2)//hoghoogh o dastmozd
                {
                    documents = db.tbDocuments.Where(p => p.TypeOfDocument == 2 && p.FK_PeymanID == Peyman_ID && p.Month == month && p.Year == year).ToList();
                    foreach (var item2 in titles.Where(p => p.TypeOfDocument == 2).ToList())
                    {
                        if (documents.Where(p => p.FK_DDOfTitle == item2.ID && p.TypeOfDocument == 2).FirstOrDefault() == null)
                        {
                            tbDocuments obj = new tbDocuments();
                            obj.FK_DDOfTitle = item2.ID;
                            obj.FK_PeymanID = item2.FK_PeymanID;
                            obj.Month = month;
                            obj.SooratNumber = soorat_number;
                            obj.TypeOfDocument = TypeOfDocument;
                            obj.Year = year;
                            documents.Add(obj);
                        }
                    }
                }
                else
                {
                   
                    documents = db.tbDocuments.Where(p => p.TypeOfDocument == 3 && p.FK_PeymanID == Peyman_ID && p.Year == year).ToList();
                    foreach (var item3 in titles.Where(p => p.TypeOfDocument == 3).ToList())
                    {
                        if (documents.Where(p => p.FK_DDOfTitle == item3.ID && p.TypeOfDocument == 3).FirstOrDefault() == null)
                        {
                            tbDocuments obj = new tbDocuments();
                            obj.FK_DDOfTitle = item3.ID;
                            obj.FK_PeymanID = item3.FK_PeymanID;
                            obj.Month = month;
                            obj.SooratNumber = soorat_number;
                            obj.TypeOfDocument = TypeOfDocument;
                            obj.Year = year;
                            documents.Add(obj);
                        }
                    }
                }
               
            }
            return PartialView(documents);
        }
        #endregion
        #region events
        [AuthorizeAAA]
        public ActionResult Export_Excel_peyman_users(int Peyman_ID = 0)
        {
            var OutPutFile = SetDataExcel(Peyman_ID);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Users_peyman" + extension);
        }
        [AuthorizeAAA]
        private Workbook SetDataExcel(int peyman_ID)
        {
            var NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/peymanUser.xlsx"));
            Row Row;
            using (SaabEntities db = new SaabEntities())
            {
                var users = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.Status == true).Select(p => p.tbUsers);
                if (users.Count() > 0)
                {
                    int counter = 1;
                    foreach (var item in users)
                    {
                        Row = new Row() { Height = 20, Index = counter };

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
                            },
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
                            },
                            new Cell
                            {
                                Value = 0,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 2
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[0].AddRow(Row);
                    }
                }
            }

            return NamingTagssampleFile;
        }

        [AuthorizeAAA]
        public string UploadAttachmentFile(int ID = 0, int FK_DDOfTitle = 0,
            int FK_PeymanID = 0, int TypeOfDocument = 0, int Year = 0, int Month = 0, string SooratNumber = "",
            IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {
                            onllineuser = User;
                        }
                    }
                    catch
                    {

                    }

                }
            }

            tbDocuments obj = new tbDocuments();

            if (ID == 0)
            {
                obj.FK_DDOfTitle = FK_DDOfTitle;
                obj.FK_PeymanID = FK_PeymanID;
                obj.TypeOfDocument = TypeOfDocument;
                obj.Year = Year;
                obj.Month = Month;
                obj.SooratNumber = SooratNumber;
                obj.FK_UserID_Saver = onllineuser.usr_ID;
            }
            else
            {
                obj = rep_Document.Find(ID);
                obj.FK_UserID_Saver = onllineuser.usr_ID;

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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/RegistrationDucumentAttachmentsFile/" + filename));
                        obj.attachmentSystemName = filename;
                        obj.attachment = file.FileName;
                        if (ID == 0)
                        {
                            return rep_Document.Create(obj).ToString();
                        }
                        else
                        {
                            return rep_Document.Update(obj).ToString();
                        }
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

        [AuthorizeAAA]
        public string UploadSummaryFile(int ID = 0, int FK_DDOfTitle = 0,
          int FK_PeymanID = 0, int TypeOfDocument = 0, int Year = 0, int Month = 0, string SooratNumber = "",
          HttpPostedFileBase files = null, FormCollection data = null)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {
                            onllineuser = User;
                        }
                    }
                    catch
                    {

                    }

                }
            }
            tbDocuments obj = new tbDocuments();

            if (ID == 0)
            {
                obj.FK_DDOfTitle = FK_DDOfTitle;
                obj.FK_PeymanID = FK_PeymanID;
                obj.TypeOfDocument = TypeOfDocument;
                obj.Year = Year;
                obj.Month = Month;
                obj.SooratNumber = SooratNumber;
                obj.FK_UserID_Saver = onllineuser.usr_ID;

            }
            else
            {
                obj = rep_Document.Find(ID);
                obj.FK_UserID_Saver = onllineuser.usr_ID;

            }




            if (files != null)
            {
              

                    if (files.ContentLength > 0)
                    {
                        var segment = files.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                    files.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/RegistrationDucumentSummery/" + filename));
                        obj.SummaryDocumentSystemName = filename;
                        obj.SummaryDocument = files.FileName;
                        if (ID == 0)
                        {
                             rep_Document.Create(obj).ToString();
                        }
                        else
                        {
                             rep_Document.Update(obj).ToString();
                        }
                    }
                    else
                    {
                        return "True";
                    }
                








                int mlfvlfsh_Year2;
                float mlfvlfsh_Value;
                string name = "";
                var us = db.tbUsers.ToList();
                var molf = db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.ID == obj.FK_DDOfTitle).FirstOrDefault();
                var fishhh=db.tbMoalefeValueFish.Where(p=>p.mlfvlfsh_Month==Month&&p.mlfvlfsh_Year==Year&&p.FK_Moalefe== molf.FK_MoalefeControll).ToList();
                //var fishh = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Month == 4 && p.FK_EXCel != null && p.FK_Moalefe == 3243).ToList();
                List<tbExcelDocument> IsFish = new List<tbExcelDocument>();
                List<int> IsFish2 = new List<int>();

                var workbook = Telerik.Web.Spreadsheet.Workbook.Load(files.InputStream, Path.GetExtension(files.FileName));
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



                                    for (int i = 1; i < count; i++)
                                    {
                                        int mlfvlfsh_Year = 0;
                                        float mlfvlfsh_Year233 = 0;
                                        float mlfvlfsh_Year23 = 0;
                                        string nammm = "";
                                        var row = workbook.Sheets[0].Rows[i];
                                        var Name = row.Cells[1];
                                        if (Name.Value != null)
                                        {
                                            bool isNumeric = int.TryParse(Name.Value.ToString(), out mlfvlfsh_Year);

                                            if (!isNumeric)
                                            {
                                                return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                            }


                                        }


                                        var Name4 = row.Cells[2];
                                        if (Name4.Value != null)
                                        {
                                            bool isNumeric2 = float.TryParse(Name4.Value.ToString(), out mlfvlfsh_Year23);

                                            if (!isNumeric2)
                                            {
                                                return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                            }


                                        }
                                        //var Name5 = row.Cells[3];
                                        //if (Name5.Value != null)
                                        //{
                                        //    bool isNumeric3 = float.TryParse(Name5.Value.ToString(), out mlfvlfsh_Year233);

                                        //    if (!isNumeric3)
                                        //    {
                                        //        return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                        //    }


                                        //}

                                        //var Name6 = row.Cells[4];
                                        //if (Name6.Value != null)
                                        //{
                                        //    nammm = Name6.Value.ToString();



                                        //}






                                        //var Name3 = row.Cells[3];
                                        //if (Name2.Value != null)
                                        //{
                                        //    bool isNumeric3 = float.TryParse(Name3.Value.ToString(), out mlfvlfsh_Year23);

                                        //    if (!isNumeric3)
                                        //    {
                                        //        return "مقدار ستون نام در سطر " + " ( " + i + " ) " + "عددی نمی باشد";
                                        //    }


                                        //}
                                        int counntt = 0;
                                        float moghayrat = 0;
                                        tbExcelDocument ayab = new tbExcelDocument();

                                        var c2 = us.Where(p => p.usr_Personal_ID == mlfvlfsh_Year).FirstOrDefault();
                                        if (c2 != null)
                                        {
                                            var ty = fishhh.Where(p => p.FK_User == c2.usr_ID).FirstOrDefault();
                                            if (ty != null)
                                            {

                                                ayab.Fk_moalfe = molf.FK_MoalefeControll;
                                                ayab.Value = mlfvlfsh_Year23;
                                                ayab.USER_ID = c2.usr_ID;
                                                ayab.Year = Year;
                                                ayab.Month = Month;
                                                ayab.FK_DEocument = obj.ID;
                                                //ayab.Padash = 0;
                                                //ayab.Jarime = 0;
                                                if (ty.mlfvlfsh_Value != mlfvlfsh_Year23)
                                                {
                                                    ayab.value_moghayerat = ty.mlfvlfsh_Value;
                                                    counntt = 1;
                                                    IsFish2.Add(counntt);

                                                }
                                                else
                                                {
                                                    ayab.value_moghayerat = null;
                                                }
                                                IsFish.Add(ayab);

                                            }

                                            //IsFish.Add(ayab);
                                            //tbJaremehandpadash ja = new tbJaremehandpadash();
                                            //ja.Fk_usr = c2.usr_ID;
                                            //ja.Month = mlfvlfsh_Year23;
                                            //ja.Year = mlfvlfsh_Year233;

                                            //ja.Titel = nammm;

                                            //var ex = db.tbMoadelPadashJarimeAyab.Where(p => p.UserID == mlfvlfsh_Year).FirstOrDefault();

                                        }


                                    }
                                    db.tbExcelDocument.AddRange(IsFish);
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
                        if (IsFish2.Count != 0)
                        {
                            return "False";

                        }
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;

                    }
                }



                return "True";
            }
            else
            {
                return "True";
            }
        }


        public string deletee()
        {
          var del=  db.tbDocuments.OrderByDescending(s => s.ID).FirstOrDefault();
            if(del != null)
            {
                var delll = db.tbExcelDocument.Where(p => p.FK_DEocument == del.ID).ToList();
                List<tbExcelDocument>excelDocuments = new List<tbExcelDocument>();
                foreach (var item in delll)
                {
                    excelDocuments.Add(item);
                }
                db.tbExcelDocument.RemoveRange(excelDocuments);
                db.SaveChanges();
                db.tbDocuments.Remove(del);
                db.SaveChanges();

            }
            return "True";
        }


        /// <summary>
        
        /// </summary>
        /// <returns></returns>
        public ActionResult viewdetail()
        {
            List<tbExcelDocument>ex=new List<tbExcelDocument> ();
            //Rigseterview dn = new Rigseterview();
            var del = db.tbDocuments.OrderByDescending(s => s.ID).FirstOrDefault();
            if(del != null)
            {
                var molf = db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.ID == del.FK_DDOfTitle).FirstOrDefault();
                if(molf != null)
                {
                    var com = db.tbExcelDocument.Where(p => p.value_moghayerat != null && p.Month == del.Month && p.Year == del.Year && p.Fk_moalfe == molf.FK_MoalefeControll).ToList();

                    foreach (var item in com)
                    {
                        ex.Add(item);
                    }
                }

            }

            //if (del != null)
            //{ 



            //}
            return View(ex);
        }

        public ActionResult taedordelview(int id)
        {
            var fin=db.tbDocuments.Where(p=>p.ID== id).FirstOrDefault();    
            return View(fin);
        }
        public ActionResult taeedv(int id)
        {
            var fin = db.tbDocuments.Where(p => p.ID == id).FirstOrDefault();
            return View("~/Areas/Contracts/Views/RegistrationDocument/taeedv.cshtml", fin);
        }
        public ActionResult viewmoghart(int id)
        {
            List<tbExcelDocument> ex = new List<tbExcelDocument>();
            //Rigseterview dn = new Rigseterview();
            var del = db.tbDocuments.Where(s => s.ID== id).FirstOrDefault();
            if (del != null)
            {
                var molf = db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.ID == del.FK_DDOfTitle).FirstOrDefault();
                if (molf != null)
                {
                    var com = db.tbExcelDocument.Where(p => p.value_moghayerat != null && p.Month == del.Month && p.Year == del.Year && p.Fk_moalfe == molf.FK_MoalefeControll).ToList();

                    foreach (var item in com)
                    {
                        ex.Add(item);
                    }
                }

            }

            //if (del != null)
            //{ 



            //}
            return View(ex);
        }
        public ActionResult viesunaccept(int id)
        {
            var fin = db.tbDocuments.Where(p => p.ID == id).FirstOrDefault();
            return View(fin);
        }
        public ActionResult viewtaed(int id) {
            var fin = db.tbDocuments.Where(p => p.ID == id).FirstOrDefault();
            return View(fin);
        }

        public string deletvieww(int ID)
        {
          var d=  db.tbDocuments.Where(p => p.ID == ID).FirstOrDefault();
            var delll = db.tbExcelDocument.Where(p => p.FK_DEocument == d.ID).ToList();
            List<tbExcelDocument> excelDocuments = new List<tbExcelDocument>();
            foreach (var item in delll)
            {
                excelDocuments.Add(item);
            }
            db.tbExcelDocument.RemoveRange(excelDocuments);
            db.SaveChanges();
            db.tbDocuments.Remove(d);
            db.SaveChanges();
            return "True";
        }


        #endregion
        #endregion

        #region کنترل مدارک
        #region pages

        [AuthorizeAAA]
        public ActionResult ControlDocuments_Manage()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult ControlDocuments_List()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {
                            List<tbDocuments> list = new List<tbDocuments>();

                            foreach (var item in db.tbDocumentsDefiOfDoctTitles_Users.Where(p => p.FK_User == User.usr_ID && p.IsTaeed == true).ToList())
                            {
                                var thisID = item.tbDocumentsDefinitionOfDocumentTitles.ID;
                                list.AddRange(db.tbDocuments.Where(p => p.FK_DDOfTitle == thisID).ToList().Distinct().ToList());
                            }
                            foreach (var item in list)
                            {
                                item.peyman_Title = item.tbPeymanContracts.pec_Title;
                            }
                            return PartialView(list);

                        }
                        else
                        {
                            return PartialView();
                        }
                    }
                    catch (Exception E)
                    {

                        return PartialView();
                    }

                }
            }
            else
            {
                return PartialView();
            }


        }
        [AuthorizeAAA]
        public ActionResult ControlDocuments_Create()
        {
            return PartialView();
        }
        [AuthorizeAAA]
        public ActionResult GetPeymansForThisUser_Accepter()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var phone = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    try
                    {
                        var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == phone).FirstOrDefault();
                        if (User != null)
                        {
                            return View("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", rep_defOfDocTitle_user.GetPeymansForThisUser_Accepter(User.usr_ID));

                        }
                        else
                        {
                            return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList());
                        }
                    }
                    catch (Exception E)
                    {

                        return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList());
                    }

                }
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/Peyman/_GetAllListPeymans_Select.cshtml", db.tbPeymanContracts.Where(p => p.Inactive != true).ToList());
            }
        }

        #endregion
        #region events

        [AuthorizeAAA]
        public string acceptDocument(int doc_ID)
        {
            return rep_Document.Accepting(doc_ID , true);
        }[AuthorizeAAA]
        public string NotacceptDocument(int doc_ID)
        {
            return rep_Document.Accepting(doc_ID , false);
        }
        #endregion
        #endregion
        #endregion
    }
}