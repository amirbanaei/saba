using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using SaabWebProject.Models.ViewModels.BusinessSide;
using SaabWebProject.Utility;
using Telerik.Web.Spreadsheet;
using MimeTypes = Telerik.Web.Spreadsheet.MimeTypes;
using Syncfusion.XlsIO;

namespace SaabWebProject.Areas.Users.Controllers
{
    public class BusinessSideController : Controller
    {
        #region متغیر ها

        private BusinessSideRepository business_SideRepository;
        private SaabEntities db;

        public BusinessSideController()
        {
            db = new SaabEntities();
            business_SideRepository = new BusinessSideRepository(db);
        }

        #endregion

        #region صفحات

        [AuthorizeAAA]
        public ActionResult Index(string message = null)
        {
            return View();
        }

        /// <summary>
        /// پارشیال ویو جدول لیست پوزیشن ها
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _GetAllBusiness()
        {
            var res = business_SideRepository.Update().OrderByDescending(p => p.up_Id).ToList();
            return PartialView("~/Areas/Users/Views/BusinessSide/_GetAllBusiness.cshtml", res);
        }

        /// <summary>
        /// پارشیال ویو مدال افزودن پوزیشن
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        [HttpGet]
        public ActionResult _CreateBusinessSide()
        {
            return PartialView("~/Areas/Users/Views/BusinessSide/_CreateBusinessSide.cshtml");
        }

        /// <summary>
        /// پارشیال ویو مربوط به ویرایش سمت
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        [HttpGet]
        public ActionResult _EditBusinessSide(int id)
        {
            var tbBusinnes = business_SideRepository.Find(id);
            return PartialView("~/Areas/Users/Views/BusinessSide/_EditBusinessSide.cshtml", tbBusinnes);
        }

        /// <summary>
        /// پارشیال ویو لیست سمت ها در قالب سلکت آپشن
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _GetAllBusinessSide_Select()
        {
            var model = business_SideRepository.GetAllBusinessSidesForSelect();

            return PartialView("~/Areas/Users/Views/BusinessSide/_GetAllBusinessSide_Select.cshtml", model);
        }
        public ActionResult _GetAllBusinessSide_SelectModiriyatFaaliat()
        {
            var model = business_SideRepository.GetAllBusinessSidesForSelect();

            return PartialView("~/Areas/Users/Views/BusinessSide/_GetAllBusinessSide_SelectModiriyatFaaliat.cshtml", model);
        }
        [AuthorizeAAA]
        public ActionResult TreeChart_BusinessSide()
        {
            StringBuilder builder = new StringBuilder();

            var users = db.tbUsers.ToList();

            foreach (var item in users)
            {
                foreach (var link in item.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
                {
                    var pos = link.tbBusinessSide;

                    if (pos != null)
                    {
                        string parent = "";
                        if (pos.FK_Parent_ID == null)
                        {
                            parent = "{\"key\":" + pos.up_Id + "," + "\"name\":\"" + pos.up_name + "\"" + "," + "\"title\":" + "\"" + item.FullName + "\"" + "," + "\"pic\":" + "\"" + item.usr_Picture + "\"" + "},";
                        }
                        else
                        {
                            parent = "{\"key\":" + pos.up_Id + "," + "\"name\":\"" + pos.up_name + "\"" + "," + "\"title\":" + "\"" + item.FullName + "\"" + "," + "\"pic\":" + "\"" + item.usr_Picture + "\"" + "," + "\"parent" + "\":" + pos.FK_Parent_ID + "},";
                        }

                        builder.Append(parent);
                    }
                }
            }

            if (builder.Length > 0)
            {
                builder.Replace(",", "", (builder.Length - 1), 1);
            }

            return View(builder);
        }

        #endregion

        #region توابع

        /// <summary>
        /// تابع افزودن سمت و زیرمجموعه آنها
        /// </summary>
        /// <param name="obj">نمونه ای از جدول پوزیشن ها</param>
        /// <returns></returns>
        [HttpPost]
        [AuthorizeAAA]
        public ActionResult _CreateBusinessSide(tbBusinessSide obj)
        {
            try
            {
                var result = business_SideRepository.Create(obj);
                if (result == "Exists_Position")
                {
                    TempData["response"] = "Exists_Position";
                }
                else
                {
                    TempData["response"] = Convert.ToBoolean(result).ToString();
                }
                return RedirectToAction("Index", "BusinessSide", new {area="Users"});
            }
            catch
            {
                TempData["response"] = "catch";
                return RedirectToAction("Index", "BusinessSide", new {area="Users"});
            }
        }

        /// <summary>
        /// تابع ویرایش سمت و زیر مجموعه آنها
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        [HttpPost]
        [AuthorizeAAA]
        public ActionResult _EditBusinessSide(tbBusinessSide obj)
        {
            try
            {
                string res = "";
                if (obj.up_Id == obj.FK_Parent_ID)
                {
                    res = "Invalid-Parent";
                    return Content(res);
                }

                res = business_SideRepository.Update(obj);
                TempData["response"] = Convert.ToBoolean(res);
                return RedirectToAction("Index", "BusinessSide", new {area="Users"});
            }
            catch
            {
                TempData["response"] = "catch";
                return RedirectToAction("Index", "BusinessSide", new {area="Users"});
            }
        }
        public ActionResult listjob()
        {
            var job = db.tbjob.ToList();
            return View(job);
        }

        public ActionResult listjobdetal()
        {
            var job = db.tbdetailjob.ToList();
            return View(job);
        }


        /// <summary>
        /// تابع مربوط به غیرفعال کردن یا فعال کردن سمت
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        [AuthorizeAAA]
        public bool DeActiveOrActive(int id)
        {
            var res = business_SideRepository.ActiveOrDeActive(id);
            return res;
        }


        [HttpGet]
        [AuthorizeAAA]
        public ActionResult Export_BusinessSide()
        {
            var OutPutFile = SetDataExcel_BusinessSide();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "BusinessSides" + extension);
        }

        private Workbook SetDataExcel_BusinessSide()
        {
            var NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/PositionsSample.xlsx"));
            Row Row;
            using (SaabEntities db = new SaabEntities())
            {
                var jobInfo = db.tbBusinessSide.ToList();
                if (jobInfo.Count() > 0)
                {
                    int counter = 1;
                    foreach (var item in jobInfo)
                    {
                        Row = new Row() { Height = 20, Index = counter };

                        Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.up_name,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 0
                            }
                        });
                        counter++;
                        NamingTagssampleFile.Sheets[1].AddRow(Row);
                    }
                }
            }

            return NamingTagssampleFile;
        }

        [HttpPost]
        [AuthorizeAAA]
        public string GetDataFromExcel_Positions(HttpPostedFileBase BusinessFile)
        {
            using (ExcelEngine xl = new ExcelEngine())
            {
                IApplication app = xl.Excel;
                var Work = app.Workbooks.Open(BusinessFile.InputStream);

                var sheets = Work.Worksheets;
                var cols = sheets[0].Columns.Count();
                var Rows = sheets[0].Rows;
                if (cols != 15 || sheets.Count() != 2)
                {
                    return "اکسل وارد شده با اکسل نمونه متفاوت است";
                }

                var rows = sheets[0].Rows.ToList();
                rows.Remove(rows[0]);
                List<tbBusinessSide> tbBusinessSides = new List<tbBusinessSide>();
                var counter = 1;
                foreach (var row in rows)
                {
                    var position = row.Cells[0].Value;
                    if (!string.IsNullOrEmpty(position))
                    {
                        var dbPos = db.tbBusinessSide.Where(p => p.up_name == position).FirstOrDefault();
                        if (dbPos != null)
                        {
                            dbPos.up_name = position;

                            var PositionParent = row.Cells[1].Value;
                            if (!string.IsNullOrEmpty(PositionParent))
                            {
                                var parent = db.tbBusinessSide.FirstOrDefault(p => p.up_name == PositionParent);
                                if (parent != null)
                                {
                                    dbPos.FK_Parent_ID = parent.up_Id;
                                }
                                else
                                {
                                    return "ستون سمت بالا دستی سطر " + " ( " + counter + " ) " + " را خالی بگذارید یا درست وارد کنید ";
                                }
                            }


                            var up_Duties = row.Cells[2].Value;
                            if (!string.IsNullOrEmpty(up_Duties))
                            {
                                dbPos.up_Duties = up_Duties;
                            }

                            var up_workTime = row.Cells[3].Value;
                            if (!string.IsNullOrEmpty(up_workTime))
                            {
                                dbPos.up_workTime = up_workTime;
                            }

                            var up_experience = row.Cells[4].Value;
                            if (!string.IsNullOrEmpty(up_experience))
                            {
                                dbPos.up_experience = up_experience;
                            }

                            var up_Safety_Skills = row.Cells[5].Value;
                            if (!string.IsNullOrEmpty(up_Safety_Skills))
                            {
                                dbPos.up_Safety_Skills = up_Safety_Skills;
                            }

                            var up_Special_Skills = row.Cells[6].Value;
                            if (!string.IsNullOrEmpty(up_Special_Skills))
                            {
                                dbPos.up_Special_Skills = up_Special_Skills;
                            }

                            var up_General_Skills = row.Cells[7].Value;
                            if (!string.IsNullOrEmpty(up_General_Skills))
                            {
                                dbPos.up_General_Skills = up_General_Skills;
                            }

                            var up_em_Gender = row.Cells[8].Value;
                            if (!string.IsNullOrEmpty(up_em_Gender))
                            {
                                switch (up_em_Gender)
                                {
                                    case "مرد":
                                    {
                                        dbPos.up_em_Gender = 2;
                                        break;
                                    }
                                    case "زن":
                                    {
                                        dbPos.up_em_Gender = 1;
                                        break;
                                    }
                                    case "هردو":
                                    {
                                        dbPos.up_em_Gender = 3;
                                        break;
                                    }
                                }

                                if (dbPos.up_em_Gender == null)
                                {
                                    return "ستون جنسیت سطر " + " ( " + counter + " ) " + " را طبق فرمت وارد کیند ";
                                }
                            }

                            var up_minimum_Edu = row.Cells[9].Value;
                            if (!string.IsNullOrEmpty(up_minimum_Edu))
                            {
                                switch (up_minimum_Edu)
                                {
                                    case "سیکل":
                                    {
                                        dbPos.up_minimum_Edu = 1;
                                        break;
                                    }
                                    case "دیپلم":
                                    {
                                        dbPos.up_minimum_Edu = 2;
                                        break;
                                    }
                                    case "فوق دیپلم":
                                    {
                                        dbPos.up_minimum_Edu = 3;
                                        break;
                                    }
                                    case "لیسانس":
                                    {
                                        dbPos.up_minimum_Edu = 4;
                                        break;
                                    }
                                    case "فوق لیسانس":
                                    {
                                        dbPos.up_minimum_Edu = 5;
                                        break;
                                    }
                                    case "دکترا":
                                    {
                                        dbPos.up_minimum_Edu = 6;
                                        break;
                                    }
                                    case "هردو":
                                    {
                                        dbPos.up_minimum_Edu = 7;
                                        break;
                                    }
                                    case "فوق دکترا":
                                    {
                                        dbPos.up_minimum_Edu = 8;
                                        break;
                                    }
                                }

                                if (dbPos.up_minimum_Edu == null)
                                {
                                    return "ستون حداقل تحصیلات سطر " + " ( " + counter + " ) " + " را طبق فرمت وارد کیند ";
                                }
                            }

                            var up_General_Equipment = row.Cells[10].Value;
                            if (!string.IsNullOrEmpty(up_General_Equipment))
                            {
                                dbPos.up_General_Equipment = up_General_Equipment;
                            }

                            var up_Specialized_Equipment = row.Cells[11].Value;
                            if (!string.IsNullOrEmpty(up_Specialized_Equipment))
                            {
                                dbPos.up_Specialized_Equipment = up_Specialized_Equipment;
                            }

                            var up_work_Space = row.Cells[12].Value;
                            if (!string.IsNullOrEmpty(up_work_Space))
                            {
                                dbPos.up_work_Space = up_work_Space;
                            }

                            var up_Related_Jobs = row.Cells[13].Value;
                            if (!string.IsNullOrEmpty(up_Related_Jobs))
                            {
                                dbPos.up_Related_Jobs = up_Related_Jobs;
                            }

                            var up_Authority = row.Cells[14].Value;
                            if (!string.IsNullOrEmpty(up_Authority))
                            {
                                dbPos.up_Authority = up_Authority;
                            }
                        }
                        else
                        {
                            var pos = new tbBusinessSide();
                            if (!string.IsNullOrEmpty(position))
                            {
                                pos.up_name = position;
                            }
                            else
                            {
                                return "ستون عنوان سمت سطر " + " ( " + counter + " ) " + " را وارد کیند ";
                            }

                            var PositionParent = row.Cells[1].Value;
                            if (!string.IsNullOrEmpty(PositionParent))
                            {
                                var parent = db.tbBusinessSide.Where(p => p.up_name == PositionParent).FirstOrDefault();
                                if (parent != null)
                                {
                                    pos.FK_Parent_ID = parent.up_Id;
                                }
                                else
                                {
                                    return "ستون سمت بالا دستی سطر " + " ( " + counter + " ) " + " را خالی بگذارید یا درست وارد کنید ";
                                }
                            }

                            var up_Duties = row.Cells[2].Value;
                            if (!string.IsNullOrEmpty(up_Duties))
                            {
                                pos.up_Duties = up_Duties;
                            }

                            var up_workTime = row.Cells[3].Value;
                            if (!string.IsNullOrEmpty(up_workTime))
                            {
                                pos.up_workTime = up_workTime;
                            }

                            var up_experience = row.Cells[4].Value;
                            if (!string.IsNullOrEmpty(up_experience))
                            {
                                pos.up_experience = up_experience;
                            }

                            var up_Safety_Skills = row.Cells[5].Value;
                            if (!string.IsNullOrEmpty(up_Safety_Skills))
                            {
                                pos.up_Safety_Skills = up_Safety_Skills;
                            }

                            var up_Special_Skills = row.Cells[6].Value;
                            if (!string.IsNullOrEmpty(up_Special_Skills))
                            {
                                pos.up_Special_Skills = up_Special_Skills;
                            }

                            var up_General_Skills = row.Cells[7].Value;
                            if (!string.IsNullOrEmpty(up_General_Skills))
                            {
                                pos.up_General_Skills = up_General_Skills;
                            }

                            var up_em_Gender = row.Cells[8].Value;
                            if (!string.IsNullOrEmpty(up_em_Gender))
                            {
                                switch (up_em_Gender)
                                {
                                    case "مرد":
                                    {
                                        pos.up_em_Gender = 2;
                                        break;
                                    }
                                    case "زن":
                                    {
                                        pos.up_em_Gender = 1;
                                        break;
                                    }
                                    case "هردو":
                                    {
                                        pos.up_em_Gender = 3;
                                        break;
                                    }
                                }

                                if (pos.up_em_Gender == null)
                                {
                                    return "ستون جنسیت سطر " + " ( " + counter + " ) " + " را طبق فرمت وارد کیند ";
                                }
                            }

                            var up_minimum_Edu = row.Cells[9].Value;
                            if (!string.IsNullOrEmpty(up_minimum_Edu))
                            {
                                switch (up_minimum_Edu)
                                {
                                    case "سیکل":
                                    {
                                        pos.up_minimum_Edu = 1;
                                        break;
                                    }
                                    case "دیپلم":
                                    {
                                        pos.up_minimum_Edu = 2;
                                        break;
                                    }
                                    case "فوق دیپلم":
                                    {
                                        pos.up_minimum_Edu = 3;
                                        break;
                                    }
                                    case "لیسانس":
                                    {
                                        pos.up_minimum_Edu = 4;
                                        break;
                                    }
                                    case "فوق لیسانس":
                                    {
                                        pos.up_minimum_Edu = 5;
                                        break;
                                    }
                                    case "دکترا":
                                    {
                                        pos.up_minimum_Edu = 6;
                                        break;
                                    }
                                    case "هردو":
                                    {
                                        pos.up_minimum_Edu = 7;
                                        break;
                                    }
                                    case "فوق دکترا":
                                    {
                                        pos.up_minimum_Edu = 8;
                                        break;
                                    }
                                }

                                if (pos.up_minimum_Edu == null)
                                {
                                    return "ستون حداقل تحصیلات سطر " + " ( " + counter + " ) " + " را طبق فرمت وارد کیند ";
                                }
                            }

                            var up_General_Equipment = row.Cells[10].Value;
                            if (!string.IsNullOrEmpty(up_General_Equipment))
                            {
                                pos.up_General_Equipment = up_General_Equipment;
                            }

                            var up_Specialized_Equipment = row.Cells[11].Value;
                            if (!string.IsNullOrEmpty(up_Specialized_Equipment))
                            {
                                pos.up_Specialized_Equipment = up_Specialized_Equipment;
                            }

                            var up_work_Space = row.Cells[12].Value;
                            if (!string.IsNullOrEmpty(up_work_Space))
                            {
                                pos.up_work_Space = up_work_Space;
                            }

                            var up_Related_Jobs = row.Cells[13].Value;
                            if (!string.IsNullOrEmpty(up_Related_Jobs))
                            {
                                pos.up_Related_Jobs = up_Related_Jobs;
                            }

                            var up_Authority = row.Cells[14].Value;
                            if (!string.IsNullOrEmpty(up_Authority))
                            {
                                pos.up_Authority = up_Authority;
                            }

                            pos.up_Status = true;
                            tbBusinessSides.Add(pos);
                        }
                    }
                    else
                    {
                        return "ستون عنوان سمت سطر " + " ( " + counter + " ) " + " را وارد کیند ";
                    }
                }

                db.tbBusinessSide.AddRange(tbBusinessSides);
            }

            return Convert.ToBoolean(db.SaveChanges()).ToString();
        }
        #endregion
    }
}