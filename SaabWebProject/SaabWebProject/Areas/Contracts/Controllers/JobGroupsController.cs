using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.BaseInformation;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Data.Entity;
using System.IO;
using Telerik.Web.Spreadsheet;


namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class JobGroupsController : Controller
    {
        SaabEntities db;
        public JobGroupsController()
        {
            db = new SaabEntities();
            
        }
        public ActionResult persianCalendar()
        {
            return View();
        }
        public ActionResult testStyle()
        {
            return View();
        }

        #region تعریف متغیر ها
        tbContractJobGroupRepository rep_jobGroup = new tbContractJobGroupRepository();
        tbContractMoalefeDastMozdiRepository rep_moalefeDastMozdi = new tbContractMoalefeDastMozdiRepository();
        PersianCalendar p = new PersianCalendar();

        #endregion

        #region صفحات
        [AuthorizeAAA]
        public ActionResult ManageJobGroups()
        {
            return View("~/Areas/Contracts/Views/JobGroups/ManageJobGroups.cshtml");
        }

        /// <summary>
        /// صفحه لیست لیست گروه های شغلی
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialListJobGroups()
        {
            return View("~/Areas/Contracts/Views/JobGroups/_PartialListJobGroups.cshtml", rep_jobGroup.Update().OrderBy(p => p.jg_Year));
        }

        /// <summary>
        /// صفحه تولید یک گروه شغلی جدید
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialCreateJobGroups(int id = 0)
        {
            tbContractJobGroup obj = new tbContractJobGroup();
            if (id != 0)
            {
                obj.jg_Year = rep_jobGroup.Find(id).jg_Year;
                obj.jg_ID = id;
            }
            return View("~/Areas/Contracts/Views/JobGroups/_PartialCreateJobGroups.cshtml", obj);

        }

        /// <summary>
        /// صفحه ویرایش یک گروه شغلی 
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialUpdateJobGroups(int id = 0, int year = 0 )
        {
            tbContractJobGroup result = null;
            if (id != 0)
            {
                result = rep_jobGroup.Find(id);
            }
            if (year != 0)
            {
                result = rep_jobGroup.FindYear(year);
            }
            // ---------------------------------------

            // ---------------------------------------

            if (result == null)
            {
                result = new tbContractJobGroup();
                List<tbContractJobGroupValues> list_tbvalues = new List<tbContractJobGroupValues>();
                for (int i = 1; i <= 40; i++)
                {
                    for (int j = 0; j <= 40; j++)
                    {
                        tbContractJobGroupValues obj = new tbContractJobGroupValues();
                        obj.SanavatYear = j;
                        obj.GroupNumber = i;
                        list_tbvalues.Add(obj);
                    }
                }
                result.tbContractJobGroupValues = list_tbvalues;
            }
            return View("~/Areas/Contracts/Views/JobGroups/_PartialUpdateJobGroups.cshtml", result);
        }
        /// <summary>
        /// صفحه نمایشی اطلاعات گروه شغلی برای صفحه قرارداد ها
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialGetJobGroupMoalefeForContracts(int groupNumber, int sanavatYear)
        {
            int year = ConvertDateTimeToShamsi.ConvertDateTimeToYearShamsi(DateTime.Now);
            var result = rep_jobGroup.GetBaseInformation(year, groupNumber, sanavatYear);
            return View("~/Areas/Contracts/Views/JobGroups/_PartialGetJobGroupMoalefeForContracts.cshtml", result);
        }
        #endregion

        #region آپلود

        /// <summary>
        /// آپلود فایل اکسل
        /// </summary>
        /// <returns></returns>
        /// 
        /*------------------------------------------------[1402/08/11]-|KH|-*/
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult ImportExcel_AddJobGroups(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_JobGroups(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }
        /*-------------------------------------------------------------|KH|-*/
        #endregion

        #region توابع

        public ActionResult tBAcceptList(int Year = 0, int Month = 0, int Peyman_ID = 0)
        {
            List<tbCommodityProject> usr2 = new List<tbCommodityProject>();

            var x2 = db.tbCommodityProject
                .Select(s => s.DateOfDefinitionCommodityProject)
                .ToList();

            foreach (var item in x2)
            {
                int shamsiYear = item.GetShamsYear();
                int shamsiMonth = item.GetShamsiMonth();

                if (shamsiYear == Year && shamsiMonth == Month)
                {
                    var c = db.tbCommodityProject
                        .Where(p => p.DateOfDefinitionCommodityProject == item && p.FK_PeymanContracts == Peyman_ID && p.P_Active == true)
                        .FirstOrDefault();

                    if (c != null)
                    {
                        usr2.Add(c);
                    }
                }
            }

            return View("~/Areas/Contracts/Views/JobGroups/tBAcceptList.cshtml", usr2);

        }

        [AuthorizeAAA]
        public string SaveJobGroup(tbContractJobGroup Filters)
        {
            var isContain = rep_jobGroup.FindYear(Filters.jg_Year ?? 0);
            if (isContain == null)
            {
                return rep_jobGroup.Create(Filters);
            }
            else
            {
                return rep_jobGroup.Update(Filters);
            }

        }
        [AuthorizeAAA]
        public string UpdateJobGroup(tbContractJobGroup Filters)
        {

            var a = rep_jobGroup.Update(Filters);
            return a;
        }

        [AuthorizeAAA]
        public List<tbContractJobGroup> ListJobGroup()
        {
            return rep_jobGroup.Update(); // not usefull
                                          //  return PartialView(rep_moalefeDastmozdi.Listt());
        }
        ///////////////////////

        /*------------------------------------------------[1402/08/11]-|KH|-*/
        [AuthorizeAAA]
        public string GetDataFromExcel_JobGroups(HttpPostedFileBase MyExcelStream)
        {

            int lastId=0;
            string mozdGroup = "";
            string sanavat = "";
            List<tbContractJobGroup> listJobGroup = new List<tbContractJobGroup>();
            List<tbContractJobGroupValues> listJobGroupValues = new List<tbContractJobGroupValues>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets > 1)
                    {
                        var data = workbook.Sheets[0];
                        var rows = data.Rows;


                        if (rows.Count >= 1)
                        {

                            //var count = workbook.Sheets[0].Rows.Count();

                            string[] check = { "سال", "شرح", "حق مسکن", "حق اولاد", "خواروبار" };
                            string[] value = { "", "", "", "", "", };



                            for (int i = 0; i < rows.Count; i++)
                            {
                                //----------------------------------------------
                                if (i < 5)
                                {
                                    if (data.Rows[i].Cells[1].Value != null)
                                    {
                                        value[i] = data.Rows[i].Cells[1].Value.ToString();
                                    }
                                    else
                                    {

                                        return $" لطفا بخش '{check[i]}' را پر کنید";
                                    }
                                }
                                if (i == 5)
                                {
                                    //1. jg_HagheMaskan = value[2];
                                    //2. jg_HagheOlad = value[3];
                                    //3. jg_KharoBar = value[4];
                                    //4. jg_Year = value[0];
                                    //5. isActive= 0;
                                    //6. title = value[1];
                                    tbContractJobGroup JobGroup = new tbContractJobGroup
                                    {
                                        jg_HagheMaskan = value[2],
                                        jg_HagheOlad = value[3],
                                        jg_KharoBar = value[4],
                                        jg_Year = Convert.ToInt32(value[0]),
                                        isActive = true,
                                        title = value[1],
                                    };
                                    listJobGroup.Add(JobGroup);
                                    rep_jobGroup.Create(JobGroup);

                                    var id = db.tbContractJobGroup
                                        .OrderByDescending(p => p.jg_ID)
                                        .FirstOrDefault();
                                    lastId = id.jg_ID;
                                }
                                //----------------------------------------------
                                if (i > 5)
                                {
                                    

                                    for (int j = 1; j < data.Rows[5].Cells.Count; j++)
                                    {

                                        //1. FK_JobGroupID = id.jg_ID
                                        //2. GroupNumber = i-6 // Cont++ 
                                        //3. SanavatYear = j-1

                                        Console.WriteLine(j - 1);
                                        if (j == 1)
                                        {
                                            //4. ValueMozdGroup = data.Rows[i].Cells[1].Value
                                            //5. ValueSanavat = null
                                            if (data.Rows[i].Cells[1].Value != null)
                                            {
                                                mozdGroup = data.Rows[i].Cells[1].Value.ToString();
                                            }
                                            else
                                            {
                                                mozdGroup = null;
                                            }
                                            sanavat = null;
                                        }
                                        else
                                        {
                                            //4. ValueMozdGroup = null
                                            //5. ValueSanavat = data.Rows[i].Cells[j].Value
                                            mozdGroup = null;
                                            if (data.Rows[i].Cells[j].Value != null)
                                            {
                                                sanavat = data.Rows[i].Cells[j].Value.ToString();
                                            }
                                            else
                                            {
                                                sanavat = null;
                                            }
                                        }
                                        tbContractJobGroupValues JobGroupValue = new tbContractJobGroupValues
                                        {
                                            FK_JobGroupID = lastId,
                                            GroupNumber = i - 5,
                                            SanavatYear = j - 1,
                                            ValueMozdGroup = mozdGroup,
                                            ValueSanavat = sanavat,
                                        };
                                        listJobGroupValues.Add(JobGroupValue);
                                        rep_jobGroup.CreateValue(JobGroupValue);

                                    }
                                   
                                }
                                //----------------------------------------------
                            }

                            ////////////////////////
                            ///////////////////////
                            ///////////////////////

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

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;
                }
            }
            return ""; /////-----------------------------------
        }
        /*-------------------------------------------------------------|KH|-*/

        #endregion
    }
}