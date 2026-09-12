using Microsoft.CodeAnalysis.CSharp.Syntax;
using SaabWebProject.Areas.Contracts.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.Repositories.BaseInformation;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Web.Spreadsheet;
using static SaabWebProject.Areas.Users.Controllers.MessageBoxController;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class MoalefeDastMozdiController : Controller
    {
        SaabEntities db;
        tbContractMoalefeDastMozdiRepository rep_moalefeDastmozdi;
        tbCategoriesRepository categoryRepo;
        tbContractPanelsRepository PanelsRepo;
        tbPeymanContractsRepository PaymanRepo;
        tbControllyPanelRepository ControllyPanelRepo;
        tbControllyPeymanRepository ControllyPeymanRepo;
        tbUnitParameterRepository UnitRepo;

        #region سازنده ها

        public MoalefeDastMozdiController()
        {
            db = new SaabEntities();
            rep_moalefeDastmozdi = new tbContractMoalefeDastMozdiRepository(db);
            categoryRepo = new tbCategoriesRepository(db);
            PanelsRepo = new tbContractPanelsRepository(db);
            PaymanRepo = new tbPeymanContractsRepository(db);
            ControllyPanelRepo = new tbControllyPanelRepository(db);
            ControllyPeymanRepo = new tbControllyPeymanRepository(db);
            UnitRepo = new tbUnitParameterRepository(db);
        }

        #endregion

        #region صفحات

        /// <summary>
        /// صفحه تعریف مولفه های دست مزدی
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialCreateMoalefeDastMozdi()
        {
            ViewBag.TitleLargModal = "تعریف مولفه دستمزدی جدید";
            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PartialCreateMoalefeDastMozdi.cshtml");
        }

        /// <summary>
        /// صفحه لیست مولفه های دست مزدی
        /// </summary>
        /// <returns></returns>
        [AuthorizeAAA]
        public ActionResult ManageMoalefeDastMozdi()
        {
            return View("~/Areas/Contracts/Views/MoalefeDastMozdi/ManageMoalefeDastMozdi.cshtml");
        }

        /// <summary>
        /// صفحه لیست مولفه های دستمزدی
        /// </summary>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialListMoalefeDastMozdi()
        {
            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PartialListMoalefeDastMozdi.cshtml", rep_moalefeDastmozdi.Update().OrderByDescending(p => p.md_ID).ToList());
        }

        /// <summary>
        /// صفحه ویرایش مولفه دستمزدی
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public ActionResult _PartialUpdateMoalefeDastMozdi(int id = 0)
        {
            ViewBag.TitleLargModal = "ویرایش مولفه";

            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PartialUpdateMoalefeDastMozdi.cshtml", rep_moalefeDastmozdi.Find(id));
        }


        [AuthorizeAAA]
        public ActionResult _Exists_MoalefeDastmozdi()
        {
            var list = GetAllMoalefeForFormulas();
            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_Exists_MoalefeDastmozdi.cshtml", list);
        }


        [AuthorizeAAA]
        public ActionResult _PartialViewGetGhararDadiMoalefeForGharardadProject()
        {
            var list = rep_moalefeDastmozdi.Update().ToList();
            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PartialViewGetGhararDadiMoalefeForGharardadProject.cshtml", list);
        }

        [AuthorizeAAA]
        public ActionResult TableSynced_Panel_Moalefe()
        {
            var st = ControllyPanelRepo.Listt();
            return PartialView("_TableSynced_Panel_Moalefe", st);
        }

        [AuthorizeAAA]
        public ActionResult TableSynced_Peyman_Moalefe()
        {
            var st = ControllyPeymanRepo.Listt();
            return PartialView("_TableSynced_Peyman_Moalefe", st);
        }

        [AuthorizeAAA]
        public ActionResult EditSynced_Peyman_Moalefe(int id)
        {
            var st = ControllyPeymanRepo.Find(id);
            var st1 = ControllyPeymanRepo.Listt().Where(c => c.FK_PeymanId == st.FK_PeymanId);

            return PartialView("_EditSynced_Peyman_Moalefe", st1);
        }

        [AuthorizeAAA]
        public ActionResult EditSynced_Panel_Moalefe(int id)
        {
            var st = ControllyPanelRepo.Find(id);
            var st1 = ControllyPanelRepo.Listt().Where(c => c.FK_PanelsContractId == st.FK_PanelsContractId);

            return PartialView("_EditSynced_Panel_Moalefe", st1);
        }



        #endregion

        #region توابع

        /// <summary>
        /// تعریف مولفه های دست مزدی
        /// </summary>
        /// <param name="obj_moalefe"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public string SaveMoalefeDastMozdi(tbContractMoalefeDastmozdi Filters)
        {
            Filters.md_IsActive = true;
            if (Filters.FK_Category_ID == -1)
            {
                Filters.FK_Category_ID = null;
            }

            if (Filters.md_variable != null && Filters.md_Type != -1 && Filters.md_Title.Trim().Length != 0)
                return rep_moalefeDastmozdi.Create(Filters).ToString();
            else
                return "false_notValid";
        }

        /// <summary>
        /// ویرایش مولفه دستمزدی
        /// </summary>
        /// <param name="Filters"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public string UpdateMoalefeDastMozdi(tbContractMoalefeDastmozdi Filters)
        {
            if (Filters.FK_Category_ID == -1)
            {
                Filters.FK_Category_ID = null;
            }

            if (Filters.md_variable != null && Filters.md_Type != -1 && Filters.md_Title.Trim().Length != 0)
                return rep_moalefeDastmozdi.Update(Filters).ToString();
            else
                return "false_notValid";
        }

        /// <summary>
        /// فعال یا غیر فعال سازی مولفه دستمزدی
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        /// 
        [AuthorizeAAA]
        public bool DisableEnableMoalefeDastMozdi(int ID = 0)
        {
            return rep_moalefeDastmozdi.Disable(ID);
        }

        [AuthorizeAAA]
        public ActionResult GetAllMoalefeForFormulasanyform()
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "شماره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_ContractNumber" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "عنوان شغل ( قراردادی)", Type = 0, GharardadColoumnName = "usc_Jobtitle" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تعداد سال سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TedadSalSanavat" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "مزد گروه ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueMozdGroup" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "مزد سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueSanavat" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "خوار و بار ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_KharoBar" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "حق مسکن ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheMaskan" });
            myList.Add(new ListOfMoalefeha { variablePersianName = "حق اولاد ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheOlad" });
            foreach (var item in rep_moalefeDastmozdi.Update())
            {
                string type = "قراردادی";
                switch (item.md_Type)
                {
                    case (int)Type_Moalefe.Gharardadi:
                        type = "قراردادی";
                        break;
                    case (int)Type_Moalefe.Amalkardi:
                        type = "عملکردی";
                        break;
                    case (int)Type_Moalefe.Karbari:
                        type = "کاربری";
                        break;
                    case (int)Type_Moalefe.Sayer:
                        type = "سایر";
                        break;
                    case (int)Type_Moalefe.Controlli:
                        type = "کنترلی";
                        break;
                    case (int)Type_Moalefe.Dastmozdi:
                        type = "دستمزدی";
                        break;
                    case (int)Type_Moalefe.Calculational:
                        type = "محاسباتی";
                        break;
                    case (int)Type_Moalefe.Karkardi:
                        type = "کارکردی";
                        break;
                    case (int)Type_Moalefe.Pishkhan:
                        type = "پیشخوان";
                        break;
                    case (int)Type_Moalefe.Fish:
                        type = "فیش‏حقوقی";

                        break;
                    case (int)Type_Moalefe.sorat:
                        type = "صورت وضعیت";
                        break;
                }

                if (item.tbFormula.Any())
                {
                    myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = true });
                }
                else
                {
                    myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = false });
                }
            }

            return PartialView("_Exists_MoalefeDastmozdi", myList);
        }
        //testmoalfeh
        public ActionResult GetAllMoalefeForFormulasbyid(int id=0)
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            if (id == 0)
            {
                myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "شماره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_ContractNumber", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "عنوان شغل ( قراردادی)", Type = 0, GharardadColoumnName = "usc_Jobtitle", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تعداد سال سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TedadSalSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد گروه ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueMozdGroup", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "خوار و بار ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_KharoBar", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق مسکن ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheMaskan", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق اولاد ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheOlad", noghrogi = 1 });
                foreach (var item in rep_moalefeDastmozdi.Update())
                {
                    string type = "قراردادی";
                    switch (item.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }

                    if (item.tbFormula.Any())
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = true, noghrogi = 1 });
                    }
                    else
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = false, noghrogi = 1 });
                    }
                }
            }
            //حقوقی
           else if (id == 1)
            {
                #region             کاربران حقوقی
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام مدیر عامل  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اختصاصی شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثبت (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اقتصادی (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل ثبت شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });
                #endregion
            }
            //پیمان ها 
            else if (id == 2)
            {
                #region             //پیمان ها 
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پروژه (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مخفف پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما اصلی (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ شروع پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ پایان پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مبلغ کل قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });


                foreach (var it in db.tbPeymanContractPrice.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -600,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( شماره قرارداد پیمان  )",
                        Type = -601,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( مبلغ کل قرارداد )",
                        Type = -602,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه شروع قرارداد)",
                        Type = -603,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال شروع قرارداد)",
                        Type = -604,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه پایان قرارداد)",
                        Type = -605,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال پایان قرارداد)",
                        Type = -606,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (کد پروژه)",
                        Type = -607,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -6,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                        Type = -7,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                        Type = -8,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                        Type = -9,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                }



             



                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    int shomareh = -330;
                //    for(int i = 0; i <= 30; i++)
                //    {
                //        shomareh += i;
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = it.ID,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName =  "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = -200,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //    }


                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}
                //int shomareh23 = -330;

                //for (int i = 0; i <= 30; i++)
                //{
                //    shomareh23 += i;
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //    //    Type = shomareh23,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 10
                //    //});
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //        Type = shomareh23,
                //        ID = -200,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });
                //}





                //foreach (var it in db.tbPeymanContracts.ToList())
                //{
                //    foreach (var item in rep_moalefeDastmozdi.Update())
                //    {
                //        string type = "قراردادی";
                //        switch (item.md_Type)
                //        {
                //            case (int)Type_Moalefe.Gharardadi:
                //                type = "قراردادی";
                //                break;
                //            case (int)Type_Moalefe.Amalkardi:
                //                type = "عملکردی";
                //                break;
                //            case (int)Type_Moalefe.Karbari:
                //                type = "کاربری";
                //                break;
                //            case (int)Type_Moalefe.Sayer:
                //                type = "سایر";
                //                break;
                //            case (int)Type_Moalefe.Controlli:
                //                type = "کنترلی";
                //                break;
                //            case (int)Type_Moalefe.Dastmozdi:
                //                type = "دستمزدی";
                //                break;
                //            case (int)Type_Moalefe.Calculational:
                //                type = "محاسباتی";
                //                break;
                //            case (int)Type_Moalefe.Karkardi:
                //                type = "کارکردی";
                //                break;
                //            case (int)Type_Moalefe.Pishkhan:
                //                type = "پیشخوان";
                //                break;
                //            case (int)Type_Moalefe.Fish:
                //                type = "فیش‏حقوقی";

                //                break;
                //            case (int)Type_Moalefe.sorat:
                //                type = "صورت وضعیت";
                //                break;
                //        }
                //        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )" + "(" + it.pec_Title + ")", ID = item.md_ID, Type = -13, ExistsFormula = false, noghrogi = 1 });

                //    }
                //}


                #endregion
            }
            //شغل
            else if (id == 3)
            {
                #region شغل ها 
                foreach (var it in db.tbdetailjob.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.detail_title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شغل  )",
                        Type = -14,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //شهرستان
            else if (id == 4)
            {
                #region شهرستان ها 
                foreach (var it in db.tbCities.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شهرستان ها   )",
                        Type = -15,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب شهرستان ها   )",
                        Type = -16,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب مالیات شهرستان   )",
                        Type = -17,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //قرارداد پرسنل
            else if (id == 5)
            {
                #region قرارداد های پرسنل
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام (قرارداد پرسنل)",
                    Type = -2,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام ونام خانوادگی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پرسنلی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد ملی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پدر (قرارداد پرسنل)",
                    Type = -2,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل صدور شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "جنسیت (قرارداد پرسنل)",
                    Type = -2,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "وضعیت تاهل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره موبایل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -12,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثابت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -13,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد  مشمول بیمه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -14,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -15,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب مالیات شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب شهرستان شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -19,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "پیمان (قرارداد پرسنل)",
                    Type = -2,
                    ID = -16,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شغل (قرارداد پرسنل)",
                    Type = -2,
                    ID = -17,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });










                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق اولاد (قرارداد های پرسنل )",
                    Type = -5,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "خواربار (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد شغل (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد سایر (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد سال سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق مسکن (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه  شروع قرارداد  (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال شروع قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                }); myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                foreach (var it in db.tbUserContractsAndMoalefeGhararDadi.GroupBy(s => s.FKMoalefeGhararDadi).ToList())
                {
                    var findca = db.tbContractMoalefeDastmozdi.Where(S => S.md_ID == it.Key).FirstOrDefault();
                    if (findca != null)
                    {
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = findca.md_Title + " (قرارداد های پرسنل )",
                            Type = -5,
                            ID = findca.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 5
                        });
                    }

                }

                #endregion
            }
            //تجهیزات
            else if (id == 6)
            {
                
                foreach (var it in db.tbEquipments.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " ( نام تجهیز )",
                        Type = -11,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " (  رقم تجهیز )",
                        Type = -12,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }


            }
            //بودجه
            else if (id == 7)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( بودجه )",
                        Type = -10,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    for (int i = 1; i <= 20; i++)
                    {
                        int sum = it.Pyman_ID;
                        sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                        // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                        string base9String = ConvertToBase9(sum);

                        // تبدیل string مبنای 9 به عدد int
                        long base9Int = long.Parse(base9String);

                        // اضافه کردن عدد 9 به انتهای عدد
                        string newBase9String = base9Int.ToString() + "9";
                        newBase9String = newBase9String.ToString() + i;
                        string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                        int base9Int2 = int.Parse(base9String2);
                        newBase9String = newBase9String.ToString() + "9";
                        newBase9String = newBase9String.ToString() + base9Int2;

                        // تبدیل دوباره به عدد int
                        xx = long.Parse(newBase9String);
                        long xxID = long.Parse(newBase9String);
                        if (base9Int == 1431)
                        {

                        }
                        shomareh += i;
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                            Type = -10000,
                            ID = xxID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                            Type = shomareh,
                            ID = -200,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                foreach (var it in db.tbCreditIndicators.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                        Type = -10,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }

            }
            //تزاز فیش
            else if (id == 8)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( سبد سایر عملیات )",
                        Type = -15,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    //for (int i = 1; i <= 20; i++)
                    //{
                    //    int sum = it.Pyman_ID;
                    //    sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                    //    // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                    //    string base9String = ConvertToBase9(sum);

                    //    // تبدیل string مبنای 9 به عدد int
                    //    long base9Int = long.Parse(base9String);

                    //    // اضافه کردن عدد 9 به انتهای عدد
                    //    string newBase9String = base9Int.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + i;
                    //    string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                    //    int base9Int2 = int.Parse(base9String2);
                    //    newBase9String = newBase9String.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + base9Int2;

                    //    // تبدیل دوباره به عدد int
                    //    xx = long.Parse(newBase9String);
                    //    long xxID = long.Parse(newBase9String);
                    //    if (base9Int == 1431)
                    //    {

                    //    }
                    //    shomareh += i;
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                    //        Type = -10000,
                    //        ID = xxID,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                    //        Type = shomareh,
                    //        ID = -200,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //}

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                //        Type = -10,
                //        ID = it.ID,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}

            }
            
            else if (id == 9)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کسورات",
                    Type = -13,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "اضافات",
                    Type = -13,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
            }

            else if (id == 10)
            {
                var find = db.tbFormula.Where(s => s.Frml_IsActive == true).ToList();
                foreach(var it in find)
                {
                    string type = "قراردادی";
                    switch (it.tbContractMoalefeDastmozdi.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName =  "(" + it.tbContractMoalefeDastmozdi.md_Title + ")" + " ( فرمول ها )",
                        Type = it.tbContractMoalefeDastmozdi.md_Type,
                        ID = (long)it.FKMoalefeDastMozdi,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                }
             
              
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }
            else if (id == 11)
            {
                var find = db.tbinstallments.GroupBy(s => s.Fk_molfe).ToList();
                foreach (var it1 in find)
                {
                    var find2 = db.tbContractMoalefeDastmozdi.Where(s => s.md_ID == it1.Key).ToList();
                    foreach (var it in find2)
                    {
                        string type = "قراردادی";
                        //switch (it.tbContractMoalefeDastmozdi.md_Type)
                        //{
                        //    case (int)Type_Moalefe.Gharardadi:
                        //        type = "قراردادی";
                        //        break;
                        //    case (int)Type_Moalefe.Amalkardi:
                        //        type = "عملکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Karbari:
                        //        type = "کاربری";
                        //        break;
                        //    case (int)Type_Moalefe.Sayer:
                        //        type = "سایر";
                        //        break;
                        //    case (int)Type_Moalefe.Controlli:
                        //        type = "کنترلی";
                        //        break;
                        //    case (int)Type_Moalefe.Dastmozdi:
                        //        type = "دستمزدی";
                        //        break;
                        //    case (int)Type_Moalefe.Calculational:
                        //        type = "محاسباتی";
                        //        break;
                        //    case (int)Type_Moalefe.Karkardi:
                        //        type = "کارکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Pishkhan:
                        //        type = "پیشخوان";
                        //        break;
                        //    case (int)Type_Moalefe.Fish:
                        //        type = "فیش‏حقوقی";

                        //        break;
                        //    case (int)Type_Moalefe.sorat:
                        //        type = "صورت وضعیت";
                        //        break;
                        //}
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(" + it.md_Title + ")" + " ( فرمول ها )",
                            Type =-17,
                            ID = (long)it.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }

                
                }


                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }

            else if (id == 12)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "معادل کار",
                    Type = -18,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
             
            }

            //کاربران حقوقی
            string قراردادپرسنل = "قرارداد A"; // فقط نمونه است، مقدار واقعی رو باید مشخص کنید








            return PartialView("_Exists_MoalefeDastmozdi", myList);
        }
        public ActionResult GetAllMoalefeForFormulasbyid2(int id = 0)
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            if (id == 0)
            {
                myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "شماره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_ContractNumber", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "عنوان شغل ( قراردادی)", Type = 0, GharardadColoumnName = "usc_Jobtitle", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تعداد سال سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TedadSalSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد گروه ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueMozdGroup", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "خوار و بار ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_KharoBar", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق مسکن ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheMaskan", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق اولاد ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheOlad", noghrogi = 1 });
                foreach (var item in rep_moalefeDastmozdi.Update())
                {
                    string type = "قراردادی";
                    switch (item.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }

                    if (item.tbFormula.Any())
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = true, noghrogi = 1 });
                    }
                    else
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = false, noghrogi = 1 });
                    }
                }
            }
            //حقوقی
            else if (id == 1)
            {
                #region             کاربران حقوقی
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام مدیر عامل  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اختصاصی شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثبت (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اقتصادی (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل ثبت شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });
                #endregion
            }
            //پیمان ها 
            else if (id == 2)
            {
                #region             //پیمان ها 
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پروژه (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مخفف پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما اصلی (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ شروع پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ پایان پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مبلغ کل قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });


                foreach (var it in db.tbPeymanContractPrice.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -600,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( شماره قرارداد پیمان  )",
                        Type = -601,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( مبلغ کل قرارداد )",
                        Type = -602,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه شروع قرارداد)",
                        Type = -603,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال شروع قرارداد)",
                        Type = -604,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه پایان قرارداد)",
                        Type = -605,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال پایان قرارداد)",
                        Type = -606,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (کد پروژه)",
                        Type = -607,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -6,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                        Type = -7,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                        Type = -8,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                        Type = -9,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                }







                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    int shomareh = -330;
                //    for(int i = 0; i <= 30; i++)
                //    {
                //        shomareh += i;
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = it.ID,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName =  "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = -200,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //    }


                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}
                //int shomareh23 = -330;

                //for (int i = 0; i <= 30; i++)
                //{
                //    shomareh23 += i;
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //    //    Type = shomareh23,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 10
                //    //});
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //        Type = shomareh23,
                //        ID = -200,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });
                //}





                //foreach (var it in db.tbPeymanContracts.ToList())
                //{
                //    foreach (var item in rep_moalefeDastmozdi.Update())
                //    {
                //        string type = "قراردادی";
                //        switch (item.md_Type)
                //        {
                //            case (int)Type_Moalefe.Gharardadi:
                //                type = "قراردادی";
                //                break;
                //            case (int)Type_Moalefe.Amalkardi:
                //                type = "عملکردی";
                //                break;
                //            case (int)Type_Moalefe.Karbari:
                //                type = "کاربری";
                //                break;
                //            case (int)Type_Moalefe.Sayer:
                //                type = "سایر";
                //                break;
                //            case (int)Type_Moalefe.Controlli:
                //                type = "کنترلی";
                //                break;
                //            case (int)Type_Moalefe.Dastmozdi:
                //                type = "دستمزدی";
                //                break;
                //            case (int)Type_Moalefe.Calculational:
                //                type = "محاسباتی";
                //                break;
                //            case (int)Type_Moalefe.Karkardi:
                //                type = "کارکردی";
                //                break;
                //            case (int)Type_Moalefe.Pishkhan:
                //                type = "پیشخوان";
                //                break;
                //            case (int)Type_Moalefe.Fish:
                //                type = "فیش‏حقوقی";

                //                break;
                //            case (int)Type_Moalefe.sorat:
                //                type = "صورت وضعیت";
                //                break;
                //        }
                //        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )" + "(" + it.pec_Title + ")", ID = item.md_ID, Type = -13, ExistsFormula = false, noghrogi = 1 });

                //    }
                //}


                #endregion
            }
            //شغل
            else if (id == 3)
            {
                #region شغل ها 
                foreach (var it in db.tbdetailjob.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.detail_title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شغل  )",
                        Type = -14,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //شهرستان
            else if (id == 4)
            {
                #region شهرستان ها 
                foreach (var it in db.tbCities.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شهرستان ها   )",
                        Type = -15,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب شهرستان ها   )",
                        Type = -16,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب مالیات شهرستان   )",
                        Type = -17,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //قرارداد پرسنل
            else if (id == 5)
            {
                #region قرارداد های پرسنل
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام (قرارداد پرسنل)",
                    Type = -2,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام ونام خانوادگی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پرسنلی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد ملی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پدر (قرارداد پرسنل)",
                    Type = -2,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل صدور شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "جنسیت (قرارداد پرسنل)",
                    Type = -2,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "وضعیت تاهل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره موبایل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -12,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثابت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -13,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد  مشمول بیمه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -14,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -15,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب مالیات شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب شهرستان شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -19,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "پیمان (قرارداد پرسنل)",
                    Type = -2,
                    ID = -16,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شغل (قرارداد پرسنل)",
                    Type = -2,
                    ID = -17,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });










                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق اولاد (قرارداد های پرسنل )",
                    Type = -5,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "خواربار (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد شغل (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد سایر (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد سال سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق مسکن (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه  شروع قرارداد  (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال شروع قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                }); myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                foreach (var it in db.tbUserContractsAndMoalefeGhararDadi.GroupBy(s => s.FKMoalefeGhararDadi).ToList())
                {
                    var findca = db.tbContractMoalefeDastmozdi.Where(S => S.md_ID == it.Key).FirstOrDefault();
                    if (findca != null)
                    {
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = findca.md_Title + " (قرارداد های پرسنل )",
                            Type = -5,
                            ID = findca.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 5
                        });
                    }

                }

                #endregion
            }
            //تجهیزات
            else if (id == 6)
            {

                foreach (var it in db.tbEquipments.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " ( نام تجهیز )",
                        Type = -11,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " (  رقم تجهیز )",
                        Type = -12,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }


            }
            //بودجه
            else if (id == 7)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( بودجه )",
                        Type = -10,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    for (int i = 1; i <= 20; i++)
                    {
                        int sum = it.Pyman_ID;
                        sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                        // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                        string base9String = ConvertToBase9(sum);

                        // تبدیل string مبنای 9 به عدد int
                        long base9Int = long.Parse(base9String);

                        // اضافه کردن عدد 9 به انتهای عدد
                        string newBase9String = base9Int.ToString() + "9";
                        newBase9String = newBase9String.ToString() + i;
                        string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                        int base9Int2 = int.Parse(base9String2);
                        newBase9String = newBase9String.ToString() + "9";
                        newBase9String = newBase9String.ToString() + base9Int2;

                        // تبدیل دوباره به عدد int
                        xx = long.Parse(newBase9String);
                        long xxID = long.Parse(newBase9String);
                        if (base9Int == 1431)
                        {

                        }
                        shomareh += i;
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                            Type = -10000,
                            ID = xxID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                            Type = shomareh,
                            ID = -200,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                foreach (var it in db.tbCreditIndicators.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                        Type = -10,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }

            }
            //تزاز فیش
            else if (id == 8)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( سبد سایر عملیات )",
                        Type = -15,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    //for (int i = 1; i <= 20; i++)
                    //{
                    //    int sum = it.Pyman_ID;
                    //    sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                    //    // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                    //    string base9String = ConvertToBase9(sum);

                    //    // تبدیل string مبنای 9 به عدد int
                    //    long base9Int = long.Parse(base9String);

                    //    // اضافه کردن عدد 9 به انتهای عدد
                    //    string newBase9String = base9Int.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + i;
                    //    string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                    //    int base9Int2 = int.Parse(base9String2);
                    //    newBase9String = newBase9String.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + base9Int2;

                    //    // تبدیل دوباره به عدد int
                    //    xx = long.Parse(newBase9String);
                    //    long xxID = long.Parse(newBase9String);
                    //    if (base9Int == 1431)
                    //    {

                    //    }
                    //    shomareh += i;
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                    //        Type = -10000,
                    //        ID = xxID,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                    //        Type = shomareh,
                    //        ID = -200,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //}

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                //        Type = -10,
                //        ID = it.ID,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}

            }

            else if (id == 9)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کسورات",
                    Type = -13,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "اضافات",
                    Type = -13,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
            }

            else if (id == 10)
            {
                var find = db.tbFormula.Where(s => s.Frml_IsActive == true).ToList();
                foreach (var it in find)
                {
                    string type = "قراردادی";
                    switch (it.tbContractMoalefeDastmozdi.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbContractMoalefeDastmozdi.md_Title + ")" + " ( فرمول ها )",
                        Type = it.tbContractMoalefeDastmozdi.md_Type,
                        ID = (long)it.FKMoalefeDastMozdi,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                }


                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }
            else if (id == 11)
            {
                var find = db.tbinstallments.GroupBy(s => s.Fk_molfe).ToList();
                foreach (var it1 in find)
                {
                    var find2 = db.tbContractMoalefeDastmozdi.Where(s => s.md_ID == it1.Key).ToList();
                    foreach (var it in find2)
                    {
                        string type = "قراردادی";
                        //switch (it.tbContractMoalefeDastmozdi.md_Type)
                        //{
                        //    case (int)Type_Moalefe.Gharardadi:
                        //        type = "قراردادی";
                        //        break;
                        //    case (int)Type_Moalefe.Amalkardi:
                        //        type = "عملکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Karbari:
                        //        type = "کاربری";
                        //        break;
                        //    case (int)Type_Moalefe.Sayer:
                        //        type = "سایر";
                        //        break;
                        //    case (int)Type_Moalefe.Controlli:
                        //        type = "کنترلی";
                        //        break;
                        //    case (int)Type_Moalefe.Dastmozdi:
                        //        type = "دستمزدی";
                        //        break;
                        //    case (int)Type_Moalefe.Calculational:
                        //        type = "محاسباتی";
                        //        break;
                        //    case (int)Type_Moalefe.Karkardi:
                        //        type = "کارکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Pishkhan:
                        //        type = "پیشخوان";
                        //        break;
                        //    case (int)Type_Moalefe.Fish:
                        //        type = "فیش‏حقوقی";

                        //        break;
                        //    case (int)Type_Moalefe.sorat:
                        //        type = "صورت وضعیت";
                        //        break;
                        //}
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(" + it.md_Title + ")" + " ( فرمول ها )",
                            Type = -17,
                            ID = (long)it.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }


                }


                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }

            else if (id == 12)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "معادل کار",
                    Type = -18,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

            }

            //کاربران حقوقی
            string قراردادپرسنل = "قرارداد A"; // فقط نمونه است، مقدار واقعی رو باید مشخص کنید




            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/GetAllMoalefeForFormulasbyid2.cshtml", myList);



            //return Json(myList, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GetAllMoalefeForFormulasbyid3(int id = 0)
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            if (id == 0)
            {
                myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "شماره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_ContractNumber", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "عنوان شغل ( قراردادی)", Type = 0, GharardadColoumnName = "usc_Jobtitle", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "تعداد سال سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TedadSalSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد گروه ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueMozdGroup", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "مزد سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueSanavat", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "خوار و بار ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_KharoBar", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق مسکن ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheMaskan", noghrogi = 1 });
                myList.Add(new ListOfMoalefeha { variablePersianName = "حق اولاد ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheOlad", noghrogi = 1 });
                foreach (var item in rep_moalefeDastmozdi.Update())
                {
                    string type = "قراردادی";
                    switch (item.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }

                    if (item.tbFormula.Any())
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = true, noghrogi = 1 });
                    }
                    else
                    {
                        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = false, noghrogi = 1 });
                    }
                }
            }
            //حقوقی
            else if (id == 1)
            {
                #region             کاربران حقوقی
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام مدیر عامل  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اختصاصی شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثبت (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد اقتصادی (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل ثبت شرکت  (قرارداد کاربران حقوقی)",
                    Type = -3,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 3
                });
                #endregion
            }
            //پیمان ها 
            else if (id == 2)
            {
                #region             //پیمان ها 
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پروژه (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مخفف پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کارفرما اصلی (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ شروع پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ پایان پیمان (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مبلغ کل قرارداد (قرارداد پیمان ها)",
                    Type = -4,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });


                foreach (var it in db.tbPeymanContractPrice.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -600,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( شماره قرارداد پیمان  )",
                        Type = -601,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( مبلغ کل قرارداد )",
                        Type = -602,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه شروع قرارداد)",
                        Type = -603,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال شروع قرارداد)",
                        Type = -604,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه پایان قرارداد)",
                        Type = -605,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    }); myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال پایان قرارداد)",
                        Type = -606,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (کد پروژه)",
                        Type = -607,
                        ID = (int)it.FKPeymanID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")",
                        Type = -6,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                        Type = -7,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                        Type = -8,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                        Type = -9,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 6
                    });

                }







                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    int shomareh = -330;
                //    for(int i = 0; i <= 30; i++)
                //    {
                //        shomareh += i;
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = it.ID,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //        myList.Add(new ListOfMoalefeha
                //        {
                //            variablePersianName =  "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //            Type = shomareh,
                //            ID = -200,
                //            ExistsFormula = false,
                //            GharardadColoumnName = "usc_CountOfChild",
                //            noghrogi = 10
                //        });
                //    }


                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}
                //int shomareh23 = -330;

                //for (int i = 0; i <= 30; i++)
                //{
                //    shomareh23 += i;
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
                //    //    Type = shomareh23,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 10
                //    //});
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                //        Type = shomareh23,
                //        ID = -200,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });
                //}





                //foreach (var it in db.tbPeymanContracts.ToList())
                //{
                //    foreach (var item in rep_moalefeDastmozdi.Update())
                //    {
                //        string type = "قراردادی";
                //        switch (item.md_Type)
                //        {
                //            case (int)Type_Moalefe.Gharardadi:
                //                type = "قراردادی";
                //                break;
                //            case (int)Type_Moalefe.Amalkardi:
                //                type = "عملکردی";
                //                break;
                //            case (int)Type_Moalefe.Karbari:
                //                type = "کاربری";
                //                break;
                //            case (int)Type_Moalefe.Sayer:
                //                type = "سایر";
                //                break;
                //            case (int)Type_Moalefe.Controlli:
                //                type = "کنترلی";
                //                break;
                //            case (int)Type_Moalefe.Dastmozdi:
                //                type = "دستمزدی";
                //                break;
                //            case (int)Type_Moalefe.Calculational:
                //                type = "محاسباتی";
                //                break;
                //            case (int)Type_Moalefe.Karkardi:
                //                type = "کارکردی";
                //                break;
                //            case (int)Type_Moalefe.Pishkhan:
                //                type = "پیشخوان";
                //                break;
                //            case (int)Type_Moalefe.Fish:
                //                type = "فیش‏حقوقی";

                //                break;
                //            case (int)Type_Moalefe.sorat:
                //                type = "صورت وضعیت";
                //                break;
                //        }
                //        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )" + "(" + it.pec_Title + ")", ID = item.md_ID, Type = -13, ExistsFormula = false, noghrogi = 1 });

                //    }
                //}


                #endregion
            }
            //شغل
            else if (id == 3)
            {
                #region شغل ها 
                foreach (var it in db.tbdetailjob.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.detail_title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شغل  )",
                        Type = -14,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //شهرستان
            else if (id == 4)
            {
                #region شهرستان ها 
                foreach (var it in db.tbCities.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شهرستان ها   )",
                        Type = -15,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب شهرستان ها   )",
                        Type = -16,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب مالیات شهرستان   )",
                        Type = -17,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 14
                    });

                }
                #endregion
            }
            //قرارداد پرسنل
            else if (id == 5)
            {
                #region قرارداد های پرسنل
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام (قرارداد پرسنل)",
                    Type = -2,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام ونام خانوادگی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد پرسنلی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کد ملی (قرارداد پرسنل)",
                    Type = -2,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تاریخ تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "نام پدر (قرارداد پرسنل)",
                    Type = -2,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل تولد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "محل صدور شناسنامه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "جنسیت (قرارداد پرسنل)",
                    Type = -2,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "وضعیت تاهل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره موبایل  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -12,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شماره ثابت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -13,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد اولاد  مشمول بیمه (قرارداد پرسنل)",
                    Type = -2,
                    ID = -14,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -15,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب مالیات شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -18,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ضزیب شهرستان شهر محل خدمت  (قرارداد پرسنل)",
                    Type = -2,
                    ID = -19,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "پیمان (قرارداد پرسنل)",
                    Type = -2,
                    ID = -16,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "شغل (قرارداد پرسنل)",
                    Type = -2,
                    ID = -17,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 2
                });










                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق اولاد (قرارداد های پرسنل )",
                    Type = -5,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "خواربار (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -3,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد شغل (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -4,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "مزد سایر (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -10,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "تعداد سال سنوات (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -11,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "حق مسکن (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -5,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });


                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه  شروع قرارداد  (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -6,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال شروع قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -7,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                }); myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "ماه پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -8,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "سال پایان قرارداد (قرارداد های پرسنل)",
                    Type = -5,
                    ID = -9,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 5
                });
                foreach (var it in db.tbUserContractsAndMoalefeGhararDadi.GroupBy(s => s.FKMoalefeGhararDadi).ToList())
                {
                    var findca = db.tbContractMoalefeDastmozdi.Where(S => S.md_ID == it.Key).FirstOrDefault();
                    if (findca != null)
                    {
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = findca.md_Title + " (قرارداد های پرسنل )",
                            Type = -5,
                            ID = findca.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 5
                        });
                    }

                }

                #endregion
            }
            //تجهیزات
            else if (id == 6)
            {

                foreach (var it in db.tbEquipments.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " ( نام تجهیز )",
                        Type = -11,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " (  رقم تجهیز )",
                        Type = -12,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 11
                    });
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }


            }
            //بودجه
            else if (id == 7)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( بودجه )",
                        Type = -10,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    for (int i = 1; i <= 20; i++)
                    {
                        int sum = it.Pyman_ID;
                        sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                        // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                        string base9String = ConvertToBase9(sum);

                        // تبدیل string مبنای 9 به عدد int
                        long base9Int = long.Parse(base9String);

                        // اضافه کردن عدد 9 به انتهای عدد
                        string newBase9String = base9Int.ToString() + "9";
                        newBase9String = newBase9String.ToString() + i;
                        string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                        int base9Int2 = int.Parse(base9String2);
                        newBase9String = newBase9String.ToString() + "9";
                        newBase9String = newBase9String.ToString() + base9Int2;

                        // تبدیل دوباره به عدد int
                        xx = long.Parse(newBase9String);
                        long xxID = long.Parse(newBase9String);
                        if (base9Int == 1431)
                        {

                        }
                        shomareh += i;
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                            Type = -10000,
                            ID = xxID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                            Type = shomareh,
                            ID = -200,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                foreach (var it in db.tbCreditIndicators.ToList())
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                        Type = -10,
                        ID = it.ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }

            }
            //تزاز فیش
            else if (id == 8)
            {
                foreach (var it in db.tbDetermining_creditline.ToList())
                {
                    long xx = 0;

                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( سبد سایر عملیات )",
                        Type = -15,
                        ID = it.Determining_creditline_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    int shomareh = -330;
                    //for (int i = 1; i <= 20; i++)
                    //{
                    //    int sum = it.Pyman_ID;
                    //    sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                    //    // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                    //    string base9String = ConvertToBase9(sum);

                    //    // تبدیل string مبنای 9 به عدد int
                    //    long base9Int = long.Parse(base9String);

                    //    // اضافه کردن عدد 9 به انتهای عدد
                    //    string newBase9String = base9Int.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + i;
                    //    string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                    //    int base9Int2 = int.Parse(base9String2);
                    //    newBase9String = newBase9String.ToString() + "9";
                    //    newBase9String = newBase9String.ToString() + base9Int2;

                    //    // تبدیل دوباره به عدد int
                    //    xx = long.Parse(newBase9String);
                    //    long xxID = long.Parse(newBase9String);
                    //    if (base9Int == 1431)
                    //    {

                    //    }
                    //    shomareh += i;
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )" + xx,
                    //        Type = -10000,
                    //        ID = xxID,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //    myList.Add(new ListOfMoalefeha
                    //    {
                    //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                    //        Type = shomareh,
                    //        ID = -200,
                    //        ExistsFormula = false,
                    //        GharardadColoumnName = "usc_CountOfChild",
                    //        noghrogi = 10
                    //    });
                    //}

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    //    Type = -7,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    //    Type = -8,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});
                    //myList.Add(new ListOfMoalefeha
                    //{
                    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    //    Type = -9,
                    //    ID = it.ID,
                    //    ExistsFormula = false,
                    //    GharardadColoumnName = "usc_CountOfChild",
                    //    noghrogi = 6
                    //});

                }
                //foreach (var it in db.tbCreditIndicators.ToList())
                //{
                //    myList.Add(new ListOfMoalefeha
                //    {
                //        variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                //        Type = -10,
                //        ID = it.ID,
                //        ExistsFormula = false,
                //        GharardadColoumnName = "usc_CountOfChild",
                //        noghrogi = 10
                //    });

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    //    Type = -7,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    //    Type = -8,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});
                //    //myList.Add(new ListOfMoalefeha
                //    //{
                //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    //    Type = -9,
                //    //    ID = it.ID,
                //    //    ExistsFormula = false,
                //    //    GharardadColoumnName = "usc_CountOfChild",
                //    //    noghrogi = 6
                //    //});

                //}

            }

            else if (id == 9)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "کسورات",
                    Type = -13,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "اضافات",
                    Type = -13,
                    ID = -2,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });
            }

            else if (id == 10)
            {
                var find = db.tbFormula.Where(s => s.Frml_IsActive == true).ToList();
                foreach (var it in find)
                {
                    string type = "قراردادی";
                    switch (it.tbContractMoalefeDastmozdi.md_Type)
                    {
                        case (int)Type_Moalefe.Gharardadi:
                            type = "قراردادی";
                            break;
                        case (int)Type_Moalefe.Amalkardi:
                            type = "عملکردی";
                            break;
                        case (int)Type_Moalefe.Karbari:
                            type = "کاربری";
                            break;
                        case (int)Type_Moalefe.Sayer:
                            type = "سایر";
                            break;
                        case (int)Type_Moalefe.Controlli:
                            type = "کنترلی";
                            break;
                        case (int)Type_Moalefe.Dastmozdi:
                            type = "دستمزدی";
                            break;
                        case (int)Type_Moalefe.Calculational:
                            type = "محاسباتی";
                            break;
                        case (int)Type_Moalefe.Karkardi:
                            type = "کارکردی";
                            break;
                        case (int)Type_Moalefe.Pishkhan:
                            type = "پیشخوان";
                            break;
                        case (int)Type_Moalefe.Fish:
                            type = "فیش‏حقوقی";

                            break;
                        case (int)Type_Moalefe.sorat:
                            type = "صورت وضعیت";
                            break;
                    }
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(" + it.tbContractMoalefeDastmozdi.md_Title + ")" + " ( فرمول ها )",
                        Type = it.tbContractMoalefeDastmozdi.md_Type,
                        ID = (long)it.FKMoalefeDastMozdi,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                }


                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }
            else if (id == 11)
            {
                var find = db.tbinstallments.GroupBy(s => s.Fk_molfe).ToList();
                foreach (var it1 in find)
                {
                    var find2 = db.tbContractMoalefeDastmozdi.Where(s => s.md_ID == it1.Key).ToList();
                    foreach (var it in find2)
                    {
                        string type = "قراردادی";
                        //switch (it.tbContractMoalefeDastmozdi.md_Type)
                        //{
                        //    case (int)Type_Moalefe.Gharardadi:
                        //        type = "قراردادی";
                        //        break;
                        //    case (int)Type_Moalefe.Amalkardi:
                        //        type = "عملکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Karbari:
                        //        type = "کاربری";
                        //        break;
                        //    case (int)Type_Moalefe.Sayer:
                        //        type = "سایر";
                        //        break;
                        //    case (int)Type_Moalefe.Controlli:
                        //        type = "کنترلی";
                        //        break;
                        //    case (int)Type_Moalefe.Dastmozdi:
                        //        type = "دستمزدی";
                        //        break;
                        //    case (int)Type_Moalefe.Calculational:
                        //        type = "محاسباتی";
                        //        break;
                        //    case (int)Type_Moalefe.Karkardi:
                        //        type = "کارکردی";
                        //        break;
                        //    case (int)Type_Moalefe.Pishkhan:
                        //        type = "پیشخوان";
                        //        break;
                        //    case (int)Type_Moalefe.Fish:
                        //        type = "فیش‏حقوقی";

                        //        break;
                        //    case (int)Type_Moalefe.sorat:
                        //        type = "صورت وضعیت";
                        //        break;
                        //}
                        myList.Add(new ListOfMoalefeha
                        {
                            variablePersianName = "(" + it.md_Title + ")" + " ( فرمول ها )",
                            Type = -17,
                            ID = (long)it.md_ID,
                            ExistsFormula = false,
                            GharardadColoumnName = "usc_CountOfChild",
                            noghrogi = 10
                        });
                    }


                }


                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = "اضافات",
                //    Type = -13,
                //    ID = -2,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 4
                //});
            }

            else if (id == 12)
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "معادل کار",
                    Type = -18,
                    ID = -1,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 4
                });

            }

            //کاربران حقوقی
            string قراردادپرسنل = "قرارداد A"; // فقط نمونه است، مقدار واقعی رو باید مشخص کنید




            return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/GetAllMoalefeForFormulasbyid3.cshtml", myList);



            //return Json(myList, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GetAllMoalefeForFormulas()
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            myList.Add(new ListOfMoalefeha { variablePersianName = " تعداد اولاد مشمول (قراردادی) ", Type = 0, GharardadColoumnName = "usc_CountOfChild", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد (قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract" , noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ شروع قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_StartTime", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تاریخ پایان قرارداد(قراردادی)", Type = 0, GharardadColoumnName = "usc_EndTime", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "دوره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_DurationTime", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "شماره قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_ContractNumber", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "عنوان شغل ( قراردادی)", Type = 0, GharardadColoumnName = "usc_Jobtitle" , noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "نوع قرارداد ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TypeOfContract" , noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "تعداد سال سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "usc_TedadSalSanavat", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "مزد گروه ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueMozdGroup", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "مزد سنوات ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_ValueSanavat", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "خوار و بار ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_KharoBar", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "حق مسکن ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheMaskan", noghrogi = 1 });
            myList.Add(new ListOfMoalefeha { variablePersianName = "حق اولاد ( قراردادی)", Type = 0, GharardadColoumnName = "jobgroup_HagheOlad" , noghrogi=1 });
            foreach (var item in rep_moalefeDastmozdi.Update())
            {
                string type = "قراردادی";
                switch (item.md_Type)
                {
                    case (int)Type_Moalefe.Gharardadi:
                        type = "قراردادی";
                        break;
                    case (int)Type_Moalefe.Amalkardi:
                        type = "عملکردی";
                        break;
                    case (int)Type_Moalefe.Karbari:
                        type = "کاربری";
                        break;
                    case (int)Type_Moalefe.Sayer:
                        type = "سایر";
                        break;
                    case (int)Type_Moalefe.Controlli:
                        type = "کنترلی";
                        break;
                    case (int)Type_Moalefe.Dastmozdi:
                        type = "دستمزدی";
                        break;
                    case (int)Type_Moalefe.Calculational:
                        type = "محاسباتی";
                        break;
                    case (int)Type_Moalefe.Karkardi:
                        type = "کارکردی";
                        break;
                    case (int)Type_Moalefe.Pishkhan:
                        type = "پیشخوان";
                        break;
                    case (int)Type_Moalefe.Fish:
                        type = "فیش‏حقوقی";

                        break;
                    case (int)Type_Moalefe.sorat:
                        type = "صورت وضعیت";
                        break;
                }

                if (item.tbFormula.Any())
                {
                    myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = true, noghrogi = 1 });
                }
                else
                {
                    myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type, ExistsFormula = false, noghrogi = 1 });
                }
            }
            string قراردادپرسنل = "قرارداد A"; // فقط نمونه است، مقدار واقعی رو باید مشخص کنید

       
            //کاربران حقوقی
            #region             کاربران حقوقی
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام مدیر عامل  (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -1,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کد اختصاصی شرکت  (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -2,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });


            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام شرکت  (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -3,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شماره ثبت (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -4,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کد اقتصادی (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -5,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });


            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "محل ثبت شرکت  (قرارداد کاربران حقوقی)",
                Type = -3,
                ID = -6,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 3
            });
            #endregion

            #region             //پیمان ها 
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام پیمان (قرارداد پیمان ها)",
                Type = -4,
                ID = -2,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کد پروژه (قرارداد پیمان ها)",
                Type = -4,
                ID = -3,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "مخفف پیمان (قرارداد پیمان ها)",
                Type = -4,
                ID = -4,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کارفرما (قرارداد پیمان ها)",
                Type = -4,
                ID = -5,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کارفرما اصلی (قرارداد پیمان ها)",
                Type = -4,
                ID = -6,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شماره قرارداد (قرارداد پیمان ها)",
                Type = -4,
                ID = -7,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تاریخ شروع پیمان (قرارداد پیمان ها)",
                Type = -4,
                ID = -8,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تاریخ پایان پیمان (قرارداد پیمان ها)",
                Type = -4,
                ID = -9,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "مبلغ کل قرارداد (قرارداد پیمان ها)",
                Type = -4,
                ID = -10,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 4
            });


            foreach(var it in db.tbPeymanContractPrice.ToList())
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName =   "(" + it.tbPeymanContracts.pec_Title + ")",
                    Type = -600,
                    ID =(int) it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( شماره قرارداد پیمان  )",
                    Type = -601,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " ( مبلغ کل قرارداد )",
                    Type = -602,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه شروع قرارداد)",
                    Type = -603,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال شروع قرارداد)",
                    Type = -604,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                }); myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (ماه پایان قرارداد)",
                    Type = -605,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                }); myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (سال پایان قرارداد)",
                    Type = -606,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = "(" + it.tbPeymanContracts.pec_Title + ")" + " (کد پروژه)",
                    Type = -607,
                    ID = (int)it.FKPeymanID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Unit+"("+ it.tbPeymanContracts.pec_Title+")",
                    Type = -6,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                    Type = -7,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                    Type = -8,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                    Type = -9,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 6
                });

            }



            foreach (var it in db.tbDetermining_creditline.ToList())
            {
                long xx = 0;

                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( بودجه )",
                    Type = -10,
                    ID = it.Determining_creditline_ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 10
                });
                int shomareh = -330;
                for (int i = 1; i <= 20; i++)
                {
                    int sum = it.Pyman_ID;
                    sum = Math.Abs(sum); // تبدیل به عدد مثبت برای تبدیل به مبنای ۹

                    // تبدیل جمع به مبنای 9 (خروجی به صورت string)
                    string base9String = ConvertToBase9(sum);

                    // تبدیل string مبنای 9 به عدد int
                    long base9Int = long.Parse(base9String);

                    // اضافه کردن عدد 9 به انتهای عدد
                    string newBase9String = base9Int.ToString() + "9";
                     newBase9String = newBase9String.ToString() +i ;
                    string base9String2 = ConvertToBase9(it.tbCreditIndicators.ID);
                    int base9Int2 = int.Parse(base9String2);
                     newBase9String = newBase9String.ToString() + "9";
                    newBase9String = newBase9String.ToString() + base9Int2;

                    // تبدیل دوباره به عدد int
                    xx = long.Parse(newBase9String);
                   long xxID = long.Parse(newBase9String);
                    if(base9Int == 1431)
                    {

                    }
                    shomareh += i;
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = it.tbCreditIndicators.Title + "(" + it.tbPeymanContracts.pec_Title + ")" + "(صورت وضعیت " + i + ")" + " (  بودجه  rrrrrrr )"+xx,
                        Type = -10000,
                        ID = xxID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
                        Type = shomareh,
                        ID = -200,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 10
                    });
                }
                
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    Type = -7,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    Type = -8,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    Type = -9,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

            }
            foreach (var it in db.tbCreditIndicators.ToList())
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" */+ " (  بودجه  تست )",
                    Type = -10,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 10
                });

                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    Type = -7,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    Type = -8,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    Type = -9,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

            }


            foreach (var it in db.tbEquipments.ToList())
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")"+"(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"+ "(کد تجهیز :" + it.codtaghiz + ")" + " ( نام تجهیز )",
                    Type = -11,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 11
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.tbEquipmentBunch.Eqpbnch_Name + "(" + it.tbPeymanContracts.pec_Title + ")" + "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")" + "(کد تجهیز :" + it.codtaghiz + ")" + " (  رقم تجهیز )",
                    Type = -12,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 11
                });
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
                //    Type = -7,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
                //    Type = -8,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});
                //myList.Add(new ListOfMoalefeha
                //{
                //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
                //    Type = -9,
                //    ID = it.ID,
                //    ExistsFormula = false,
                //    GharardadColoumnName = "usc_CountOfChild",
                //    noghrogi = 6
                //});

            }




            //foreach (var it in db.tbCreditIndicators.ToList())
            //{
            //    int shomareh = -330;
            //    for(int i = 0; i <= 30; i++)
            //    {
            //        shomareh += i;
            //        myList.Add(new ListOfMoalefeha
            //        {
            //            variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
            //            Type = shomareh,
            //            ID = it.ID,
            //            ExistsFormula = false,
            //            GharardadColoumnName = "usc_CountOfChild",
            //            noghrogi = 10
            //        });
            //        myList.Add(new ListOfMoalefeha
            //        {
            //            variablePersianName =  "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
            //            Type = shomareh,
            //            ID = -200,
            //            ExistsFormula = false,
            //            GharardadColoumnName = "usc_CountOfChild",
            //            noghrogi = 10
            //        });
            //    }
               

            //    //myList.Add(new ListOfMoalefeha
            //    //{
            //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( نرخ )",
            //    //    Type = -7,
            //    //    ID = it.ID,
            //    //    ExistsFormula = false,
            //    //    GharardadColoumnName = "usc_CountOfChild",
            //    //    noghrogi = 6
            //    //});

            //    //myList.Add(new ListOfMoalefeha
            //    //{
            //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( تعداد )",
            //    //    Type = -8,
            //    //    ID = it.ID,
            //    //    ExistsFormula = false,
            //    //    GharardadColoumnName = "usc_CountOfChild",
            //    //    noghrogi = 6
            //    //});
            //    //myList.Add(new ListOfMoalefeha
            //    //{
            //    //    variablePersianName = it.Unit + "(" + it.tbPeymanContracts.pec_Title + ")" + " ( قیمت )",
            //    //    Type = -9,
            //    //    ID = it.ID,
            //    //    ExistsFormula = false,
            //    //    GharardadColoumnName = "usc_CountOfChild",
            //    //    noghrogi = 6
            //    //});

            //}
            //int shomareh23 = -330;

            //for (int i = 0; i <= 30; i++)
            //{
            //    shomareh23 += i;
            //    //myList.Add(new ListOfMoalefeha
            //    //{
            //    //    variablePersianName = it.Title + "(صورت وضعیت " + i + ")" + " (  بودجه  تست )",
            //    //    Type = shomareh23,
            //    //    ID = it.ID,
            //    //    ExistsFormula = false,
            //    //    GharardadColoumnName = "usc_CountOfChild",
            //    //    noghrogi = 10
            //    //});
            //    myList.Add(new ListOfMoalefeha
            //    {
            //        variablePersianName = "(رقم نهایی صورت وضعیت  " + i + ")" + " (  بودجه  تست )",
            //        Type = shomareh23,
            //        ID = -200,
            //        ExistsFormula = false,
            //        GharardadColoumnName = "usc_CountOfChild",
            //        noghrogi = 10
            //    });
            //}





            //foreach (var it in db.tbPeymanContracts.ToList())
            //{
            //    foreach (var item in rep_moalefeDastmozdi.Update())
            //    {
            //        string type = "قراردادی";
            //        switch (item.md_Type)
            //        {
            //            case (int)Type_Moalefe.Gharardadi:
            //                type = "قراردادی";
            //                break;
            //            case (int)Type_Moalefe.Amalkardi:
            //                type = "عملکردی";
            //                break;
            //            case (int)Type_Moalefe.Karbari:
            //                type = "کاربری";
            //                break;
            //            case (int)Type_Moalefe.Sayer:
            //                type = "سایر";
            //                break;
            //            case (int)Type_Moalefe.Controlli:
            //                type = "کنترلی";
            //                break;
            //            case (int)Type_Moalefe.Dastmozdi:
            //                type = "دستمزدی";
            //                break;
            //            case (int)Type_Moalefe.Calculational:
            //                type = "محاسباتی";
            //                break;
            //            case (int)Type_Moalefe.Karkardi:
            //                type = "کارکردی";
            //                break;
            //            case (int)Type_Moalefe.Pishkhan:
            //                type = "پیشخوان";
            //                break;
            //            case (int)Type_Moalefe.Fish:
            //                type = "فیش‏حقوقی";

            //                break;
            //            case (int)Type_Moalefe.sorat:
            //                type = "صورت وضعیت";
            //                break;
            //        }
            //        myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )" + "(" + it.pec_Title + ")", ID = item.md_ID, Type = -13, ExistsFormula = false, noghrogi = 1 });

            //    }
            //}


            #endregion
            #region شغل ها 
            foreach(var it  in db.tbdetailjob.ToList())
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.detail_title /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شغل  )",
                    Type = -14,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 14
                });

            }
            #endregion

            #region شهرستان ها 
            foreach (var it in db.tbCities.ToList())
            {
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  شهرستان ها   )",
                    Type = -15,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 14
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب شهرستان ها   )",
                    Type = -16,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 14
                });
                myList.Add(new ListOfMoalefeha
                {
                    variablePersianName = it.Name /*+ "(" + it.tbPeymanContracts.pec_Title + ")" *//*+*//* "(دسته :" + it.tbEquipmentBunch.tbEquipmentGroup.Eqpgrp_Name + ")"*/ + " (  ضریب مالیات شهرستان   )",
                    Type = -17,
                    ID = it.ID,
                    ExistsFormula = false,
                    GharardadColoumnName = "usc_CountOfChild",
                    noghrogi = 14
                });

            }
            #endregion

            #region قرارداد های پرسنل
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام (قرارداد پرسنل)",
                Type = -2,
                ID = -1,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام ونام خانوادگی (قرارداد پرسنل)",
                Type = -2,
                ID = -2,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کد پرسنلی (قرارداد پرسنل)",
                Type = -2,
                ID = -3,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });


            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "کد ملی (قرارداد پرسنل)",
                Type = -2,
                ID = -4,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شماره شناسنامه (قرارداد پرسنل)",
                Type = -2,
                ID = -5,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تاریخ تولد (قرارداد پرسنل)",
                Type = -2,
                ID = -6,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "نام پدر (قرارداد پرسنل)",
                Type = -2,
                ID = -7,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });


            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "محل تولد (قرارداد پرسنل)",
                Type = -2,
                ID = -8,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "محل صدور شناسنامه (قرارداد پرسنل)",
                Type = -2,
                ID = -9,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "جنسیت (قرارداد پرسنل)",
                Type = -2,
                ID = -10,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "وضعیت تاهل  (قرارداد پرسنل)",
                Type = -2,
                ID = -11,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شماره موبایل  (قرارداد پرسنل)",
                Type = -2,
                ID = -12,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شماره ثابت  (قرارداد پرسنل)",
                Type = -2,
                ID = -13,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تعداد اولاد (قرارداد پرسنل)",
                Type = -2,
                ID = -18,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تعداد اولاد  مشمول بیمه (قرارداد پرسنل)",
                Type = -2,
                ID = -14,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شهر محل خدمت  (قرارداد پرسنل)",
                Type = -2,
                ID = -15,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "ضزیب مالیات شهر محل خدمت  (قرارداد پرسنل)",
                Type = -2,
                ID = -18,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "ضزیب شهرستان شهر محل خدمت  (قرارداد پرسنل)",
                Type = -2,
                ID = -19,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "پیمان (قرارداد پرسنل)",
                Type = -2,
                ID = -16,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "شغل (قرارداد پرسنل)",
                Type = -2,
                ID = -17,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 2
            });










            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "حق سنوات (قرارداد های پرسنل)",
                Type = -5,
                ID = -1,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "حق اولاد (قرارداد های پرسنل )",
                Type = -5,
                ID = -2,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });

            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "خواربار (قرارداد های پرسنل)",
                Type = -5,
                ID = -3,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "مزد شغل (قرارداد های پرسنل)",
                Type = -5,
                ID = -4,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "مزد سایر (قرارداد های پرسنل)",
                Type = -5,
                ID = -10,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "تعداد سال سنوات (قرارداد های پرسنل)",
                Type = -5,
                ID = -11,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "حق مسکن (قرارداد های پرسنل)",
                Type = -5,
                ID = -5,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });


            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "ماه  شروع قرارداد  (قرارداد های پرسنل)",
                Type = -5,
                ID = -6,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "سال شروع قرارداد (قرارداد های پرسنل)",
                Type = -5,
                ID = -7,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            }); myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "ماه پایان قرارداد (قرارداد های پرسنل)",
                Type = -5,
                ID = -8,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            myList.Add(new ListOfMoalefeha
            {
                variablePersianName = "سال پایان قرارداد (قرارداد های پرسنل)",
                Type = -5,
                ID = -9,
                ExistsFormula = false,
                GharardadColoumnName = "usc_CountOfChild",
                noghrogi = 5
            });
            foreach (var it in db.tbUserContractsAndMoalefeGhararDadi.GroupBy(s => s.FKMoalefeGhararDadi).ToList())
            {
                var findca = db.tbContractMoalefeDastmozdi.Where(S => S.md_ID == it.Key).FirstOrDefault();
                if (findca != null)
                {
                    myList.Add(new ListOfMoalefeha
                    {
                        variablePersianName = findca.md_Title + " (قرارداد های پرسنل )",
                        Type = -5,
                        ID = findca.md_ID,
                        ExistsFormula = false,
                        GharardadColoumnName = "usc_CountOfChild",
                        noghrogi = 5
                    });
                }
               
            }

            #endregion
            return PartialView("_Exists_MoalefeDastmozdi", myList);
        }
        public static string ConvertToBase9(int number)
        {
            if (number == 0)
            {
                return "0"; // اگر عدد صفر بود، به سادگی "0" را برمی‌گردانیم
            }

            string result = "";

            // تبدیل به مبنای ۹ با تقسیم باقی‌مانده
            while (number > 0)
            {
                int remainder = number % 9;  // باقی‌مانده تقسیم بر ۹
                result = remainder.ToString() + result;  // باقی‌مانده را به ابتدای رشته اضافه می‌کنیم
                number /= 9;  // تقسیم عدد بر ۹
            }

            return result;
        }

        [AuthorizeAAA]
        public ActionResult GetAllMoalefeForFormulas_Select()
        {
            List<ListOfMoalefeha> myList = new List<ListOfMoalefeha>();
            //myList.Add(new ListOfMoalefeha
            //{
            //    variablePersianName = " تعداد اولاد مشمول (قراردادی) ",
            //    Type = 0,
            //    GharardadColoumnName = "usc_CountOfChild"
            //});
            //myList.Add(new ListOfMoalefeha
            //{
            //    variablePersianName = "نوع قرارداد (قراردادی)",
            //    Type = 0,
            //    GharardadColoumnName = "usc_TypeOfContract"
            //});
            //myList.Add(new ListOfMoalefeha
            //{
            //    variablePersianName = "تاریخ شروع قرارداد(قراردادی)",
            //    Type = 0,
            //    GharardadColoumnName = "usc_StartTime"
            //});
            //myList.Add(new ListOfMoalefeha
            //{
            //    variablePersianName = "تاریخ پایان قرارداد(قراردادی)",
            //    Type = 0,
            //    GharardadColoumnName = "usc_EndTime"
            //});
            //myList.Add(new ListOfMoalefeha
            //{
            //    variablePersianName = "دوره قرارداد ( قراردادی)",
            //    Type = 0,
            //    GharardadColoumnName = "usc_DurationTime"
            //});

            foreach (var item in rep_moalefeDastmozdi.Update())
            {
                string type = "قراردادی";
                switch (item.md_Type)
                {
                    case (int)Type_Moalefe.Gharardadi:
                        type = "قراردادی";
                        break;
                    case (int)Type_Moalefe.Amalkardi:
                        type = "عملکردی";
                        break;
                    case (int)Type_Moalefe.Karbari:
                        type = "کاربری";
                        break;
                    case (int)Type_Moalefe.Sayer:
                        type = "سایر";
                        break;
                    case (int)Type_Moalefe.Controlli:
                        type = "کنترلی";
                        break;
                    case (int)Type_Moalefe.Dastmozdi:
                        type = "دستمزدی";
                        break;
                    case (int)Type_Moalefe.Calculational:
                        type = "محاسباتی";
                        break;
                    case (int)Type_Moalefe.Karkardi:
                        type = "کارکردی";
                        break;
                    case (int)Type_Moalefe.Fish:
                        type = "فیش‏حقوقی";
                        break;
                    case (int)Type_Moalefe.sorat:
                        type = "صورت وضعیت";
                        break;
                }

                myList.Add(new ListOfMoalefeha { variablePersianName = item.md_Title + " ( " + type + " )", ID = item.md_ID, Type = item.md_Type });
            }

            return PartialView("GetAllMoalefeForFormulas_Select", myList);
        }

        [AuthorizeAAA]
        public ActionResult ExportAddMoalefeExcel()
        {
            var OutPutFile = SetDataExcel_AddMoalefe();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            return File(stream.ToArray(), mimeType, "AddMoalefe" + extension);
        }

        [AuthorizeAAA]
        public Workbook SetDataExcel_AddMoalefe()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/AddMoalefe.xlsx"));
            Row Row;

            //name and lastname and groupjob
            int counter = 1;
            var category = categoryRepo.GetAllActiveCategory();
            foreach (var item in category)
            {
                Row = new Row() { Height = 20, Index = counter };


                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Category_Name,
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
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            int counterunit = 1;
            var unit = UnitRepo.ListActiveUnit();
            foreach (var item in unit)
            {
                Row = new Row() { Height = 20, Index = counterunit };


                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.unt_Name,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 3
                    }
                });
                counterunit++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }
            return Moalefeexcelfile;
        }


        [AuthorizeAAA]
        [HttpPost]
        public ActionResult ImportExcel_AddMoalefe(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_Moalefe(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }















        [AuthorizeAAA]
        public string GetDataFromExcel_Moalefe(HttpPostedFileBase MyExcelStream)
        {

            string name, type, category, unit, jens = "";
            float value;
            List<tbContractMoalefeDastmozdi> lstMoalefe = new List<tbContractMoalefeDastmozdi>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets > 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        int Cols = 0;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
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
                                var row = workbook.Sheets[0].Rows[i];
                                var Name = row.Cells[0];
                                if (Name.Value != null || Name.Value.ToString() != "")
                                {
                                    name = Name.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                }

                                ///////////////
                                var Type = row.Cells[1];
                                if (Type.Value != null || Type.Value.ToString() != "")
                                {
                                    type = Type.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون نوع مولفه سطر " + " ( " + i + " ) " + "را پر کنید ";
                                }

                                ////////////////////////
                                var Category = row.Cells[2];
                                if (Category.Value != null || Category.Value.ToString() != "")
                                {
                                    category = Category.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون دسته بندی مولفه سطر " + " ( " + i + " ) " + "را پر کنید ";
                                }
                                var Unit = row.Cells[4];
                                if (Unit.Value != null || Unit.Value.ToString() != "")
                                {
                                    unit = Unit.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون واحد مولفه سطر " + " ( " + i + " ) " + "را پر کنید ";
                                }
                                ///////////////////////////////
                                var Jens = row.Cells[3];
                                if (Jens.Value != null || Jens.Value.ToString() != "")
                                {
                                    jens = Jens.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون جنس مولفه سطر " + " ( " + i + " ) " + "را پر کنید ";
                                }
                                var unitID = db.tbUnitParameter.FirstOrDefault(p => p.unt_Name == unit).unt_ID;
                                var categoryID = db.tbCategories.FirstOrDefault(p => p.Category_Name == category).Category_ID;
                                var jensvalue = "";
                                switch (jens)
                                {
                                    case "عدد":
                                        jensvalue = "int";
                                        break;
                                    case "رشته":
                                        jensvalue = "string";
                                        break;
                                    case "تاریخ":
                                        jensvalue = "date";
                                        break;
                                    case "زمان":
                                        jensvalue = "time";
                                        break;
                                    case "تیک":
                                        jensvalue = "tik";
                                        break;
                                    default:
                                        break;
                                }

                                int typeID = 0;
                                switch (type)
                                {
                                    case "قراردادی":
                                        typeID = (int)Type_Moalefe.Gharardadi;
                                        break;
                                    case "عملکردی":
                                        typeID = (int)Type_Moalefe.Amalkardi;
                                        break;
                                    case "کاربری":
                                        typeID = (int)Type_Moalefe.Karbari;
                                        break;
                                    case "سایر":
                                        typeID = (int)Type_Moalefe.Sayer;
                                        break;
                                    case "کنترلی":
                                        typeID = (int)Type_Moalefe.Controlli;
                                        break;
                                    case "دستمزدی":
                                        typeID = (int)Type_Moalefe.Dastmozdi;
                                        break;
                                    case "کارکردی":
                                        typeID = (int)Type_Moalefe.Karkardi;
                                        break;
                                    case "محاسباتی":
                                        typeID = (int)Type_Moalefe.Calculational;
                                        break;
                                    case
                                        "پیشخوان":
                                        typeID = (int)Type_Moalefe.Pishkhan;
                                        break;
                                    case
                                        "فیش‏حقوقی":
                                        typeID = (int)Type_Moalefe.Fish;
                                        break;
                                    case
                                        "صورت وضعیت":
                                        typeID = (int)Type_Moalefe.sorat;
                                        break;
                                }

                                tbContractMoalefeDastmozdi moalefe = new tbContractMoalefeDastmozdi
                                {
                                    FK_Category_ID = categoryID,
                                    md_IsActive = true,
                                    md_Title = name,
                                    md_Type = typeID,
                                    md_variable = jensvalue,
                                    FK_UnitParameter = unitID
                                };
                                lstMoalefe.Add(moalefe);
                            }

                            if (!rep_moalefeDastmozdi.AddRange(lstMoalefe))
                            {
                                transaction.Rollback();
                                return "ثبت کردن مولفه با خطا مواجه شد";
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
                    return "با موفقیت انجام شد";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return ex.Message;
                }
            }
        }

        [AuthorizeAAA]
        public ActionResult MoalefeKarbari_Select()
        {
            var moalefeKarbari = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == (int)Type_Moalefe.Karbari).ToList();
            if (moalefeKarbari != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_MoalefeKarbari_Select.cshtml", moalefeKarbari);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_MoalefeKarbari_Select.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }

        [AuthorizeAAA]
        public ActionResult AllMoalefeList()
        {
            var Model = rep_moalefeDastmozdi.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }
        public ActionResult AllMoalefeListAnddetai()
        {
            var Model = rep_moalefeDastmozdi.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeListAnddetai.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeListAnddetai.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }
        public ActionResult AllMoalefeListAnddetaiazefat()
        {
            var Model = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
            if (Model != null)
            {
                var x = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title== "کمک هزینه ایاب و ذهاب(ریال)").FirstOrDefault();
                if (x != null)
                {
                    Model.Add(x);

                }
                var find = db.tbUserContractsAndMoalefeGhararDadi.GroupBy(s => s.FKMoalefeGhararDadi).ToList();
               foreach(var it in find)
                {
                    var findmoal= Model.Where(s=>s.md_ID==it.Key).FirstOrDefault();
                    if(findmoal != null)
                    {
                        Model.Remove(findmoal);
                    }
                }
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeListAnddetaiazefat.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeListAnddetaiazefat.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }
        [AuthorizeAAA]
        public ActionResult AllMoalefeList_MultiSelect()
        {
            var Model = rep_moalefeDastmozdi.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }
        public ActionResult AllMoalefeList_MultiSelectfish()
        {
            var Model = db.tbContractMoalefeDastmozdi.Where(p=>p.md_Type==10).ToList();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelectfish.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelectfish.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }


        [AuthorizeAAA]
        public ActionResult ControllyMoalefeList()
        {
            var Model = rep_moalefeDastmozdi.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }
        public ActionResult AllMoalefeList_MultiSelectForTOOl(int IsFor = 0)
        {

            if (IsFor == 3)
            {
                var exist = db.tbContractMoalefeDastmozdi
                    .Where(p => p.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان")
                    .ToList();

                if (exist != null)
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", exist);
                }
                else
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
                }
            }
            else
            {
                var exist = db.tbContractMoalefeDastmozdi
                    .Where(p => p.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ابزار پیمان")
                    .ToList();

                if (exist != null)
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", exist);
                }
                else
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
                }
            }

        }
        public ActionResult AllMoalefeList_MultiSelectForproje(int IsFor = 0)
        {
            if (IsFor == 6)
            {
                var exist = db.tbContractMoalefeDastmozdi
              .Where(p => p.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان")
              .ToList();

                if (exist != null)
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", exist);
                }
                else 
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
                }
            }
            else if (IsFor == 7)
            {
                var exist = db.tbContractMoalefeDastmozdi
          .Where(p => p.md_Title == "پیشخوان-پیمان-مستندات وضعیت-فیش واریز لیست تامین اجتماعی")
          .ToList();

                if (exist != null)
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", exist);
                }
                else
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
                }
            }
            else
            {
                var exist = db.tbContractMoalefeDastmozdi
      .Where(p => p.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان")
      .ToList();

                if (exist != null)
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", exist);
                }
                else
                {
                    return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/AllMoalefeList_MultiSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
                }
            }
        }

        [AuthorizeAAA]
        public ActionResult ControllyMoalefeList_SingleSelect()
        {
            var Model =db.tbContractMoalefeDastmozdi.ToList();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList_SingleSelect.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList_SingleSelect.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }

        [AuthorizeAAA]
        public ActionResult ControllyMoalefeList2()
        {
            var Model = rep_moalefeDastmozdi.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList2.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_ControllyMoalefeList2.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }

        [AuthorizeAAA]
        public ActionResult PanelMoalefeList()
        {
            var Model = PanelsRepo.Listt();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PanelMoalefe.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PanelMoalefe.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }

        [AuthorizeAAA]
        public ActionResult PeymanList()
        {
            var Model = PaymanRepo.Update();
            if (Model != null)
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PeymanTitleList.cshtml", Model);
            }
            else
            {
                return PartialView("~/Areas/Contracts/Views/MoalefeDastMozdi/_PeymanTitleList.cshtml", new List<tbContractMoalefeDastmozdi>());
            }
        }

        [AuthorizeAAA]
        public ActionResult PanelControlly_Submit(string moalefe, string[] CotrollyMoalefe)
        {
            for (int i = 0; i < CotrollyMoalefe.Length; i++)
            {
                tbControllyPanel st = new tbControllyPanel();
                st.FK_PanelsContractId = Int32.Parse(moalefe);
                st.FK_ControllyMoalefeId = Int32.Parse(CotrollyMoalefe[i]);

                ControllyPanelRepo.Create(st);
            }
            ControllyPanelRepo.SaveChanges();
            return Content("1");
        }

        [AuthorizeAAA]
        public ActionResult PeymanControlly_Submit(string peyman, string[] CotrollyMoalefe)
        {
            for (int i = 0; i < CotrollyMoalefe.Length; i++)
            {
                tbControllyPeyman st = new tbControllyPeyman();
                st.FK_PeymanId = Int32.Parse(peyman);
                st.FK_ContractControllyId = Int32.Parse(CotrollyMoalefe[i]);

                ControllyPeymanRepo.Create(st);
            }
            ControllyPeymanRepo.SaveChanges();
            return Content("1");
        }

        #endregion
    }
}