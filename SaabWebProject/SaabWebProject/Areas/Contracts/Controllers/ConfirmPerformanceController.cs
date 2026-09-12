using ExcelLibrary.BinaryFileFormat;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Models.ViewModels.ConfirmPerformance;
using SaabWebProject.Models.ViewModels.Contracts.Function;
using SaabWebProject.Utility;
using Stimulsoft.Blockly.Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using static Stimulsoft.Report.StiRecentConnections;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class ConfirmPerformanceController : Controller
    {

        #region متغیر
        SaabEntities db = new SaabEntities();
        private tbContractMoalefeDastMozdiRepository rep_moalefeDastmozdi = new tbContractMoalefeDastMozdiRepository();
        RegistrationAndConfirmationProceduresController registerController = new RegistrationAndConfirmationProceduresController();


        #endregion
        #region سازنده

        #endregion
        // GET: Contracts/ConfirmPerformance
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult _TaedKarkard()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedKarkard.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult _TaedKarkardDetail()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedKarkardDetail.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult _TaedKarkardTajhizat()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedKarkardTajhizat.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult _TaedKarkardDetailKarkard()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedKarkardDetailKarkard.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult _TaedVaziyatProjectKalaye()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedVaziyatProjectKalaye.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult _TaedVaziyatDetailProjectKalaye()
        {
            return View("~/Areas/Contracts/Views/ConfirmPerformance/_TaedVaziyatDetailProjectKalaye.cshtml");

        }

        public ActionResult Filter(int Py = 0, int Basteh = 0)
        {
            ModelFilter mo = new ModelFilter();
            mo.fkBasteh = Basteh;
            mo.fkpy = Py;
            return View(mo);
        }
        public ActionResult viewoneee(int fkmoalfe = 0, int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == Py && p.Inactive != true).FirstOrDefault();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();

            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.tbPeymanContracts.Inactive != true && p.Status==true).Select(p => p.FK_User_ID).ToList();
            var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p=>p.md_ID==fkmoalfe).ToList();

            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh && s.Deleted != true).FirstOrDefault();

            List<string> MoalefeKarkardi = new List<string>();
            List<int> MoalefeKarkardi11 = new List<int>();

            MoalefeKarkardi.Add(moalfe.FirstOrDefault(p => p.md_ID == fkmoalfe).md_Title);

            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }




            FunctionModel2 lstMoalefeexcel2 = new FunctionModel2();
            DateTime FromDate = default; DateTime ToDate = default;
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = System.Convert.ToInt32(Month),
                MoalfeVal_Year = System.Convert.ToInt32(Year),
                PeymanName = PymN.pec_Title,
                BastehName = BastehName2.Title,
                FromDate = FromDate,
                ToDate = ToDate,


            };
            var exi = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month)).OrderByDescending(s => s.svdfunc_ID).FirstOrDefault();
            if (exi != null)
            {
                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = moalfe.Where(p => p.md_ID == item2.md_ID).FirstOrDefault().md_ID;
                        var Value = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 && p.FK_SavedFunctionsID == exi.svdfunc_ID && p.MoalfeVal_FKUser == item.usr_ID).Select(s => s.MoalfeVal_Value).FirstOrDefault();
                        double value; // Change the type to double

                        //if (Value == null)
                        //{
                        //    value = 0;
                        //}
                        //else
                        //{
                        //    value = Value.Value;
                        //}

                        var stringValue = Value.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {

                            MoalefeTitle = moalfe.Where(p => p.md_ID == item2.md_ID).FirstOrDefault().md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = MoalefeID2

                        });
                    }
                    lstMoalefeexcel2.UsersMoalefe2.Add(obj); //اطلاعات یک نفر اضافه شود 

                }


            }


            return PartialView("~/Areas/Salaries/Views/Moalefe/Viewonemoalfe.cshtml", lstMoalefeexcel2);
        }

        public ActionResult viewone(int fkusr=0, int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == Py && p.Inactive != true).FirstOrDefault();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();

            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true && p.tbPeymanContracts.Inactive != true).Select(p => p.FK_User_ID).ToList();
            var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);

            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh&&s.Deleted != true).FirstOrDefault();

            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(moalfe.FirstOrDefault(p => p.md_ID == item).md_Title);
            }


            List<tbUsers> UserList = new List<tbUsers>();
            //foreach (var item in usrID)
            //{
            //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    UserList.Add(user);
            //}


            var usr = db.tbUsers.Where(p => p.usr_ID == fkusr).FirstOrDefault();
            UserList.Add(usr);

            FunctionModel2 lstMoalefeexcel2 = new FunctionModel2();
            DateTime FromDate = default; DateTime ToDate = default;
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = System.Convert.ToInt32(Month),
                MoalfeVal_Year = System.Convert.ToInt32(Year),
                PeymanName = PymN.pec_Title,
                BastehName = BastehName2.Title,
                FromDate = FromDate,
                ToDate = ToDate,


            };
            var exi = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month)).OrderByDescending(s => s.svdfunc_ID).FirstOrDefault();
            if (exi != null)
            {
                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_ID;
                        var Value = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 && p.FK_SavedFunctionsID == exi.svdfunc_ID && p.MoalfeVal_FKUser == item.usr_ID).Select(s => s.MoalfeVal_Value).FirstOrDefault();
                        double value; // Change the type to double

                        //if (Value == null)
                        //{
                        //    value = 0;
                        //}
                        //else
                        //{
                        //    value = Value.Value;
                        //}

                        var stringValue = Value.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {

                            MoalefeTitle = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = MoalefeID2

                        });
                    }
                    lstMoalefeexcel2.UsersMoalefe2.Add(obj); //اطلاعات یک نفر اضافه شود 

                }


            }


            return PartialView("~/Areas/Salaries/Views/Moalefe/Viewonepersone.cshtml", lstMoalefeexcel2);
        }


        public ActionResult Viewonmoalanus(int fkmoalfe = 0, int fkusr = 0, int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == Py && p.Inactive != true).FirstOrDefault();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.tbPeymanContracts.Inactive != true && p.Status == true).Select(p => p.FK_User_ID).ToList();
            var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_ID == fkmoalfe).Select(o=>o.md_ID).ToList();
            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh && s.Deleted != true).FirstOrDefault();

            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(moalfe.FirstOrDefault(p => p.md_ID == item).md_Title);
            }


            List<tbUsers> UserList = new List<tbUsers>();
            //foreach (var item in usrID)
            //{
            //    var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
            //    UserList.Add(user);
            //}


            var usr = db.tbUsers.Where(p => p.usr_ID == fkusr).FirstOrDefault();
            UserList.Add(usr);

            FunctionModel2 lstMoalefeexcel2 = new FunctionModel2();
            DateTime FromDate = default; DateTime ToDate = default;
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = System.Convert.ToInt32(Month),
                MoalfeVal_Year = System.Convert.ToInt32(Year),
                PeymanName = PymN.pec_Title,
                BastehName = BastehName2.Title,
                FromDate = FromDate,
                ToDate = ToDate,


            };
            var exi = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month)).OrderByDescending(s => s.svdfunc_ID).FirstOrDefault();
            if (exi != null)
            {
                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_ID;
                        var Value = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 && p.FK_SavedFunctionsID == exi.svdfunc_ID && p.MoalfeVal_FKUser == item.usr_ID).Select(s => s.MoalfeVal_Value).FirstOrDefault();
                        double value; // Change the type to double

                        //if (Value == null)
                        //{
                        //    value = 0;
                        //}
                        //else
                        //{
                        //    value = Value.Value;
                        //}

                        var stringValue = Value.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {

                            MoalefeTitle = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = MoalefeID2

                        });
                    }
                    lstMoalefeexcel2.UsersMoalefe2.Add(obj); //اطلاعات یک نفر اضافه شود 

                }


            }


            return PartialView("~/Areas/Salaries/Views/Moalefe/ViewForonepersononemoalfe.cshtml", lstMoalefeexcel2);
        }

        public ActionResult viewforAccept(int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = db.tbPeymanContracts.Where(p => p.pec_ID == Py && p.Inactive != true).FirstOrDefault();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();

            List<int?> listpymn = new List<int?>();
            for (int i = 0; i < pymn.Count(); i++)
            {
                listpymn.Add(System.Convert.ToInt32(pymn[i]));
            }
            var usrID = db.Link_User_And_Peyman.Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.tbPeymanContracts.Inactive != true && p.Status == true).Select(p => p.FK_User_ID).ToList();
            var MoalefeID = registerController.GetAllMoalefeInReffrenceSaveLevel(Basteh);

            var BastehName2 = db.tbReffrenceSaveLevel.Where(s => s.ID == Basteh && s.Deleted != true).FirstOrDefault();

            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(moalfe.FirstOrDefault(p => p.md_ID == item).md_Title);
            }


            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }
            FunctionModel2 lstMoalefeexcel2 = new FunctionModel2();
            DateTime FromDate = default; DateTime ToDate = default;
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = System.Convert.ToInt32(Month),
                MoalfeVal_Year = System.Convert.ToInt32(Year),
                PeymanName = PymN.pec_Title,
                BastehName = BastehName2.Title,
                FromDate = FromDate,
                ToDate = ToDate,


            };
            var exi = db.tbSavedFunctions.Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true && p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month)).OrderByDescending(s => s.svdfunc_ID).FirstOrDefault();
            if (exi != null)
            {
                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_ID;
                        var Value = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 && p.FK_SavedFunctionsID == exi.svdfunc_ID && p.MoalfeVal_FKUser == item.usr_ID).Select(s => s.MoalfeVal_Value).FirstOrDefault();
                        double value; // Change the type to double

                        //if (Value == null)
                        //{
                        //    value = 0;
                        //}
                        //else
                        //{
                        //    value = Value.Value;
                        //}

                        var stringValue = Value.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {

                            MoalefeTitle = moalfe.Where(p => p.md_ID == item2).FirstOrDefault().md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = MoalefeID2

                        });
                    }
                    lstMoalefeexcel2.UsersMoalefe2.Add(obj); //اطلاعات یک نفر اضافه شود 

                }


            }


            return PartialView("~/Areas/Salaries/Views/Moalefe/ViewforAcceptMoalfe.cshtml", lstMoalefeexcel2);
        }
        public async Task<ActionResult> viewforAccept2(int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = await db.tbPeymanContracts.FirstOrDefaultAsync(p => p.pec_ID == Py && p.Inactive != true);

            List<int?> listpymn = pymn.Select(p => (int?)System.Convert.ToInt32(p)).ToList();
            var usrID = await db.Link_User_And_Peyman
                .Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true)
                .Select(p => p.FK_User_ID)
                .ToListAsync();
            var moalfe = db.tbContractMoalefeDastmozdi.ToList();
            var MoalefeID = await registerController.GetAllMoalefeInReffrenceSaveLevel2(Basteh); 
            var BastehName2 = await db.tbReffrenceSaveLevel.FirstOrDefaultAsync(s => s.ID == Basteh && s.Deleted != true);

            var MoalefeKarkardi = MoalefeID.Select(item => db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item)?.md_Title).ToList();

            var UserList = await db.tbUsers.Where(p => usrID.Contains(p.usr_ID)).ToListAsync();

            var lstMoalefeexcel2 = new FunctionModel2();
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = Month,
                MoalfeVal_Year = Year,
                PeymanName = PymN?.pec_Title,
                BastehName = BastehName2?.Title,
                FromDate = default,
                ToDate = default
            };

            var exi = await db.tbSavedFunctions
                .Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true &&
                            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month))
                .OrderByDescending(s => s.svdfunc_ID)
                .FirstOrDefaultAsync();

            if (exi != null)
            {
                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = moalfe.FirstOrDefault(p => p.md_ID == item2)?.md_ID;
                        var Value = await db.tbMoalefeDastmozdiValueFromExcel
                            .Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 &&
                                        p.FK_SavedFunctionsID == exi.svdfunc_ID &&
                                        p.MoalfeVal_FKUser == item.usr_ID)
                            .Select(s => s.MoalfeVal_Value)
                            .FirstOrDefaultAsync();
                        var stringValue = Value?.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {
                            MoalefeTitle = moalfe.FirstOrDefault(p => p.md_ID == item2)?.md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = (int)MoalefeID2
                        });
                    }
                    lstMoalefeexcel2.UsersMoalefe2.Add(obj);
                }
            }

            return PartialView("~/Areas/Salaries/Views/Moalefe/ViewforAcceptMoalfe.cshtml", lstMoalefeexcel2);
        }
        public async Task<ActionResult> viewforAccept4(int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = await db.tbPeymanContracts.FirstOrDefaultAsync(p => p.pec_ID == Py && p.Inactive != true);

            List<int?> listpymn = pymn.Select(p => (int?)System.Convert.ToInt32(p)).ToList();
            var usrID = await db.Link_User_And_Peyman
                .Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.Status == true)
                .Select(p => p.FK_User_ID)
                .ToListAsync();

            var MoalefeID = await registerController.GetAllMoalefeInReffrenceSaveLevel2(Basteh);
            var BastehName2 = await db.tbReffrenceSaveLevel.FirstOrDefaultAsync(s => s.ID == Basteh && s.Deleted != true);

            var MoalefeKarkardi = MoalefeID.Select(item => db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item)?.md_Title).ToList();

            var UserList = await db.tbUsers.Where(p => usrID.Contains(p.usr_ID)).ToListAsync();

            var lstMoalefeexcel2 = new FunctionModel2();
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = Month,
                MoalfeVal_Year = Year,
                PeymanName = PymN?.pec_Title,
                BastehName = BastehName2?.Title,
                FromDate = default,
                ToDate = default
            };

            var exi = await db.tbSavedFunctions
                .Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true &&
                            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month))
                .OrderByDescending(s => s.svdfunc_ID)
                .FirstOrDefaultAsync();

            if (exi != null)
            {
                var userMoalefeList = new List<MoalefeUserInfo2>();

                foreach (var item in UserList)
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();
                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item2)?.md_ID;
                        var Value = await db.tbMoalefeDastmozdiValueFromExcel
                            .Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 &&
                                        p.FK_SavedFunctionsID == exi.svdfunc_ID &&
                                        p.MoalfeVal_FKUser == item.usr_ID)
                            .Select(s => s.MoalfeVal_Value)
                            .FirstOrDefaultAsync();
                        var stringValue = Value?.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {
                            MoalefeTitle = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item2)?.md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = (int)MoalefeID2
                        });
                    }
                    userMoalefeList.Add(obj);

                }
                lstMoalefeexcel2.UsersMoalefe2.AddRange(userMoalefeList);

            }

            return PartialView("~/Areas/Salaries/Views/Moalefe/ViewforAcceptMoalfe.cshtml", lstMoalefeexcel2);
        }
        public async Task<ActionResult> viewforAccept3(int Py = 0, int Basteh = 0, int Year = 0, int Month = 0)
        {
            var pymn = Py.ToString().Split(',');
            var PymN = await db.tbPeymanContracts.FirstOrDefaultAsync(p => p.pec_ID == Py && p.Inactive != true);

            List<int?> listpymn = pymn.Select(p => (int?)System.Convert.ToInt32(p)).ToList();
            var usrID = await db.Link_User_And_Peyman
                .Where(p => listpymn.Contains(p.FK_Peyman_ID) && p.tbPeymanContracts.Inactive != true && p.Status == true)
                .Select(p => p.FK_User_ID)
                .ToListAsync();

            var MoalefeID = await registerController.GetAllMoalefeInReffrenceSaveLevel2(Basteh);
            var BastehName2 = await db.tbReffrenceSaveLevel.FirstOrDefaultAsync(s => s.ID == Basteh && s.Deleted != true);

            var MoalefeKarkardi = MoalefeID.Select(item => db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item)?.md_Title).ToList();

            var UserList = await db.tbUsers.Where(p => usrID.Contains(p.usr_ID)).ToListAsync();

            var lstMoalefeexcel2 = new FunctionModel2();
            lstMoalefeexcel2.UsersMoalefe2 = new List<MoalefeUserInfo2>();
            lstMoalefeexcel2.SabtHeader = new Header
            {
                MoalfeVal_Month = Month,
                MoalfeVal_Year = Year,
                PeymanName = PymN?.pec_Title,
                BastehName = BastehName2?.Title,
                FromDate = default,
                ToDate = default
            };

            var exi = await db.tbSavedFunctions
                .Where(p => p.svdfunc_BastehID == Basteh && p.svdfunc_pymnID == Py && p.tbPeymanContracts.Inactive != true &&
                            p.tbMoalefeDastmozdiValueFromExcel.Any(m => m.MoalfeVal_Year == Year && m.MoalfeVal_Month == Month))
                .OrderByDescending(s => s.svdfunc_ID)
                .FirstOrDefaultAsync();

            if (exi != null)
            {
                var userMoalefeList = new List<MoalefeUserInfo2>();

                Parallel.ForEach(UserList, async item =>
                {
                    var obj = new MoalefeUserInfo2();
                    obj.FullName = item.FullName;
                    obj.PersonalCode = System.Convert.ToInt32(item.usr_Personal_ID);
                    obj.listMoalefe = new List<MoalefeInfo2>();

                    foreach (var item2 in MoalefeID)
                    {
                        var MoalefeID2 = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item2)?.md_ID;
                        var Value = await db.tbMoalefeDastmozdiValueFromExcel
                            .Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID2 &&
                                        p.FK_SavedFunctionsID == exi.svdfunc_ID &&
                                        p.MoalfeVal_FKUser == item.usr_ID)
                            .Select(s => s.MoalfeVal_Value)
                            .FirstOrDefaultAsync();
                        var stringValue = Value?.ToString();
                        obj.listMoalefe.Add(new MoalefeInfo2
                        {
                            MoalefeTitle = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item2)?.md_Title,
                            MoalefeValue = stringValue,
                            Moalefenum = (int)MoalefeID2
                        });
                    }
                    userMoalefeList.Add(obj);
                });

                lstMoalefeexcel2.UsersMoalefe2.AddRange(userMoalefeList);
            }

            return PartialView("~/Areas/Salaries/Views/Moalefe/ViewforAcceptMoalfe.cshtml", lstMoalefeexcel2);
        }


        public ActionResult ViewForAcceptModal(int Moalfe, int name, string Py, string Basteh, int Year, int Month)
        {


            //if (string.Equals(c, X2, StringComparison.OrdinalIgnoreCase))
            //{
            //    var MoalefeID8 = db.tbContractMoalefeDastmozdi
            //        .Where(p => string.Equals(p.md_Title.Trim(), Moalfe, StringComparison.OrdinalIgnoreCase))
            //        .FirstOrDefault();
            //}
            var pyy = db.tbPeymanContracts.Where(p => p.pec_Title == Py && p.Inactive != true).Select(s => s.pec_ID).FirstOrDefault();
            var Bas = db.tbReffrenceSaveLevel.Where(p => p.Title == Basteh && p.tbReffrenceSave.FK_PeymanID == pyy && p.Deleted != true).Select(s => s.ID).FirstOrDefault();

            viewModel f = new viewModel();


            var usr2 = db.tbUsers.Where(p => p.usr_Personal_ID == name).FirstOrDefault();
            var usr = 1710;
            //f.usr = usr;
            var x = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == Moalfe && p.MoalfeVal_FKUser == usr2.usr_ID && p.MoalfeVal_Year == Year && p.MoalfeVal_Month == Month
            && p.tbSavedFunctions.svdfunc_pymnID == pyy && p.tbSavedFunctions.tbPeymanContracts.Inactive != true && p.tbSavedFunctions.svdfunc_BastehID == Bas
            ).ToList();

            foreach (var r in x)
            {
                var ex = db.tbSavedFunctions.Where(p => p.svdfunc_ID == r.FK_SavedFunctionsID).FirstOrDefault();
                var us = db.tbUsers.Where(p => p.usr_ID == ex.svdfunc_UserSaveID).FirstOrDefault();
                if (ex.For_Accept == true)
                {
                    f.AddData(us.FullName, "تایید کننده", (double)r.MoalfeVal_Value);

                }
                else
                {
                    f.AddData(us.FullName, "ثبت کننده", (double)r.MoalfeVal_Value);

                }
            }




            var z = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == pyy && p.tbPeymanContracts.Inactive != true && p.svdfunc_BastehID == Bas && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Year == Year

            && s.MoalfeVal_FKUser == usr && s.MoalfeVal_Month == Month && s.MoalfeVal_FKMoalafeDastmozdi == Moalfe)).ToList();

            return View(f);

        }


        public string cHEAK(int MoalfeVal_Year, int MoalfeVal_Month, int Basteh, int baste_peyman1)
        {
            int userid = 0;
            var User = new tbUsers();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];

            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }

                if (User != null)
                {
                    userid = User.usr_ID;
                }
            }

            List<tbSavedFunctions> time1 = new List<tbSavedFunctions>();
            var ex = db.tbSavedFunctions
                .Where(p => p.svdfunc_pymnID == baste_peyman1 && p.tbPeymanContracts.Inactive != true && p.svdfunc_BastehID == Basteh && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == MoalfeVal_Month && s.MoalfeVal_Year == MoalfeVal_Year))
                .ToList();

            if (ex.Count == 0)
            {

                return "False2";





                // If the loop is not entered, return a default value

            }
            else
            {
                return "true";
            }
        }
    }
}