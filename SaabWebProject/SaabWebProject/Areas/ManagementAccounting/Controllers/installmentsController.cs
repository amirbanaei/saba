using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.ManagementAccounting;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Areas.ManagementAccounting.Controllers
{
    public class installmentsController : Controller
    {
        SaabEntities db=new SaabEntities();
        installmentsRepositories rep_def_rsn = new installmentsRepositories();

        // GET: ManagementAccounting/installments
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult _Manual()
        {
            return View("~/Areas/ManagementAccounting/Views/installments/_Manaual.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult _Automatic()
        {
            return View("~/Areas/ManagementAccounting/Views/installments/_Automatic.cshtml");
        }
        [AuthorizeAAA]
        public ActionResult typeOfCalculate()
        {
            return PartialView("~/Areas/ManagementAccounting/Views/installments/_PartialViewTypeOfCalculate.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult InstallmentManually(Int64 totalAmountInput)
        {
            return PartialView("~/Areas/ManagementAccounting/Views/installments/_InstallmentManually.cshtml", totalAmountInput);
        }
        public ActionResult InstallmentManually2(Int64 totalAmountInput)
        {
            return PartialView("~/Areas/ManagementAccounting/Views/installments/InstallmentManually2.cshtml", totalAmountInput);
        }
        public ActionResult installments_List()
        {
            return View("~/Areas/ManagementAccounting/Views/installments/installments_List.cshtml", rep_def_rsn.Update3());
        }

        public ActionResult installments_Edit(int id)
        {
            return View("~/Areas/ManagementAccounting/Views/installments/installments_Edit.cshtml", rep_def_rsn.Find(id));
        }

        public ActionResult installments_details(int id)
        {
            return View("~/Areas/ManagementAccounting/Views/installments/installments_details.cshtml", rep_def_rsn.Find(id));
        }

        public ActionResult moalfe_fk()
        {
            return View(db.tbContractMoalefeDastmozdi.ToList());
        }
        public ActionResult moalfe_fk2()
        {
            return View(db.tbContractMoalefeDastmozdi.ToList());
        }





        public string UploadAttachmentFile(int Peyman_ID = 0, int Moalfe_fk = 0, int cheackbox = 0,
  int User_ID = 0, int Total_Amount = 0, bool Calculation_type = true, int Debtore = 0, int Creditor = 0, string Title = "", int numberOfInstallment = 0,
  IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null,int Month=0,int Year=0, bool? actions = null)
        {


            tbinstallments obj = new tbinstallments();

            if (Peyman_ID == 0)
            {
                obj.Peyman_ID = null;
            }
            else
            {
                obj.Peyman_ID = Peyman_ID;
            }

            if (User_ID == 0)
            {
                obj.User_ID = null;
            }
            else
            {
                obj.User_ID = User_ID;
            }
            obj.Year = Year;
            obj.Month=Month;
            obj.Total_Amount = Total_Amount;
            obj.Calculation_type = Calculation_type;
            obj.numberOfInstallment = numberOfInstallment;
            obj.Title = Title;
            obj.actions= actions;
            obj.Fk_molfe = Moalfe_fk;
            obj.For_Fish = false;
            if (cheackbox == 3)
            {
                obj.Action_sorat = true;

            }
            else if (cheackbox == 2)
            {
                obj.Action_Fish = true;

            }
            else if (cheackbox == 1)
            {
                obj.Action_traz = true;

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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/installmentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;


                        return rep_def_rsn.Create(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return rep_def_rsn.Create(obj).ToString();
            }
        }



        public string Edit_UploadAttachmentFile(int installment_ID = 0, int Year = 0, int Month = 0, int Moalfe_fk = 0, int cheackbox = 0,
 int Total_Amount = 0, bool Calculation_type = true, int Debtore = 0, int Creditor = 0, string Title = "", int numberOfInstallment = 0,
  IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {


            tbinstallments obj = new tbinstallments();
            obj.installment_ID = installment_ID;
            obj.Total_Amount = Total_Amount;
            obj.Calculation_type = Calculation_type;
            obj.numberOfInstallment = numberOfInstallment;
            obj.Title = Title;
            obj.Month = Month;
            obj.Year=Year;
            obj.For_Fish = false;
            obj.Fk_molfe = Moalfe_fk;

            if (cheackbox == 3)
            {
                obj.Action_sorat = true;

            }
            else if (cheackbox == 2)
            {
                obj.Action_Fish = true;

            }
            else if (cheackbox == 1)
            {
                obj.Action_traz = true;

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
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/installmentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;


                        return rep_def_rsn.Update(obj).ToString();

                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return rep_def_rsn.Update(obj).ToString();
            }
        }





        [AuthorizeAAA]
        [HttpPost]
        public string installments_Edit(int Peyman_ID = 0,int installment_ID=0,
  int User_ID = 0, int Total_Amount = 0, bool Calculation_type = true, int Debtore = 0, int Creditor = 0, string Title = "", int numberOfInstallment = 0,
  IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {

            tbinstallments obj = new tbinstallments();
            obj.installment_ID = installment_ID;

            if (Peyman_ID == 0)
            {
                obj.Peyman_ID = null;
            }
            else
            {
                obj.Peyman_ID = Peyman_ID;
            }

            if (User_ID == 0)
            {
                obj.User_ID = null;
            }
            else
            {
                obj.User_ID = User_ID;
            }

            obj.Total_Amount = Total_Amount;
            obj.Calculation_type = Calculation_type;
            obj.numberOfInstallment = numberOfInstallment;
            obj.Title = Title;


            
            if (files != null)
            {

                foreach (var file in files)
                {

                    if (file.ContentLength > 0)
                    {
                        var segment = file.FileName.Split('.');
                        string file_type = segment[segment.Length - 1];
                        var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                        file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/ManagementAccounting/Contents/installmentAttachmentsFile/" + filename));
                        obj.File_SystemName = filename;
                        obj.File_RealName = file.FileName;





                        return rep_def_rsn.Update(obj).ToString();
                    }
                    else
                    {

                    }
                }
                return "True";
            }
            else
            {

                return rep_def_rsn.Update(obj).ToString();
            }
        }
        public string installmentsDelete(int ID)
        {
            return rep_def_rsn.Disable(ID).ToString();
        }

        public string tbinstallments_Add(List<tb_Subset_of_installments> Filters)
        {
            return rep_def_rsn.Create(Filters);
        }
        public string tbinstallments_Edit(List<tb_Subset_of_installments> Filters)
        {
            return rep_def_rsn.Create(Filters);
        }

        //public string tbinstallments_Edit(tb_Subset_of_installments Filters)
        //{
        //    return rep_def_rsn.Update2(Filters);
        //}
    }
}