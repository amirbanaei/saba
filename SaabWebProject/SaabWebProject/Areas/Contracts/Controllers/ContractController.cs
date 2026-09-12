using System.Linq;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories;
using SaabWebProject.Utility;

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class ContractController : Controller
    {
        // GET: Contract
        #region تعریف متغیر ها
        tbUserContractsRepository rep_UserContracts;
        SaabEntities Context = new SaabEntities(); 
        #endregion

        #region سازنده ها
        public ContractController()
        {
            rep_UserContracts = new tbUserContractsRepository(Context);
        }
        #endregion

        #region صفحات
        [AuthorizeAAA]
        public ActionResult ManageContracts()
        {
            return View("~/Areas/Contracts/Views/Contract/ManageContracts.cshtml");
        }

        #endregion

        #region توابع





        #endregion
        [AuthorizeAAA]
        public ActionResult ShowContractEmzaShode()
        {
            try
            {
                using (SaabEntities db = new SaabEntities())
                {

                    int? personal = 0;
                    var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                    if (cookie_user != null)
                    {
                        var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                        var user = db.tbUsers.FirstOrDefault(p => p.usr_NationalCode == nationalcode);
                        personal = user.usr_Personal_ID;
                    }

                    byte[] filebyyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Content/SignQarardad/" + personal + ".pdf"));

                    return File(filebyyte, System.Net.Mime.MediaTypeNames.Application.Octet, personal + ".pdf");
                }

            }
            catch (System.Exception)
            {

                return Content("0");

            }


        }

    }
}