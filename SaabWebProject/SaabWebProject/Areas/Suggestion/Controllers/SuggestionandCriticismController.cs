using DocumentFormat.OpenXml.Office2010.Excel;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Suggestion.Controllers
{
    public class SuggestionandCriticismController : Controller
    {
        // GET: Suggestion/SuggestionandCriticism

        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

        [AuthorizeAAA]
        public ActionResult _Sabt_Pasokh(int suggestion_ID)
        {
            int user_id = 0;
            #region به دست آوردن ایدی کاربر لاگین شده
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    user_id = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).Select(p => p.usr_ID).FirstOrDefault();
                }

            }
            #endregion
            using (var db = new SaabEntities())
            {
                var list = db.tbSuggestion.Where(p => p.ID == suggestion_ID || p.FKParentID == suggestion_ID).OrderBy(p => p.ID).ToList();
                var flag = false;
                foreach (var item in list)
                {
                    if (user_id != item.FK_UserID)
                    {
                        if (item.IsSeen == false)
                        {
                            flag = true;
                            item.IsSeen = true;
                        }
                    }
                }
                if (flag == true)
                {
                    db.SaveChanges();
                }
                return View("~/Areas/Suggestion/Views/SuggestionandCriticism/_Sabt_Pasokh.cshtml", list);

            }
        }





        public bool Save(IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, string Description = "", string Title = "", bool Type = false)
        {
            try
            {
                var obj = new tbSuggestion();
                if (files != null)
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Suggestion/Contents/" + filename));
                            obj.FileSystemName = filename;
                            obj.FileName = file.FileName;
                        }
                    }
                }
                obj.Description = Description;
                obj.Title = Title;
                obj.Datetime = DateTime.Now;
                obj.FKParentID = null;
                obj.IsSeen = false;
                obj.IsSuggest = Type;

                #region به دست آوردن ایدی کاربر لاگین شده
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    using (SaabEntities db = new SaabEntities())
                    {
                        obj.FK_UserID = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).Select(p => p.usr_ID).FirstOrDefault();
                    }

                }
                #endregion

                using (var db = new SaabEntities())
                {
                    db.tbSuggestion.Add(obj);
                    db.SaveChanges();
                }
                return true;
            }
            catch (Exception)
            {

                return false;
            }

        }
        public ActionResult List()
        {
            int id = 0;
            List<tbSuggestion> list = new List<tbSuggestion>();
            #region به دست آوردن ایدی کاربر لاگین شده
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    id = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).Select(p => p.usr_ID).FirstOrDefault();
                    var fin= db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == id).FirstOrDefault();

                    if (fin != null && fin.FK_up_ID==9)
                    {
                        list = db.tbSuggestion.Where(p => p.FKParentID == null).ToList();
                        var flag = false;
                        foreach (var item in list)
                        {
                            if (item.IsSeen == false)
                            {
                                flag = true;
                                item.IsSeen = true;
                            }
                        }
                        if (flag == true)
                            db.SaveChanges();
                    }
                    else
                    {
                        list = db.tbSuggestion.Where(p => p.FK_UserID == id && p.FKParentID == null).ToList();
                    }
                    foreach (var item in list)
                    {
                        item.HasNotSeen = db.tbSuggestion.Where(p => p.FKParentID == item.ID && p.FK_UserID != id && p.IsSeen == false).Count();
                    }
                }

            }

            return PartialView(list.OrderByDescending(p=>p.HasNotSeen).ThenByDescending(x => x.Datetime).ToList());
            #endregion
        }
        public bool Reply(IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, string Description = "", int parent_ID = 0)
        {
            try
            {
                var obj = new tbSuggestion();
                if (files != null)
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Suggestion/Contents/" + filename));
                            obj.FileSystemName = filename;
                            obj.FileName = file.FileName;
                        }
                    }
                }
                obj.Description = Description;
                obj.Datetime = DateTime.Now;
                obj.FKParentID = parent_ID;
                obj.IsSeen = false;

                #region به دست آوردن ایدی کاربر لاگین شده
                var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
                if (cookie_user != null)
                {
                    var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                    using (SaabEntities db = new SaabEntities())
                    {
                        obj.FK_UserID = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).Select(p => p.usr_ID).FirstOrDefault();
                    }

                }
                #endregion

                using (var db = new SaabEntities())
                {
                    db.tbSuggestion.Add(obj);
                    db.SaveChanges();
                }
                return true;
            }
            catch (Exception e)
            {

                return false;
            }
        }
        public int Count_suggestion1()
        {
            int user_ID = 0;
            #region به دست آوردن ایدی کاربر لاگین شده
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    user_ID = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).Select(p => p.usr_ID).FirstOrDefault();
                }
            }
            #endregion
            using (var db = new SaabEntities())
            {
                var fin = db.tbUser_link_BusinessSide.Where(p => p.FK_u_ID == user_ID).FirstOrDefault();
                if (fin != null && fin.FK_up_ID == 9)
                {
                    return db.tbSuggestion.Where(p => p.FK_UserID != user_ID && p.IsSeen == false).Count();

                }
                //if (user_ID == 1710)
                //{
                //}
                else
                {
                    var a = db.tbSuggestion.Where(p => p.FK_UserID == user_ID && p.FKParentID == null).ToList();
                    int counter = 0;
                    foreach (var item in a)
                    {
                        counter += db.tbSuggestion.Where(p => p.FKParentID == item.ID && p.IsSeen == false).Count();
                    }
                    return counter;
                }
            }
        }
        public int Count_suggestion()
        {
            int user_ID = 0;

            var cookie = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie == null) return 0;

            var nationalcode = Utility.Base64.Base64Decode(cookie.Value);

            using (var db = new SaabEntities())
            {
                user_ID = db.tbUsers
                    .Where(p => p.usr_NationalCode == nationalcode)
                    .Select(p => p.usr_ID)
                    .FirstOrDefault();

                var fin = db.tbUser_link_BusinessSide
                    .Where(p => p.FK_u_ID == user_ID)
                    .Select(p => p.FK_up_ID)
                    .FirstOrDefault();

                if (fin == 9)
                {
                    return db.tbSuggestion.Count(p => p.FK_UserID != user_ID && p.IsSeen != true);
                }

                // همه Parent ها
                var parentIds = db.tbSuggestion
                    .Where(p => p.FK_UserID == user_ID && p.FKParentID == null)
                    .Select(p => p.ID);

                // شمارش جواب‌های دیده‌نشده
                return db.tbSuggestion.Count(p =>
                    parentIds.Contains(p.FKParentID.Value) && p.IsSeen != true);
            }
        }

    }
}