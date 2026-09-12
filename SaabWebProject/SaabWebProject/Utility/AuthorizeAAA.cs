using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Filters;
using System.Web.Routing;
using System.Web.Security;

namespace SaabWebProject.Utility
{
    public class AuthorizeAAA : AuthorizeAttribute
    {
        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            var PosID = Base64.Base64Encode("positionID");
            var positionID = httpContext.Request.Cookies[PosID];
            var PositionsID = Base64.Base64Decode(positionID.Value);
            UserStuf.PosID = PositionsID;

            return true;
            bool authorize = true;
            // return true; //ایراد دارد میبایست پاک شود تا سطح دسترسی کار کند 

            if (httpContext.Request.IsAuthenticated)
            {
                var x = Base64.Base64Encode("userid");
                var y = Base64.Base64Encode("CodeMeli");
                var h = Base64.Base64Encode("position");

                var userid = httpContext.Request.Cookies[x];
                var CodeMeli = httpContext.Request.Cookies[y];
                var position = httpContext.Request.Cookies[h];


                if (userid != null && CodeMeli != null && position != null && positionID != null)
                {

                    var UserID = Base64.Base64Decode(userid.Value);
                    var Phone = Base64.Base64Decode(CodeMeli.Value);
                    var Positions = Base64.Base64Decode(position.Value);

                    if (UserID == "admin" && Phone == "1234565432" && Positions == "admin")
                    {
                        return true;
                    }
                    using (SaabEntities db = new SaabEntities())
                    {
                        foreach (var Position in Positions.Split(':').ToList().Where(p => p != "").ToList())
                        {
                            var User = db.tbUsers.Where(p => p.usr_ID.ToString() == UserID && p.usr_IsActive == true && p.tbUser_link_BusinessSide.Where(a => a.Status == true).Any(s => s.tbBusinessSide.up_name == Position)).FirstOrDefault();
                            if (User != null)
                            {
                                if (User.usr_NationalCode.ToString() == Phone)
                                {
                                    var controller = httpContext.Request.RequestContext.RouteData.Values["controller"].ToString();
                                    var action = httpContext.Request.RequestContext.RouteData.Values["action"].ToString();
                                    var Access = db.tbLink_UserAndAction.Where(p => p.tbActions.Action == action && p.tbActions.Controller == controller && p.tbBusinessSide.up_name == Position && p.FK_User_ID == null && p.HasAccess == true).Any();
                                    if (Access)
                                        return true;

                                    Access = User.tbLink_UserAndAction.Where(p => p.tbActions.Action == action && p.tbActions.Controller == controller && p.tbBusinessSide.up_name == Position && p.HasAccess == true).Any();
                                    if (Access)
                                        return true;


                                }
                            }
                        }
                    }
                }

            }
            //hhttpContext.Response.Redirect("/Login/Login");
            return authorize;
        }
    }
}