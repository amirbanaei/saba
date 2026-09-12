using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Ussers;
using System;
using System.Buffers.Text;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;


namespace SaabWebProject.Areas.Users.Controllers
{
    public class LoginController : Controller
    {
        #region متغیر ها

        
        SaabEntities db = new SaabEntities(); // should be deleted
        tbUsersRepository rep_user = new tbUsersRepository();

        #endregion

        #region صفحات

        /// <summary>
        /// صفحه ورود برای اولین بار توسط شماره موبایل
        /// </summary>
        /// <returns></returns>
        public ActionResult FirstLogin()
        {
            return View("~/Areas/Users/Views/Login/FirstLogin.cshtml");
        }

        public ActionResult newpassword()
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
                        return View("~/Areas/Users/Views/Login/SetNewPassword.cshtml", User.usr_PhoneNumber);

                    }
                    else
                    {
                        return View("~/Areas/Users/Views/Login/SetNewPassword.cshtml", 09000000000);

                    }
                }
            }
            else
            {
                return View("~/Areas/Users/Views/Login/SetNewPassword.cshtml", 09000000000);

            }
        }

        /// <summary>
        /// صفحه ایجاد رمز کاربری جدید
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult SetNewPassword(long phoneNumber)
        {
            if (phoneNumber != 0)
            {
                var User = rep_user.CheckExsitsUser(phoneNumber);
                if (User.usr_SmsCodeRecievedTime.Value.AddMinutes(10) > DateTime.Now)
                {
                    return View("~/Areas/Users/Views/Login/SetNewPassword.cshtml", (long)User.usr_PhoneNumber);
                }
                return View("~/Areas/Users/Views/Login/FirstLogin_PhoneNumber.cshtml", (long)User.usr_PhoneNumber);

            }
            else
            {
                return HttpNotFound("404");
            }
        }

        /// <summary>
        /// صفحه ورود با شما موبایل و پسورد
        /// </summary>
        /// <returns></returns>
        /// 
        [Route("Login")]
        public ActionResult Login(string message = null)
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if(User != null) { 
                    
                    var find =db.dbAsnadsalvorodi.Where(p=>p.FK_usr==User.usr_ID).FirstOrDefault();
                        if(find != null)
                        {
                            find.login = false;
                            db.SaveChanges();




                        }
                    }
                }
            }
            //else
            //{
            //    var fi = db.dbrelatinperson.ToList();

            //}
            ViewBag.Message = message;
            return View("~/Areas/Users/Views/Login/Login.cshtml");
        }

        /// <summary>
        /// صفحه وارد کردن کد تایید
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult FirstLogin_PhoneNumber(long PhoneNumber)
        {
            if (PhoneNumber != 0)
            {
                var User = rep_user.CheckExsitsUser(PhoneNumber);
                if(User.usr_SmsCodeRecievedTime != null)
                {
                    if (User.usr_SmsCodeRecievedTime.Value.AddMinutes(10) > DateTime.Now)
                    {
                        return View("~/Areas/Users/Views/Login/FirstLogin_PhoneNumber.cshtml", PhoneNumber); 
                    }
                }
                SendCode(PhoneNumber);
                return View("~/Areas/Users/Views/Login/FirstLogin_PhoneNumber.cshtml", PhoneNumber); 
            }
            return View("~/Areas/Users/Views/Login/Login.cshtml"); // should be chnage view to add sms code

        }

        #endregion

        #region توابع

        /// <summary>
        /// تابع ارسال کد تایید
        /// </summary>
        /// <param name="phonenumber"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SendCode(long phonenumber)
        {
            Random rand = new Random();
            var smsCode = rand.Next(1000, 9999);
            var User = rep_user.Update_SmsCode(phonenumber, smsCode);
            if (User == "True")
            {
                //تابع sms پنل
                WSDLSample.PanelSMS.smsserver client = new WSDLSample.PanelSMS.smsserver();
                var username = "09000000000";
                var password = "REPLACE_ME";
                var fromNum = "3000505";
                string[] toNum = { "0" + phonenumber.ToString() };

                var patternCode = "e0zetx17wu3k4e9";
                string fullname = "";
                using(var db = new SaabEntities())
                {
                    var old = db.tbUsers.Where(p => p.usr_PhoneNumber == phonenumber).FirstOrDefault();
                    fullname = old.FullName;

                }

                var data = new WSDLSample.PanelSMS.input_data_type[] { new WSDLSample.PanelSMS.input_data_type() { key = "AcceptNumber", value = smsCode.ToString() }, new WSDLSample.PanelSMS.input_data_type() { key = "FullName", value = fullname } };

                var response = client.sendPatternSms(fromNum, toNum, username, password, patternCode, data);
            }

            return Content(User);
        }

        public string FirstLogin_SmsCode(long PhoneNumber, int smsCode)
        {
            return rep_user.CheckSmsCodeForLogin(PhoneNumber, smsCode);
        }

        /// <summary>
        /// تابع تماس صوتی با کاربر برای خواندن کد sms
        /// </summary>
        /// <returns></returns>
        public ActionResult FirstLogin_Call()
        {
            return null;
        }

        /// <summary>
        /// تابع تغییر رمز کاربری
        /// </summary>
        /// <param name="tbuser"></param> 
        /// <returns></returns>
        public ActionResult EventSetNewPassword(string password = "REPLACE_ME", long PhoneNumber=0)
        {
            var User = rep_user.CheckExsitsUser(PhoneNumber);
            if (User != null)
            {
                var HashPass = FormsAuthentication.HashPasswordForStoringInConfigFile(password, "MD5");
                User.usr_Password = HashPass;
                User.usr_first_payment = false;
                var bit = rep_user.Update(User);
                return Content(bit);
            }
            else
            {
                return Content("false");
            }
        }
  

        /// <summary>
        ///  تابع ورود کاربر از دفعه دوم به بعد
        /// </summary>
        /// <param name="login">شماره همراه و پسورد را میفرستد</param>
        /// <returns></returns>
        [HttpPost]
        [Route("Login")]
        public ActionResult Login(tbUsers login)
         {
            string hashPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(login.usr_Password, "MD5");
            //string kkii = FormsAuthentication.HashPasswordForStoringInConfigFile("Emad@1399", "MD5");
            var user = db.tbUsers.SingleOrDefault(u => u.usr_NationalCode == login.usr_NationalCode);
            if (user == null)
            {
                //string username = "saabadmin";
                //string password = "REPLACE_ME";
                //string password1 = "82E98110BC044FD153932F211138F2C9";
                //string userId = "admin"; 
                //string CodeMeli = "1234565432";
                //string position = "admin";
                ////string phone1 = "09000000000";

                //if (login.usr_NationalCode.ToString() == CodeMeli && (hashPassword == password || hashPassword == password1))
                //{

                //    UserStuf.SetOnlineUser(new tbUsers()
                //    {
                //        usr_NationalCode = "1234565432",
                //        usr_ID = 0,
                //        usr_Name = "saabadmin",
                //        usr_PhoneNumber = 09000000000
                //    });

                //    System.Text.StringBuilder builder = new System.Text.StringBuilder();
                //    var x = Utility.Base64.Base64Encode("userid");
                //    var y = Utility.Base64.Base64Encode("CodeMeli");
                //    var h = Utility.Base64.Base64Encode("position");

                //    var userid = Utility.Base64.Base64Encode(userId);
                //    var phone = Utility.Base64.Base64Encode(CodeMeli);
                //    position = Utility.Base64.Base64Encode(position);

                //    if (Request.Cookies[x] != null)
                //    {
                //        Request.Cookies[x].Value = userid;
                //        Request.Cookies[x].Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(Request.Cookies[x]);
                //    }
                //    else
                //    {
                //        HttpCookie httpCookie = new HttpCookie(x);
                //        httpCookie.Value = userid;
                //        httpCookie.Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(httpCookie);
                //    }

                //    if (Request.Cookies[y] != null)
                //    {
                //        Request.Cookies[y].Value = phone;
                //        Request.Cookies[y].Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(Request.Cookies[y]);
                //    }
                //    else
                //    {
                //        HttpCookie httpCookie1 = new HttpCookie(y);
                //        httpCookie1.Value = phone;
                //        httpCookie1.Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(httpCookie1);
                //    }

                //    if (Request.Cookies[h] != null)
                //    {
                //        Request.Cookies[h].Value = position;
                //        Request.Cookies[h].Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(Request.Cookies[h]);
                //    }
                //    else
                //    {
                //        HttpCookie httpCookie1 = new HttpCookie(h);
                //        httpCookie1.Value = position;
                //        httpCookie1.Expires = DateTime.Now.AddDays(7);
                //        Response.Cookies.Add(httpCookie1);
                //    }
                //    FormsAuthentication.SetAuthCookie(CodeMeli, login.usr_RememberMe);
                //    return Content("true-" + "/Reporter/Dashboard/Dashboard2");
                //}
            }

            if (user != null && user.usr_IsActive == true)
            {
                if (user.usr_Password != hashPassword)
                {

                  
                        dblog dblog = new dblog();
                        dblog.FK_usr = user.usr_ID;
                        dblog.personalID = user.usr_Personal_ID;
                        dblog.FULLname = user.FullName;
                        dblog.time = DateTime.Now;
                        dblog.Action = "users/login/login";
                        dblog.sharh = "رمز  اشتباه وارد شده بود ";
                    PersianCalendar pc = new PersianCalendar();
                    DateTime now = DateTime.Now;

                    string persianDateTime =
                        $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                        $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                    dblog.PersianDate = persianDateTime;
                    dblog.MACADDRESS = "";
                        db.dblog.Add(dblog);

                      
                        db.SaveChanges();
                        //return Content("true-" + "/Reporter/Dashboard/MaMoorin_Dashboard");
                        //return Content("true-/Users/Login/SetNewPassword?phoneNumber=" + user.usr_PhoneNumber);
                    
                    return Content("false-" + "لطفا اطلاعات صحیح خودرا وارد نمایید");

                }
                else if(user.usr_first_payment != false)
                {
                    dblog dblog = new dblog();
                    dblog.FK_usr = user.usr_ID;
                    dblog.personalID = user.usr_Personal_ID;
                    dblog.FULLname = user.FullName;
                    dblog.time = DateTime.Now;
                    dblog.Action = "users/login/login";
                    dblog.sharh = "دفعه اوله وارد سامانه شده ";
                    PersianCalendar pc = new PersianCalendar();
                    DateTime now = DateTime.Now;

                    string persianDateTime =
                        $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                        $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                    dblog.PersianDate = persianDateTime;
                    dblog.MACADDRESS = "";
                    db.dblog.Add(dblog);


                    db.SaveChanges();
                    return Content("true-/Users/Login/FirstLogin_PhoneNumber?phoneNumber=" + user.usr_PhoneNumber);
                }

                var HassAcc = user.tbLink_UserAndAction.Where(p => p.HasAccess == true).FirstOrDefault();
                var UrlReturn = "";
                if (HassAcc != null)
                {
                    if (HassAcc.tbActions != null)
                    {
                        UrlReturn = "true-" + HassAcc.tbActions.Area + "/" + HassAcc.tbActions.Controller + "/" + HassAcc.tbActions.Action;
                    }
                }

                foreach (var item in user.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
                {
                    HassAcc = db.tbLink_UserAndAction.Where(p => p.FK_Position_ID == item.FK_up_ID && p.FK_User_ID == null && p.HasAccess == true).FirstOrDefault();

                    if (HassAcc != null)
                    {
                        UrlReturn = "true-" + HassAcc.tbActions.Area + "/" + HassAcc.tbActions.Controller + "/" + HassAcc.tbActions.Action;
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(UrlReturn))
                {
                    var userid = Utility.Base64.Base64Encode(user.usr_ID.ToString());
                    var CodeMeli = Utility.Base64.Base64Encode(user.usr_NationalCode.ToString());
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();
                    System.Text.StringBuilder builderID = new System.Text.StringBuilder();

                    foreach (var item in user.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
                    {
                        builder.Append(":" + item.tbBusinessSide.up_name);
                        builderID.Append(item.FK_up_ID);
                    }

                    string position = "";
                    string positionID = "";

                    if (builder.Length > 0)
                    {
                        position = Utility.Base64.Base64Encode(builder.ToString());
                        positionID = Utility.Base64.Base64Encode(builderID.ToString());
                    }

                    var x = Utility.Base64.Base64Encode("userid");
                    var y = Utility.Base64.Base64Encode("CodeMeli");
                    var h = Utility.Base64.Base64Encode("position");
                    var PosID = Utility.Base64.Base64Encode("positionID");


                    if (Request.Cookies[x] != null)
                    {
                        Request.Cookies[x].Value = userid;
                        Request.Cookies[x].Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(Request.Cookies[x]);
                    }
                    else
                    {
                        HttpCookie httpCookie = new HttpCookie(x);
                        httpCookie.Value = userid;
                        httpCookie.Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(httpCookie);
                    }

                    if (Request.Cookies[y] != null)
                    {
                        Request.Cookies[y].Value = CodeMeli;
                        Request.Cookies[y].Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(Request.Cookies[y]);
                    }
                    else
                    {
                        HttpCookie httpCookie1 = new HttpCookie(y);
                        httpCookie1.Value = CodeMeli;
                        httpCookie1.Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(httpCookie1);
                    }

                    if (Request.Cookies[h] != null)
                    {
                        Request.Cookies[h].Value = position;
                        Request.Cookies[h].Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(Request.Cookies[h]);
                    }
                    else
                    {
                        HttpCookie httpCookie1 = new HttpCookie(h);
                        httpCookie1.Value = position;
                        httpCookie1.Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(httpCookie1);
                    }
                    if (Request.Cookies[PosID] != null)
                    {
                        Request.Cookies[PosID].Value = positionID;
                        Request.Cookies[PosID].Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(Request.Cookies[PosID]);
                    }
                    else
                    {
                        HttpCookie httpCookie1 = new HttpCookie(PosID);
                        httpCookie1.Value = positionID;
                        httpCookie1.Expires = DateTime.Now.AddDays(7);
                        Response.Cookies.Add(httpCookie1);
                    }

                    FormsAuthentication.SetAuthCookie(user.usr_NationalCode.ToString(), login.usr_RememberMe);
                    //return Content(UrlReturn);
                    dblog dblog = new dblog();
                    dblog.FK_usr = user.usr_ID;
                    dblog.personalID = user.usr_Personal_ID;
                    dblog.FULLname = user.FullName;
                    dblog.time = DateTime.Now;
                    dblog.Action = "users/login/login";
                    dblog.sharh = "با موفقیت وارد سامانه شد";
                    dblog.MACADDRESS = "";
                    PersianCalendar pc = new PersianCalendar();
                    DateTime now = DateTime.Now;

                    string persianDateTime =
                        $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                        $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                    dblog.PersianDate = persianDateTime;
                    db.dblog.Add(dblog);
                    db.SaveChanges();
                    return Content("true-" + "/Reporter/Dashboard/_banner");
                    //return Content("true-" + "/Setting/Orders/vieweee1");





                }

                return Content("false-" + "کاربر گرامی شما فاقد سطح دسترسی می باشید");
            }
            else if (user == null)
            {
                dblog dblog = new dblog();
                dblog.FK_usr = null;
                dblog.personalID = null;
                dblog.FULLname = login.usr_NationalCode;
                dblog.time = DateTime.Now;
                dblog.Action = "users/login/login";
                dblog.sharh = "همچین یوزری در سامانه نداریم ";
                dblog.MACADDRESS = "";
                PersianCalendar pc = new PersianCalendar();
                DateTime now = DateTime.Now;

                string persianDateTime =
                    $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                    $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                dblog.PersianDate = persianDateTime;

                db.dblog.Add(dblog);
                db.SaveChanges();
                return Content("false-" + "لطفا اطلاعات صحیح خودرا وارد نمایید");
            }
            else if (user.usr_IsActive == false || user.usr_IsActive == null)
            {
                dblog dblog = new dblog();
                dblog.FK_usr = user.usr_ID;
                dblog.personalID = user.usr_Personal_ID;
                dblog.FULLname = user.FullName;
                dblog.time = DateTime.Now;
                dblog.Action = "users/login/login";
                dblog.sharh = "اعتبار کاربر تمام شده  ";
                PersianCalendar pc = new PersianCalendar();
                DateTime now = DateTime.Now;

                string persianDateTime =
                    $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                    $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                dblog.PersianDate = persianDateTime;
                dblog.MACADDRESS = "";
                db.dblog.Add(dblog);
                db.SaveChanges();

                return Content("false-" + "اعتبار کاربری شما به پایان رسیده است");
            }
            else
            {
                dblog dblog = new dblog();
                dblog.FK_usr = user.usr_ID;
                dblog.personalID = user.usr_Personal_ID;
                dblog.FULLname = user.FullName;
                dblog.time = DateTime.Now;
                dblog.Action = "users/login/login";
                dblog.sharh = "خطا در اعتبار سنجی ";
                PersianCalendar pc = new PersianCalendar();
                DateTime now = DateTime.Now;

                string persianDateTime =
                    $"{pc.GetDayOfMonth(now)}/{pc.GetMonth(now)}/{pc.GetYear(now)}" +
                    $"_{pc.GetHour(now):00}:{pc.GetMinute(now):00}:{pc.GetSecond(now):00}";
                dblog.PersianDate = persianDateTime;
                dblog.MACADDRESS = "";
                db.dblog.Add(dblog);
                db.SaveChanges();
                return Content("false-" + "خطا در اعتبار سنجی");
            }
        }
        //public ActionResult Login2()
        //{
        //    return View();
        //}
        #region
        //public ActionResult Login(tbUsers login)
        //{
        //    string hashPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(login.usr_Password, "MD5");

        //    var user = db.tbUsers.SingleOrDefault(u => u.usr_NationalCode == login.usr_NationalCode);
        //    if (user == null)
        //    {

        //        //string username = "saabadmin";
        //        //string password = "REPLACE_ME";
        //        //string password1 = "82E98110BC044FD153932F211138F2C9";
        //        //string userId = "admin";
        //        //string CodeMeli = "1234565432";
        //        //string position = "admin";
        //        ////string phone1 = "09000000000";

        //        //if (login.usr_NationalCode.ToString() == CodeMeli && (hashPassword == password || hashPassword == password1))
        //        //{

        //        //    UserStuf.SetOnlineUser(new tbUsers()
        //        //    {
        //        //        usr_NationalCode = "1234565432",
        //        //        usr_ID = 0,
        //        //        usr_Name = "saabadmin",
        //        //        usr_PhoneNumber = 09000000000
        //        //    });

        //        //    System.Text.StringBuilder builder = new System.Text.StringBuilder();
        //        //    var x = Utility.Base64.Base64Encode("userid");
        //        //    var y = Utility.Base64.Base64Encode("CodeMeli");
        //        //    var h = Utility.Base64.Base64Encode("position");

        //        //    var userid = Utility.Base64.Base64Encode(userId);
        //        //    var phone = Utility.Base64.Base64Encode(CodeMeli);
        //        //    position = Utility.Base64.Base64Encode(position);

        //        //    if (Request.Cookies[x] != null)
        //        //    {
        //        //        Request.Cookies[x].Value = userid;
        //        //        Request.Cookies[x].Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(Request.Cookies[x]);
        //        //    }
        //        //    else
        //        //    {
        //        //        HttpCookie httpCookie = new HttpCookie(x);
        //        //        httpCookie.Value = userid;
        //        //        httpCookie.Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(httpCookie);
        //        //    }

        //        //    if (Request.Cookies[y] != null)
        //        //    {
        //        //        Request.Cookies[y].Value = phone;
        //        //        Request.Cookies[y].Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(Request.Cookies[y]);
        //        //    }
        //        //    else
        //        //    {
        //        //        HttpCookie httpCookie1 = new HttpCookie(y);
        //        //        httpCookie1.Value = phone;
        //        //        httpCookie1.Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(httpCookie1);
        //        //    }

        //        //    if (Request.Cookies[h] != null)
        //        //    {
        //        //        Request.Cookies[h].Value = position;
        //        //        Request.Cookies[h].Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(Request.Cookies[h]);
        //        //    }
        //        //    else
        //        //    {
        //        //        HttpCookie httpCookie1 = new HttpCookie(h);
        //        //        httpCookie1.Value = position;
        //        //        httpCookie1.Expires = DateTime.Now.AddDays(7);
        //        //        Response.Cookies.Add(httpCookie1);
        //        //    }
        //        //    FormsAuthentication.SetAuthCookie(CodeMeli, login.usr_RememberMe);
        //        //    return Content("true-" + "/Reporter/Dashboard/Dashboard2");
        //        //}
        //    }

        //    if (user != null && user.usr_IsActive == true)
        //    {
        //        if (user.usr_Password != hashPassword)
        //        {
        //            UserStuf.SetOnlineUser(user);
        //            if (user.usr_first_payment == null || (bool)user.usr_first_payment)
        //            {
        //                Random rand = new Random();
        //                var smsCode = rand.Next(1000, 9999);
        //                user.usr_SmsCodeRecievedTime = DateTime.Now;
        //                user.usr_SmsCode = smsCode;
        //                db.SaveChanges();
        //                //return Content("true-" + "/Reporter/Dashboard/MaMoorin_Dashboard");
        //                return Content("true-/Users/Login/SetNewPassword?phoneNumber=" + user.usr_PhoneNumber);
        //            }
        //        }

        //        var HassAcc = user.tbLink_UserAndAction.Where(p => p.HasAccess == true).FirstOrDefault();
        //        var UrlReturn = "";
        //        if (HassAcc != null)
        //        {
        //            if (HassAcc.tbActions != null)
        //            {
        //                UrlReturn = "true-" + HassAcc.tbActions.Area + "/" + HassAcc.tbActions.Controller + "/" + HassAcc.tbActions.Action;
        //            }
        //        }

        //        foreach (var item in user.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
        //        {
        //            HassAcc = db.tbLink_UserAndAction.Where(p => p.FK_Position_ID == item.FK_up_ID && p.FK_User_ID == null && p.HasAccess == true).FirstOrDefault();

        //            if (HassAcc != null)
        //            {
        //                UrlReturn = "true-" + HassAcc.tbActions.Area + "/" + HassAcc.tbActions.Controller + "/" + HassAcc.tbActions.Action;
        //                break;
        //            }
        //        }

        //        if (!string.IsNullOrEmpty(UrlReturn))
        //        {
        //            var userid = Utility.Base64.Base64Encode(user.usr_ID.ToString());
        //            var CodeMeli = Utility.Base64.Base64Encode(user.usr_NationalCode.ToString());
        //            System.Text.StringBuilder builder = new System.Text.StringBuilder();

        //            foreach (var item in user.tbUser_link_BusinessSide.Where(p => p.Status == true).ToList())
        //            {
        //                builder.Append(":" + item.tbBusinessSide.up_name);
        //            }

        //            string position = "";

        //            if (builder.Length > 0)
        //            {
        //                position = Utility.Base64.Base64Encode(builder.ToString());
        //            }

        //            var x = Utility.Base64.Base64Encode("userid");
        //            var y = Utility.Base64.Base64Encode("CodeMeli");
        //            var h = Utility.Base64.Base64Encode("position");


        //            if (Request.Cookies[x] != null)
        //            {
        //                Request.Cookies[x].Value = userid;
        //                Request.Cookies[x].Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(Request.Cookies[x]);
        //            }
        //            else
        //            {
        //                HttpCookie httpCookie = new HttpCookie(x);
        //                httpCookie.Value = userid;
        //                httpCookie.Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(httpCookie);
        //            }

        //            if (Request.Cookies[y] != null)
        //            {
        //                Request.Cookies[y].Value = CodeMeli;
        //                Request.Cookies[y].Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(Request.Cookies[y]);
        //            }
        //            else
        //            {
        //                HttpCookie httpCookie1 = new HttpCookie(y);
        //                httpCookie1.Value = CodeMeli;
        //                httpCookie1.Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(httpCookie1);
        //            }

        //            if (Request.Cookies[h] != null)
        //            {
        //                Request.Cookies[h].Value = position;
        //                Request.Cookies[h].Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(Request.Cookies[h]);
        //            }
        //            else
        //            {
        //                HttpCookie httpCookie1 = new HttpCookie(h);
        //                httpCookie1.Value = position;
        //                httpCookie1.Expires = DateTime.Now.AddDays(7);
        //                Response.Cookies.Add(httpCookie1);
        //            }

        //            FormsAuthentication.SetAuthCookie(user.usr_NationalCode.ToString(), login.usr_RememberMe);
        //            return Content(UrlReturn);
        //           // return Content("true-" + "/Reporter/Dashboard/MaMoorin_Dashboard");

        //        }

        //        return Content("false-" + "کاربر گرامی شما فاقد سطح دسترسی می باشید");
        //    }
        //    else if (user == null)
        //    {
        //        return Content("false-" + "لطفا اطلاعات صحیح خودرا وارد نمایید");
        //    }
        //    else if (user.usr_IsActive == false || user.usr_IsActive == null)
        //    {
        //        return Content("false-" + "اعتبار کاربری شما به پایان رسیده است");
        //    }
        //    else
        //    {
        //        return Content("false-" + "خطا در اعتبار سنجی");
        //    }
        //}
        #endregion
        public ActionResult LogOff()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    //var fin = db.dbrelatinperson.Where(p => p.tbperson.FK_usr == User.usr_ID).ToList();
                    if (User != null)
                    {

                        var find = db.dbAsnadsalvorodi.Where(p => p.FK_usr == User.usr_ID).FirstOrDefault();
                        if (find != null)
                        {
                            find.login = false;
                            db.SaveChanges();
                        }
                    }
                }
            }
            FormsAuthentication.SignOut();
            return Redirect("/");
        }

        #endregion
    }
}