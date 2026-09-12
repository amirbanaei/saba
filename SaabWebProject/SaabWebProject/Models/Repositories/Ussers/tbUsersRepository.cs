using Newtonsoft.Json;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.ViewModels.Users;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using PagedList;

namespace SaabWebProject.Models.Repositories.Ussers
{
    public class tbUsersRepository : IInterFace<tbUsers>
    {
        SaabEntities db;

        public tbUsersRepository(SaabEntities Context)
        {
            db = Context;
        }

        public tbUsersRepository()
        {
            db = new SaabEntities();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Create(tbUsers user)
        {
            //chi neveshte !!!
            try
            {
                db.tbUsers.Add(user);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }
            catch
            {
                return false.ToString();
            }
        }

        public tbUsers Find(int id)
        {
            return db.tbUsers.Find(id);
        } 
        public tbUsers FindByNationalCode(string NationalCode)
        {
            return db.tbUsers.Where(p=>p.usr_NationalCode == NationalCode).FirstOrDefault();
        }

        public List<tbUsers> Update()
        {
            return db.tbUsers.ToList();
        }

        public List<tbUsers> ActiveList()
        {
            return db.tbUsers.Where(p => p.usr_IsActive == true).ToList();
        }
        public List<tbUsers> HoghooghiUser()
        {
            return db.tbUsers.Where(p => p.usr_IsUser != true).ToList();
        }
        public List<tbUsers> HaghighiUser()
        {
            return db.tbUsers.Where(p => p.usr_IsUser != false).ToList();
        }

        public string Update(tbUsers entity)
        {
            if (entity != null)
            {
                try
                {
                    db.Entry(entity).State = EntityState.Modified;
                    return Convert.ToBoolean(db.SaveChanges()).ToString();
                }
                catch
                {
                    return false.ToString();
                }
            }
            else
            {
                return true.ToString();
            }
        }

        /// <summary>
        /// بروز رسانی عدد sms شده
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <param name="smsCode"></param>
        /// <returns></returns>
        public string Update_SmsCode(long phoneNumber, int smsCode)
        {
            var old = db.tbUsers.Where(p => p.usr_PhoneNumber == phoneNumber).FirstOrDefault();
            if (old != null)
            {
                if (old.usr_IsActive == false || old.usr_IsActive == null)
                {
                    return "User_NotActive";
                }

                old.usr_SmsCode = smsCode;
                old.usr_SmsCodeRecievedTime = DateTime.Now;
                try
                {
                    return Convert.ToBoolean(db.SaveChanges()).ToString();
                }
                catch (Exception e)
                {
                    return false.ToString();
                }
            }
            else
            {
                return "User_NotFound";
            }
        }

        /// <summary>
        /// چمک کردن صحت کد sms 
        /// </summary>
        /// <param name="phoneNumber"></param>
        /// <param name="smsCode"></param>
        /// <returns></returns>
        public string CheckSmsCodeForLogin(long phoneNumber, int smsCode)
        {
            var find_user = db.tbUsers.Where(p => p.usr_PhoneNumber == phoneNumber).FirstOrDefault();
            if (find_user != null)
            {
                if (find_user.usr_SmsCodeRecievedTime.Value.AddMinutes(2) < DateTime.Now.AddMinutes(2))
                {
                    if (smsCode == find_user.usr_SmsCode)
                    {
                        return "true";
                    }
                    else
                    {
                        return "false_WrongCode";
                    }
                }
                else
                {
                    return "false_Expired";
                }
            }
            else
            {
                return "false_NotFoundUser";
            }
        }

        public List<tbUser_link_BusinessSide> ListUserAndJobsOfKarfarmayan()
        {
            return db.tbUser_link_BusinessSide.Where(p => p.tbUsers.usr_IsUser != true).ToList();
        }

        public List<tbUsers> ListUserWithoutKarfarmayan()
        {
            return db.tbUsers.Where(p => p.usr_IsUser == true).ToList();
        }

        public bool Disable(int ID)
        {
            try
            {
                var user = db.tbUsers.Find(ID);
                if ((bool)user.usr_IsActive)
                {
                    user.usr_IsActive = false;
                }
                else
                {
                    user.usr_IsActive = true;
                }

                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }

        /// <summary>mn
        /// تابع چک کردن وجود داشتن کاربر با شماره همراه
        /// </summary>
        /// <returns></returns>
        public tbUsers CheckExsitsUser(long PhoneNumber)
        {
            return db.tbUsers.Where(p => p.usr_PhoneNumber == PhoneNumber).FirstOrDefault();
        }

        public List<CityViewModel> GetAllCity(string Path)
        {
            var json = System.IO.File.ReadAllText(Path);
            var obj = JsonConvert.DeserializeObject<List<CityViewModel>>(json);
            List<CityViewModel> obj2 = new List<CityViewModel>();
            var counter = 1;
            foreach (var i in obj)
            {
                i.cityId = counter;


                obj2.Add(i);
                counter++;
            }


            return obj2;
        }

        public int? IfExitsParentUser(string ParentName)
        {
            try
            {
                if (ParentName != null)
                {
                    tbUsers parent;
                    var res = ParentName.Split(' ').ToList();
                    if (res.Count >= 2)
                    {
                        var name = res[0];
                        res.Remove(res[0]);
                        var join = string.Join(" ", res.ToArray());
                        parent = db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == join).FirstOrDefault();
                    }
                    else
                    {
                        parent = db.tbUsers.Where(p => p.usr_Family.Trim() == ParentName.Trim() || p.usr_Name.Trim() == ParentName.Trim()).FirstOrDefault();
                    }

                    if (parent != null)
                    {
                        return parent.usr_ID;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public bool CheckHasAccess(long PhoneNumber, int posid, string Action, string Controller)
        {
            var user = db.tbUsers.Where(p => p.usr_PhoneNumber == PhoneNumber).FirstOrDefault();

            var acc = user.tbLink_UserAndAction.Where(p => p.tbActions.Action == Action && p.tbActions.Controller == Controller).FirstOrDefault();
            if (acc != null)
            {
                return true;
            }
            else
            {
                foreach (var item in user.tbUser_link_BusinessSide.ToList())
                {
                    var acc2 = db.tbLink_UserAndAction.Where(p => p.FK_Position_ID == item.tbBusinessSide.up_Id && p.FK_User_ID == null && p.tbActions.Action == Action && p.tbActions.Controller == Controller).FirstOrDefault();
                    if (acc2 != null)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public List<tbActions> GetAllActionLevel()
        {
            return db.tbActions.Where(p => p.Parent_ID != null).ToList();
        }

        public List<tbActions> GetActionsForController(int id)
        {
            return db.tbActions.Where(p => p.Parent_ID == id).ToList();
        }

        public bool SetAccessLevel(List<string> Actions, int UserID)
        {
            var user = db.tbUsers.Find(UserID);
            foreach (var item in user.tbLink_UserAndAction.ToList())
            {
                db.tbLink_UserAndAction.Remove(item);
            }

            if (Actions != null)
            {
                foreach (var item in Actions.ToList())
                {
                    var it = item.Split('_');
                    if (it[0] == "child")
                    {
                        if (it[1] == "0")
                        {
                            foreach (var item1 in db.tbActions.ToList())
                            {
                                var link = new tbLink_UserAndAction();
                                link.UserAction_ID = UserID;
                                link.FK_Action_ID = item1.Action_ID;
                                link.FK_Position_ID = user.tbUser_link_BusinessSide.Where(p=>p.Status==true).FirstOrDefault().FK_up_ID;
                                link.HasAccess = true;
                                user.tbLink_UserAndAction.Add(link);
                            }

                            break;
                        }
                        else
                        {
                            var link = new tbLink_UserAndAction();
                            link.UserAction_ID = UserID;
                            link.FK_Action_ID = Convert.ToInt32(it[1]);
                            link.FK_Position_ID = user.tbUser_link_BusinessSide.Where(p=>p.Status==true).FirstOrDefault().FK_up_ID;
                            link.HasAccess = true;
                            user.tbLink_UserAndAction.Add(link);
                        }
                    }
                }
            }

            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbUsers> GetAllUsersInSamePosition(int positionId)
        {
            var temp = db.tbUser_link_BusinessSide.Where(p => p.FK_up_ID == positionId && p.tbUsers != null).Select(p => p.tbUsers).ToList();
            if(temp != null)
            {
                var dis = temp.Distinct();
                if(dis != null)
                {
                    return dis.ToList();
                }
                else
                {
                    return new List<tbUsers>();

                }

            }
            return new List<tbUsers>();

        }

    }
}