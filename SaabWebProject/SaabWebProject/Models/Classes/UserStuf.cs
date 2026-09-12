using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Models.Classes
{
    public static class UserStuf
    {
        public static tbUsers User;
        public static string PosID;
        public static void SetOnlineUser(tbUsers _user)
        {
            User = _user; 
        }
        public static tbUsers GetOlineUser()
        {
            return User;
        }
    }
}