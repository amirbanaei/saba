using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Extensions.Helper
{
    public static class extentions
    {
        public static string ToLowerFirst(this string str)
        {
            return str.Substring(0, 1).ToLower() + str.Substring(1);
        }
    }



}