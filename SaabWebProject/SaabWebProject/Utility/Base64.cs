using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Utility
{
    public class Base64
    {
        public static string Base64Encode(string text)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(text);
            var res = System.Convert.ToBase64String(plainTextBytes);
            var u = res.Count(p => p.ToString() == "=");
            if (u != 0)
            {
                for (var i = 0; i < u; i++)
                {
                    res = res.Replace('=', ' ').Trim() + ":";
                }
            }

            return res;
        }
        public static string Base64Decode(string text)
        {
            var res = text;
            var u = res.Count(p => p.ToString() == ":");
            if (u != 0)
            {
                for (var i = 0; i < u; i++)
                {
                    res = res.Replace(':', ' ').Trim() + "=";
                }
            }
            var base64EncodedBytes = System.Convert.FromBase64String(res);
            res = System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
            return res;
        }
    }
}