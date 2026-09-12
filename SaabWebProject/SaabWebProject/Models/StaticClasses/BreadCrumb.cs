using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.StaticClasses
{
    public static class BreadCrumb
    {
        static string title = "";
        public static string LargModalTitle
        {
            get { return title; }
            set { title = value; }
        }

    }
}