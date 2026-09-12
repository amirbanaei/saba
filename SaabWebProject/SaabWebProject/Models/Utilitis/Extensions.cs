using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Utilitis
{
    public static class Extensions
    {
        public static string ToShamsi(this DateTime dateTime)
        {
             PersianCalendar pc = new PersianCalendar();

            var dt = dateTime;

            return (pc.GetYear(dt) + "/" + pc.GetMonth(dt) + "/" + pc.GetDayOfMonth(dt));

        }
        public static int GetShamsiDayOfMonth(this DateTime dateTime)
        {
            PersianCalendar pc = new PersianCalendar();

            var dt = dateTime;

            return pc.GetDayOfMonth(dt);

        }
        public static int GetShamsiMonth(this DateTime dateTime)
        {
            PersianCalendar pc = new PersianCalendar();

            var dt = dateTime;

            return pc.GetMonth(dt);

        }
        public static int GetShamsYear(this DateTime dateTime)
        {
            PersianCalendar pc = new PersianCalendar();

            var dt = dateTime;

            return pc.GetYear(dt);

        }

    }
}