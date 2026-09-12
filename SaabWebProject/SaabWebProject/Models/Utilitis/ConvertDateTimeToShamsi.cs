using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Utilitis
{
    public static class ConvertDateTimeToShamsi
    {
        static PersianCalendar pc = new PersianCalendar();

        public static string ConvertDateTimeToShamsi1(DateTime? dt1)
        {
            if(dt1 != null)
            {
                var dt = dt1 ?? DateTime.Now;

                return (pc.GetYear(dt) + "/" + pc.GetMonth(dt) + "/" + pc.GetDayOfMonth(dt));
            }
            else
            {
                return null;
            }
        }
        public static int ConvertDateTimeToYearShamsi(DateTime dt)
        {
            return pc.GetYear(dt);
        }

        public static DateTime ConvertShamsiYearToYear(int _shamsiYear , int _shamsiMonth)
        {
            return  new DateTime(_shamsiYear,_shamsiMonth,29, pc);
        }
      
    }
}