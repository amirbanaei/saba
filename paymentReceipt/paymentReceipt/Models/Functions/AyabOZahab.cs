using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace paymentReceipt.Models.Functions
{
    public class AyabOZahab
    {
        /// <summary>
        /// به ازای هر کار مبلغ ایاب و ذهاب تعلق گرفته را محاسبه می کند
        /// </summary>
        /// <param name="ayabozahab">مبلغ ایاب و ذهاب</param>
        /// <param name="count">تعداد ایاب و ذهاب </param>
        /// <returns></returns>
        public double ayabozahabPerWork(double ayabozahab, int count)
        {
            return ayabozahab * count;
        }
        /// <summary>
        /// جمع مبالغ تمام ایاب و ذهاب ها می باشد
        /// </summary>
        /// <param name="lstayabozahab">لیستی از ایاب و ذهاب های بدست آمده را دریافت می کند</param>
        /// <returns></returns>
        public double Totalayabozahab(List<double> lstayabozahab)
        {
            double result = 0;
            foreach (var item in lstayabozahab)
            {
                result = result + item;
            }
            return result;
        }
    }
}