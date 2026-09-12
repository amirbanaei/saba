using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace paymentReceipt.Models.Functions
{
    public class Performance
    {
        public Performance()
        {

        }
        /// <summary>
        /// استاندارد ماهانه را بدست می آورد
        /// </summary>
        /// <param name="WorkDay">تعداد روزهای کاری سال</param>
        /// <param name="DayStandard">استاندارد روزانه</param>
        /// <returns></returns>
        public double CreateMontlyStandard(int WorkDay,int DayStandard)
        {
            if(DayStandard==0)
            {
                return 0;
            }
            else
            {
                return WorkDay * DayStandard / 12;

            }
        }
        /// <summary>
        /// بدست آوردن میزان کارکرد به ازای هر کار
        /// </summary>
        /// <param name="work">تعداد کار</param>
        /// <param name="WorkDay">تعداد روزهای کاری سال</param>
        /// <param name="DayStandard">استاندارد روزانه</param>
        /// <returns></returns>
        public double CalculatePerformancePerWork(int DayStandard, int work,int workDay)
        {
            return work / CreateMontlyStandard( workDay,DayStandard);
        }
        /// <summary>
        /// جمع کارکرد حساب می کند که حاصل آن چند درصد از حقوق می باشد
        /// </summary>
        /// <param name="lstperformance">لیستی از کارکردهای بدست آمده بر اساس استانداردهای تعریف شده می باشد</param>
        /// <returns></returns>
        public double sumperformance(List<double> lstperformance)
        {
            double result=0;
            foreach (var item in lstperformance)
            {
                result = result + item;
            }
            return result;
        }
 



    }
}