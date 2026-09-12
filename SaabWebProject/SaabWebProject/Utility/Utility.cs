using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Ajax.Utilities;
using SaabWebProject.Utility;

namespace System
{
    public static class Utility
    {
        public static string GetTypeOfMoalefe(this int? value)
        {
            switch (value)
            {
                case (int)Type_Moalefe.Amalkardi:
                    {
                        return "عملکردی";
                    }
                case (int)Type_Moalefe.Controlli:
                    {
                        return "کنترلی";
                    }
                case (int)Type_Moalefe.Gharardadi:
                    {
                        return "قراردادی";
                    }
                case (int)Type_Moalefe.Karbari:
                    {
                        return "کاربری";
                    }
                case (int)Type_Moalefe.Sayer:
                    {
                        return "سایر";
                    }
                case (int)Type_Moalefe.Pishkhan:
                    {
                        return "پیشخوان";
                    }

                case (int)Type_Moalefe.Dastmozdi:
                    {
                        return "دستمزدی";
                    }
                case (int)Type_Moalefe.Karkardi:
                    {
                        return "کارکردی";
                    }
                case (int)Type_Moalefe.Calculational:
                    {
                        return "محاسباتی";
                    }
                case (int)Type_Moalefe.Fish:
                    {
                        return "فیش‏حقوقی";
                    }
                case (int)Type_Moalefe.sorat:
                    return "صورت وضعیت";
                   

            }

            return "نامشخص";
        }

        public static string GetPersianMonth(this int Month)
        {
            switch (Month)
            {
                case 1:
                    {
                        return "فروردین";
                    }
                case 2:
                    {
                        return "اردیبهشت";
                    }
                case 3:
                    {
                        return "خرداد";
                    }
                case 4:
                    {
                        return "تیر";
                    }
                case 5:
                    {
                        return "مرداد";
                    }
                case 6:
                    {
                        return "شهریور";
                    }
                case 7:
                    {
                        return "مهر";
                    }
                case 8:
                    {
                        return "آبان";
                    }
                case 9:
                    {
                        return "آذر";
                    }
                case 10:
                    {
                        return "دی";
                    }
                case 11:
                    {
                        return "بهمن";
                    }
                case 12:
                    {
                        return "اسفند";
                    }
                default:
                    {
                        return "";
                    }
            }
        }
    }
}