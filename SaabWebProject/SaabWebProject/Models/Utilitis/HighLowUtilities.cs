using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace SaabWebProject.Models.Utilitis
{
    public class HighLowUtilities
    {
        SaabEntities db;
        FishUtilities Fishuti;
        tbMoalefeValueFishRepositories moalefeValueFishRepo;
        public HighLowUtilities()
        {
            db = new SaabEntities();
        }
        public HighLowUtilities(SaabEntities Context)
        {
            db = Context;
            Fishuti = new FishUtilities(db);
            moalefeValueFishRepo = new tbMoalefeValueFishRepositories(db);
        }
        public ManualFishDetail CalculateandinserttotableForHighLow(int UserID, int Month, int Year, double Value, int MoalefeID)
        {            
            ManualFishDetail Result = new ManualFishDetail();


            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);
            var mozdsayer = "0";
            var haghmaskanbase = "0";
            var aghlammasrafikhanevarbase = "0";
            var hagholadbase = "0";
            float zaribkasrikar = 1;
            float zaribmamoriat = 1;
            float zaribqeybat = 1;





            bool Karmozdi = false;
            if (db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID).usc_TypeOfContract == true)
            {
                Karmozdi = true;
            }


            double tedadroozestelaji = 0;
            double tedadsaatezafkar = 0;
            double tedadroozmamoriat = 0;
            double tedadsaatkasrekar = 0;

            double MoadelKarkard = 0;
            double tedadroozQeybat = 0;
            double tedadrooztatil = 0;
            double teadaroozjome = 0;

            double ayabozahab = 0;
            double? padash = 0;
            double? jarime = 0;


            var moalefehayeexcel = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID && p.Finalaccept == true).ToList();

            double? Saatezafkar = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد ساعات اضافه کار")
            {
                Saatezafkar = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار") != null)
                {
                    Saatezafkar = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").mlfvlfsh_Value;
                }
            }



            double? stelaji = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز استعلاجی")
            {
                stelaji = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی") != null)
                {
                    stelaji = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی").mlfvlfsh_Value;
                }
            }



            double? roozmamoriat = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز ماموریت")
            {
                roozmamoriat = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت") != null)
                {
                    roozmamoriat = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت").mlfvlfsh_Value;
                }
            }



            double? saatkasrkar = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد ساعات کسر کار")
            {
                saatkasrkar = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
                {
                    saatkasrkar = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
                }
            }




            double? roozqybt = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز غیبت")
            {
                roozqybt = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت") != null)
                {
                    roozqybt = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت").mlfvlfsh_Value;
                }
            }





            double? ayabzahab = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه ایاب و ذهاب(ریال)")
            {
                ayabzahab = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                {
                    ayabzahab = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                }
                else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
                {
                    ayabzahab = Fishuti.CalculateAyabOzahab(UserID, Month, Year);
                }
            }





            double? jome = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد ساعات کارکرد جمعه")
            {
                jome = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه") != null)
                {
                    jome = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه").mlfvlfsh_Value;
                }
            }

            var tedadroozmah = Fishuti.TeadaroozMonth(Month, Year);




            double? tatil = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد ساعات کارکرد روز تعطیل")
            {

                tatil = Value;
            }
            else
            {
                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل") != null)
                {
                    tatil = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل").mlfvlfsh_Value;
                }
            }



            int teadadroozkarkard = Fishuti.teadadroozkarkardforfishmain((int)roozqybt, tedadroozmah, (int)stelaji);


            #region آپدیت کردن مولفه ولیو اکسل با های لو کاربر
            var exist = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == MoalefeID && p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year).FirstOrDefault();
            if (exist != null)
            {
                exist.MoalfeVal_Value = Value;
                db.SaveChanges();
            }
            #endregion

            var moadelkarkar = Fishuti.CalculateMoadelKarkard(UserID, Month, Year);




            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "پاداش(ریال)")
            {
                padash = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)") != null)
                {
                    padash = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)").mlfvlfsh_Value;
                }
                else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
                {
                    padash = Fishuti.CalculatePadashForkarmozdi(UserID, Month, Year);
                }
            }





            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جریمه(ریال)")
            {
                jarime = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)") != null)
                {
                    jarime = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)").mlfvlfsh_Value;
                }
                else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
                {
                    jarime = Fishuti.CalculateJarimeForKarmozdi(UserID, Month, Year);
                }
            }




            if (!Karmozdi)//روزمزدی
            {

                tedadroozestelaji = (double)stelaji;
                tedadsaatezafkar = (double)Saatezafkar;
                tedadroozmamoriat = (double)roozmamoriat;
                tedadsaatkasrekar = (double)saatkasrkar;


                tedadroozQeybat = (double)roozqybt;

                ayabozahab = (double)ayabzahab;
            }
            else//کارمزدی
            {
                tedadroozestelaji = (double)stelaji;

                var res = Fishuti.Calculatesaatkasrkarvaezafkarforkarmozdi(teadadroozkarkard, moadelkarkar, tedadroozmah, tedadroozestelaji);

                tedadsaatezafkar = res.Item1;
                tedadroozmamoriat = (double)roozmamoriat;
                tedadsaatkasrekar = res.Item2;


                MoadelKarkard = moadelkarkar;
                tedadroozQeybat = (double)roozqybt;
                tedadrooztatil = (double)tatil;
                teadaroozjome = (double)jome;

                ayabozahab = (double)ayabzahab;
            }








            var PeymanID = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == UserID && p.Status == true).FK_Peyman_ID;
            var zaribpeyman = db.tbPeymanZaribForFish.FirstOrDefault(p => p.FKPeyman == PeymanID);

            if (zaribpeyman != null)
            {
                zaribkasrikar = (float)zaribpeyman.pymnzrbfsh_zaribkasrekar;
                zaribqeybat = (float)zaribpeyman.pymnzrbfsh_zaribqeybat;
                zaribmamoriat = (float)zaribpeyman.pymnzrbfsh_zaribmamoriat;

            }
            List<FishValue> fishValues = new List<FishValue>();
            #region مولفه داینامیک قرارداد و ریختن در فیش ولیو
            var contractinfo = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_EndTime >= myDatetime && p.usc_StartTime <= myDatetime).FirstOrDefault();



            var listmoalefeqarardadi = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == contractinfo.usc_ID).ToList();
            List<FishValue> lstmoalefeqarardadfish = new List<FishValue>();
            foreach (var item in listmoalefeqarardadi)
            {
                FishValue moalefeqaradadfish = new FishValue
                {
                    Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.FKMoalefeGhararDadi).md_Title,
                    Value = (int)item.Value
                };
                lstmoalefeqarardadfish.Add(moalefeqaradadfish);
                fishValues.Add(moalefeqaradadfish);
            }

            Result.FishValueQaradad = lstmoalefeqarardadfish;
            #endregion
            #region مولفه های قرارداد

            var mozdsanavatroozane = Regex.Replace(contractinfo.jobgroup_ValueSanavat, ",", "");
            if (mozdsanavatroozane == null)
            {
                mozdsanavatroozane = "0";
            }
            var mozdshoqlroozane = Regex.Replace(contractinfo.jobgroup_ValueMozdGroup, ",", "");
            if (mozdshoqlroozane == null)
            {
                mozdshoqlroozane = "0";
            }
            if (contractinfo.jobgroup_MozdSayer != null)
            {
                mozdsayer = Regex.Replace(contractinfo.jobgroup_MozdSayer, ",", "");

                if (mozdsayer == null)
                {
                    mozdsayer = "0";
                }
            }
            if (contractinfo.jobgroup_HagheMaskan != null)
            {
                haghmaskanbase = Regex.Replace(contractinfo.jobgroup_HagheMaskan, ",", "");
                if (haghmaskanbase == null)
                {
                    haghmaskanbase = "0";
                }

            }
            if (contractinfo.jobgroup_HagheOlad != null)
            {

                hagholadbase = Regex.Replace(contractinfo.jobgroup_HagheOlad, ",", "");
                if (hagholadbase == null)
                {
                    hagholadbase = "0";
                }
            }
            if (contractinfo.jobgroup_KharoBar != null)
            {

                aghlammasrafikhanevarbase = Regex.Replace(contractinfo.jobgroup_KharoBar, ",", "");
                if (aghlammasrafikhanevarbase == null)
                {
                    aghlammasrafikhanevarbase = "0";
                }
            }
            #endregion


            #region مولفه هایی که در اکسل پر می شوند

            var mozdmabna = Convert.ToDouble(mozdsanavatroozane) + Convert.ToDouble(mozdshoqlroozane) + Convert.ToDouble(mozdsayer);
            var haghmorkhasi = ((((Convert.ToDouble(haghmaskanbase) + Convert.ToDouble(hagholadbase) + Convert.ToDouble(aghlammasrafikhanevarbase)) / 30) + mozdmabna) * 9) / 12;
            var ezafekar = (mozdmabna / 7.3333) * 1.4;
            var karkardjome = ezafekar * 1.4;
            var karkarrooztatil = ezafekar;




            FishValue valuefish11 = new FishValue
            {
                Title = "تعداد روز کارکرد",
                Value = teadadroozkarkard /*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز کارکرد").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish11);


            var nesbatkarbkolmah = Fishuti.Calculatekarkardtaghsimbarkolemah(Convert.ToInt32(valuefish11.Value), Fishuti.TeadaroozMonth(Month, Year));



            double? tedadrooznobatkari1 = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز نوبتکاری1 ( صبح وعصر)")
            {
                tedadrooznobatkari1 = Value;
            }
            else
            {

                var x = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری1 ( صبح وعصر)");
                if (x != null)
                {
                    tedadrooznobatkari1 = x.mlfvlfsh_Value;
                }
            }


            double? tedadrooznobatkari2 = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز نوبتکاری2 (صبح وشب)")
            {
                tedadrooznobatkari2 = Value;
            }
            else
            {

                var y = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری2 (صبح وشب)");
                if (y != null)
                {
                    tedadrooznobatkari2 = y.mlfvlfsh_Value;
                }
            }



            double? tedadrooznobatkari3 = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز نوبتکاری3 (عصروشب)")
            {
                tedadrooznobatkari3 = Value;
            }
            else
            {

                var z = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری3 (عصروشب)");
                if (z != null)
                {
                    tedadrooznobatkari3 = z.mlfvlfsh_Value;
                }
            }



            double? tedadrooznobatkari4 = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد روز نوبتکاری4 (صبح وعصروشب)")
            {
                tedadrooznobatkari4 = Value;
            }
            else
            {
                var xx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری4 (صبح وعصروشب)");
                if (xx != null)
                {
                    tedadrooznobatkari4 = xx.mlfvlfsh_Value;
                }


            }





            double? tedadroozshabkari = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "تعداد شب های کاری")
            {
                tedadroozshabkari = Value;
            }
            else
            {

                var xy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد شب های کاری");
                if (tedadroozshabkari == null)
                {
                    tedadroozshabkari = xy.mlfvlfsh_Value;
                }
            }





            double? xz = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "اصلاحیه کمک هزینه اولاد(ریال)")
            {
                xz = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)") != null)
                {
                    xz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)").mlfvlfsh_Value;
                }
            }

            FishValue eslahiyekomakhazineolad = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اولاد(ریال)",
                Value = (double)xz
            };
            fishValues.Add(eslahiyekomakhazineolad);





            double? xxx = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "اصلاحیه کمک هزینه مسکن(ریال)")
            {
                xxx = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)") != null)
                {
                    xxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)").mlfvlfsh_Value;
                }
            }
            FishValue valuefish2 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه مسکن(ریال)",
                Value = (double)xxx
            };
            fishValues.Add(valuefish2);





            double? xxy = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)")
            {
                xxy = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)") != null)
                {
                    xxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)").mlfvlfsh_Value;
                }
            }
            FishValue valuefish3 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)",
                Value = (double)xxy
            };
            fishValues.Add(valuefish3);







            FishValue valuefish4 = new FishValue
            {
                Title = "کمک هزینه ایاب و ذهاب(ریال)",
                Value = ayabozahab
            };
            fishValues.Add(valuefish4);





            FishValue valuefish5 = new FishValue
            {
                Title = "پاداش(ریال)",
                Value = (double)padash
            };
            fishValues.Add(valuefish5);




            FishValue valuefish6 = new FishValue
            {
                Title = "جریمه(ریال)",
                Value = (double)jarime
            };
            fishValues.Add(valuefish6);





            double? xyz = 0;


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه ابزار کار(ریال)")
            {
                xyz = Value;
            }
            else
            {
                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)") != null)
                {
                    xyz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)").mlfvlfsh_Value;
                }

            }
            FishValue valuefish7 = new FishValue
            {
                Title = "کمک هزینه ابزار کار(ریال)",
                Value = (double)xyz
            };
            fishValues.Add(valuefish7);





            double? yxx = 0;

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه تبلت و رایانه(ریال)")
            {
                yxx = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه") != null)
                {
                    yxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه").mlfvlfsh_Value * moadelkarkar;
                }
            }
            FishValue valuefish8 = new FishValue
            {
                Title = "کمک هزینه تبلت و رایانه(ریال)",
                Value = (double)yxx
            };
            fishValues.Add(valuefish8);




            double? yxy = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)")
            {
                yxy = Value;
            }
            else
            {

                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)") != null)
                {
                    yxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)").mlfvlfsh_Value;
                }
            }
            FishValue valuefish9 = new FishValue
            {
                Title = "کمک هزینه بیمه تکمیل درمان (ریال)",
                Value = (double)yxy
            };
            fishValues.Add(valuefish9);


            //سوال شود
            //FishValue valuefish10 = new FishValue
            //{
            //    Title = "اقساط بیمه تکمیل درمان (ریال)",
            //    Value = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اقساط بیمه تکمیل درمان (ریال)").mlfvlfsh_Value
            //};
            //fishValues.Add(valuefish10);






            FishValue valuefish12 = new FishValue
            {
                Title = "تعداد ساعات اضافه کار",
                Value = tedadsaatezafkar/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish12);

            FishValue valuefish13 = new FishValue
            {
                Title = "تعداد ساعات کسر کار",
                Value = tedadsaatkasrekar/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish13);

            FishValue valuefish14 = new FishValue
            {
                Title = "تعداد روز ماموریت",
                Value = tedadroozmamoriat/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish14);

            FishValue valuefish15 = new FishValue
            {
                Title = "تعداد روز استعلاجی",
                Value = tedadroozestelaji/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish15);

            FishValue valuefish16 = new FishValue
            {
                Title = "تعداد روز غیبت",
                Value = tedadroozQeybat/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish16);

            FishValue valuefish17 = new FishValue
            {
                Title = "تعداد ساعات کارکرد جمعه",
                Value = teadaroozjome/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish17);

            FishValue valuefish18 = new FishValue
            {
                Title = "تعداد ساعات کارکرد روز تعطیل",
                Value = tedadrooztatil/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل").mlfvlfsh_Value*/
            };
            #endregion
            #region سایر مولفه ها


            var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 / tedadroozmah);
            var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 / tedadroozmah);
            var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 / tedadroozmah);
            var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 / tedadroozmah);
            var shabkari = (mozdmabna * 35 / 100) * (tedadroozshabkari / tedadroozmah);
            var eydivapadash = mozdmabna * 5;
            var sanavatkhdmat = mozdmabna * 2.5;

            if (contractinfo != null)
            {
                // Result.FishValues = new List<FishValue>();
                FishValue fv = new FishValue

                {
                    Title = "مزد سنوات (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane)
                };

                fishValues.Add(fv);


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "نوبتکاری (ریال)")
                {
                    FishValue nobatkari = new FishValue

                    {
                        Title = "نوبتکاری (ریال)",
                        Value = Value
                    };
                    fishValues.Add(nobatkari);

                }
                else
                {
                    FishValue nobatkari = new FishValue

                    {
                        Title = "نوبتکاری (ریال)",
                        Value = (double)(nobatkari1 + nobatkari2 + nobatkari3 + nobatkari4 + shabkari)
                    };
                    fishValues.Add(nobatkari);
                }



                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "مزد سنوات (ریال)")
                {
                    FishValue fv85 = new FishValue

                    {
                        Title = "مزد سنوات (ریال)",
                        Value = Value
                    };

                    fishValues.Add(fv85);
                }
                else
                {
                    FishValue fv85 = new FishValue

                    {
                        Title = "مزد سنوات (ریال)",
                        Value = Convert.ToDouble(mozdsanavatroozane) * teadadroozkarkard
                    };

                    fishValues.Add(fv85);
                }





                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "مزد شغل (روزانه-ریال)")
                {
                    FishValue fv2 = new FishValue
                    {
                        Title = "مزد شغل (روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv2);
                }
                else
                {
                    FishValue fv2 = new FishValue
                    {
                        Title = "مزد شغل (روزانه-ریال)",
                        Value = Convert.ToDouble(mozdshoqlroozane)
                    };
                    fishValues.Add(fv2);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "مزد گروه (شغل)")
                {
                    FishValue fv200 = new FishValue
                    {
                        Title = "مزد گروه (شغل)",
                        Value = Value
                    };
                    fishValues.Add(fv200);
                }
                else
                {
                    FishValue fv200 = new FishValue
                    {
                        Title = "مزد گروه (شغل)",
                        Value = Convert.ToDouble(mozdshoqlroozane) * teadadroozkarkard
                    };
                    fishValues.Add(fv200);
                }




                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "مزد سایر ( روزانه-ریال)")
                {
                    FishValue fv3 = new FishValue
                    {
                        Title = "مزد سایر ( روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv3);
                }
                else
                {
                    FishValue fv3 = new FishValue
                    {
                        Title = "مزد سایر ( روزانه-ریال)",
                        Value = Convert.ToDouble(mozdsayer)
                    };
                    fishValues.Add(fv3);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جمع مزد مبنا ( روزانه-ریال)")
                {
                    FishValue fv4 = new FishValue
                    {
                        Title = "جمع مزد مبنا ( روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv4);
                }
                else
                {
                    FishValue fv4 = new FishValue
                    {
                        Title = "جمع مزد مبنا ( روزانه-ریال)",
                        Value = mozdmabna
                    };
                    fishValues.Add(fv4);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "اضافه کار ( هرساعت-ریال)")
                {
                    FishValue fv5 = new FishValue
                    {
                        Title = "اضافه کار ( هرساعت-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv5);
                }
                else
                {
                    FishValue fv5 = new FishValue
                    {
                        Title = "اضافه کار ( هرساعت-ریال)",
                        Value = Math.Round(ezafekar)
                    };
                    fishValues.Add(fv5);
                }




                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کارکرد جمعه (روزانه-ریال)")
                {
                    FishValue fv6 = new FishValue
                    {
                        Title = "کارکرد جمعه (روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv6);
                }
                else
                {
                    FishValue fv6 = new FishValue
                    {
                        Title = "کارکرد جمعه (روزانه-ریال)",
                        Value = karkardjome
                    };
                    fishValues.Add(fv6);
                }



                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کارکرد روز تعطیل (روزانه-ریال)")
                {
                    FishValue fv7 = new FishValue
                    {
                        Title = "کارکرد روز تعطیل (روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv7);
                }
                else
                {
                    FishValue fv7 = new FishValue
                    {
                        Title = "کارکرد روز تعطیل (روزانه-ریال)",
                        Value = karkarrooztatil
                    };
                    fishValues.Add(fv7);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "نوبتکاری1 ( صبح وعصر-روزانه-ریال)")
                {
                    FishValue fv8 = new FishValue
                    {
                        Title = "نوبتکاری1 ( صبح وعصر-روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv8);
                }
                else
                {
                    FishValue fv8 = new FishValue
                    {
                        Title = "نوبتکاری1 ( صبح وعصر-روزانه-ریال)",
                        Value = (double)nobatkari1
                    };
                    fishValues.Add(fv8);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "نوبتکاری2 (صبح وشب-روزانه-ریال)")
                {
                    FishValue fv9 = new FishValue
                    {
                        Title = "نوبتکاری2 (صبح وشب-روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv9);
                }
                else
                {
                    FishValue fv9 = new FishValue
                    {
                        Title = "نوبتکاری2 (صبح وشب-روزانه-ریال)",
                        Value = (double)nobatkari2
                    };
                    fishValues.Add(fv9);
                }




                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "نوبتکاری3 (عصروشب-روزانه-ریال)")
                {
                    FishValue fv10 = new FishValue
                    {
                        Title = "نوبتکاری3 (عصروشب-روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv10);
                }
                else
                {
                    FishValue fv10 = new FishValue
                    {
                        Title = "نوبتکاری3 (عصروشب-روزانه-ریال)",
                        Value = (double)nobatkari3
                    };
                    fishValues.Add(fv10);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)")
                {
                    FishValue fv11 = new FishValue
                    {
                        Title = "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv11);
                }
                else
                {
                    FishValue fv11 = new FishValue
                    {
                        Title = "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)",
                        Value = (double)nobatkari4
                    };
                    fishValues.Add(fv11);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "شبکاری (هرشب-ریال)")
                {
                    FishValue fv12 = new FishValue
                    {
                        Title = "شبکاری (هرشب-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv12);
                }
                else
                {
                    FishValue fv12 = new FishValue
                    {
                        Title = "شبکاری (هرشب-ریال)",
                        Value = (double)shabkari
                    };
                    fishValues.Add(fv12);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "عیدی و پاداش ( ماهانه-ریال)")
                {
                    FishValue fv13 = new FishValue
                    {
                        Title = "عیدی و پاداش ( ماهانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv13);
                }
                else
                {
                    FishValue fv13 = new FishValue
                    {
                        Title = "عیدی و پاداش ( ماهانه-ریال)",
                        Value = eydivapadash * nesbatkarbkolmah
                    };
                    fishValues.Add(fv13);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "سنوات خدمت (ماهانه-ریال)")
                {
                    FishValue fv14 = new FishValue
                    {
                        Title = "سنوات خدمت (ماهانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv14);
                }
                else
                {
                    FishValue fv14 = new FishValue
                    {
                        Title = "سنوات خدمت (ماهانه-ریال)",
                        Value = sanavatkhdmat * nesbatkarbkolmah
                    };
                    fishValues.Add(fv14);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "حق مرخصی (ریال)")
                {
                    FishValue fv15 = new FishValue
                    {
                        Title = "حق مرخصی (ریال)",
                        Value = Value

                    };
                    fishValues.Add(fv15);
                }
                else
                {
                    FishValue fv15 = new FishValue
                    {
                        Title = "حق مرخصی (ریال)",
                        Value = haghmorkhasi * nesbatkarbkolmah

                    };
                    fishValues.Add(fv15);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه مسکن (ریال)")
                {
                    FishValue fv16 = new FishValue
                    {
                        Title = "کمک هزینه مسکن (ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv16);
                }
                else
                {
                    FishValue fv16 = new FishValue
                    {
                        Title = "کمک هزینه مسکن (ریال)",
                        Value = Convert.ToDouble(haghmaskanbase) * nesbatkarbkolmah
                    };
                    fishValues.Add(fv16);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه اقلام مصرفی خانوار(ریال)")
                {
                    FishValue f17 = new FishValue
                    {
                        Title = "کمک هزینه اقلام مصرفی خانوار(ریال)",
                        Value = Value
                    };
                    fishValues.Add(f17);
                }
                else
                {
                    FishValue f17 = new FishValue
                    {
                        Title = "کمک هزینه اقلام مصرفی خانوار(ریال)",
                        Value = Convert.ToDouble(aghlammasrafikhanevarbase) * nesbatkarbkolmah
                    };
                    fishValues.Add(f17);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کمک هزینه اولاد(ریال)")
                {
                    FishValue fv18 = new FishValue
                    {
                        Title = "کمک هزینه اولاد(ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv18);
                }
                else
                {
                    FishValue fv18 = new FishValue
                    {
                        Title = "کمک هزینه اولاد(ریال)",
                        Value = Convert.ToDouble(hagholadbase) * nesbatkarbkolmah
                    };
                    fishValues.Add(fv18);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کسر کار (هرساعت-ریال)")
                {
                    FishValue fv19 = new FishValue
                    {
                        Title = "کسر کار (هرساعت-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv19);
                }
                else
                {
                    FishValue fv19 = new FishValue
                    {
                        Title = "کسر کار (هرساعت-ریال)",
                        Value = (mozdmabna / 7.3333) * zaribkasrikar
                    };
                    fishValues.Add(fv19);
                }



                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "ماموریت ( روزانه-ریال)")
                {
                    FishValue fv20 = new FishValue
                    {
                        Title = "ماموریت ( روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv20);
                }
                else
                {
                    FishValue fv20 = new FishValue
                    {
                        Title = "ماموریت ( روزانه-ریال)",
                        Value = (mozdmabna * zaribmamoriat)
                    };
                    fishValues.Add(fv20);
                }


                if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "غیبت (روزانه-ریال)")
                {
                    FishValue fv21 = new FishValue
                    {
                        Title = "غیبت (روزانه-ریال)",
                        Value = Value
                    };
                    fishValues.Add(fv21);
                }
                else
                {
                    FishValue fv21 = new FishValue
                    {
                        Title = "غیبت (روزانه-ریال)",
                        Value = (mozdmabna * zaribqeybat)
                    };
                    fishValues.Add(fv21);
                }



            }
            #endregion


            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "ماموریت (ریال)")
            {
                FishValue Mamoriat = new FishValue
                {
                    Title = "ماموریت (ریال)",
                    Value = Value
                };
                fishValues.Add(Mamoriat);
            }
            else
            {
                FishValue Mamoriat = new FishValue
                {
                    Title = "ماموریت (ریال)",
                    Value = Fishuti.CalculateMamoriat((mozdmabna * zaribmamoriat), (double)roozmamoriat)
                };
                fishValues.Add(Mamoriat);
            }



            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "اضافه کاری (ریال)")
            {
                FishValue ezafkari = new FishValue
                {
                    Title = "اضافه کاری (ریال)",
                    Value = Value
                };
                fishValues.Add(ezafkari);
            }
            else
            {
                FishValue ezafkari = new FishValue
                {
                    Title = "اضافه کاری (ریال)",
                    Value = Fishuti.CalculateEzafkar(ezafekar, tedadsaatezafkar)
                };
                fishValues.Add(ezafkari);
            }





            #region پر کردن اضافات و ریختن در فیش ولیو
            var ezafat = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
            List<FishValue> ezafatValue = new List<FishValue>();
            foreach (var item in ezafat)
            {
                FishValue ezafvalue = new FishValue
                {
                    Title = item.md_Title,
                    Value = (double)db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID).mlfvlfsh_Value
                };
                ezafatValue.Add(ezafvalue);
                fishValues.Add(ezafvalue);
            }
            Result.FishValuesEzafat = ezafatValue;
            #endregion
            #region  پر کردن کسورات و ریختن در فیش ولیو
            var kosoratlist = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 1 && p.IsMain != true).ToList();
            List<FishValue> koosratvalue = new List<FishValue>();
            foreach (var item in kosoratlist)
            {
                FishValue kasrvalue = new FishValue
                {
                    Title = item.md_Title,
                    Value = (double)db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID).mlfvlfsh_Value
                };
                koosratvalue.Add(kasrvalue);
                fishValues.Add(kasrvalue);
            }
            Result.FishValuesKosoorat = koosratvalue;
            #endregion


            #region پر کردن جدول برای محاسبه بیمه و مالیات
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in fishValues)
                    {
                        var fkMoalefe = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == item.Title).md_ID;
                        var existMoalefe = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (existMoalefe == null)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                            {
                                FK_Moalefe = fkMoalefe,
                                mlfvlfsh_Submit = false,
                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishRepo.Create(moalefefish);
                        }
                        else if (existMoalefe.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            existMoalefe.FK_Moalefe = fkMoalefe;
                            existMoalefe.mlfvlfsh_Submit = false;
                            existMoalefe.FK_User = UserID;
                            existMoalefe.mlfvlfsh_Month = Month;
                            existMoalefe.mlfvlfsh_Year = Year;
                            existMoalefe.mlfvlfsh_Value = item.Value;
                            db.SaveChanges();
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception)
                {

                    transaction.Rollback();
                }
            }

            #endregion




            #region محاسبه بیمه
            double jamebime = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جمع کل مشمول بیمه (ریال)")
            {
                jamebime = Value;
            }
            else
            {
                var bimelist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 1 || p.included == 3).ToList();
                foreach (var item in bimelist)
                {

                    var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID);

                    if (varObject != null)
                    {
                        jamebime += (double)varObject.mlfvlfsh_Value;
                    }
                    else
                    {
                        jamebime += 0;

                    }
                }

            }
            FishValue valuejamebime = new FishValue
            {
                Title = "جمع کل مشمول بیمه (ریال)",
                Value = Math.Round(jamebime)
            };
            fishValues.Add(valuejamebime);

            FishValue bimesahmkarmand = new FishValue();

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "حق بیمه سهم کارمند (ریال)")
            {

                bimesahmkarmand.Title = "حق بیمه سهم کارمند (ریال)";
                bimesahmkarmand.Value = Value;

                fishValues.Add(bimesahmkarmand);
            }
            else
            {
                bimesahmkarmand.Title = "حق بیمه سهم کارمند (ریال)";
                bimesahmkarmand.Value = Math.Round(Fishuti.CalculateBime(jamebime, Year));

                fishValues.Add(bimesahmkarmand);
            }

            #endregion



            #region محاسبه مالیات

            var eydivapadashsalane = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "عیدی و پاداش (سالیانه - ریال)")
            {
                eydivapadashsalane = Convert.ToInt32(Value);
            }
            else
            {
                int startmonth = contractinfo.ShamsiStartTime_month;
                var Moalefeeydivapadash = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "عیدی و پاداش ( ماهانه-ریال)");
                int moalefeeydivapadashID = 0;
                if (Moalefeeydivapadash != null)
                {
                    moalefeeydivapadashID = Moalefeeydivapadash.md_ID;
                    if (Month == 12)
                    {
                        for (int i = startmonth; i < 13; i++)
                        {

                            eydivapadashsalane += Convert.ToInt32(db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == moalefeeydivapadashID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == i && p.FK_User == UserID).mlfvlfsh_Value);
                        }
                    }
                }

            }



            FishValue eydivapadashsaliyaneh = new FishValue
            {
                Title = "عیدی و پاداش (سالیانه - ریال)",
                Value = eydivapadashsalane
            };
            fishValues.Add(eydivapadashsaliyaneh);






            double jammaliat = 0;
            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جمع کل مشمول مالیات (ریال)")
            {
                FishValue valuejammaliat = new FishValue
                {
                    Title = "جمع کل مشمول مالیات (ریال)",
                    Value = Value
                };
                fishValues.Add(valuejammaliat);
            }
            else
            {

                var maliatlist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 2 || p.included == 3).ToList();
                foreach (var item in maliatlist)
                {

                    var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID);

                    if (varObject != null)
                    {
                        jammaliat += (double)varObject.mlfvlfsh_Value;
                    }
                    else
                    {
                        jammaliat += 0;

                    }
                }
                jammaliat = jammaliat - ((7 / 7) * bimesahmkarmand.Value) - Fishuti.getBimetakmiliValue(Year, Month, UserID);
                FishValue valuejammaliat = new FishValue
                {
                    Title = "جمع کل مشمول مالیات (ریال)",
                    Value = Math.Round(jammaliat + eydivapadashsalane)
                };
                fishValues.Add(valuejammaliat);
            }




            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "مالیات سهم کارمند (ریال)")
            {
                FishValue maliatsahmkarmand = new FishValue
                {
                    Title = "مالیات سهم کارمند (ریال)",
                    Value = Value
                };
                fishValues.Add(maliatsahmkarmand);
            }
            else
            {
                FishValue maliatsahmkarmand = new FishValue
                {
                    Title = "مالیات سهم کارمند (ریال)",
                    Value = Fishuti.CalculateTaxTajamoe(jammaliat, Year, Month, UserID, eydivapadashsalane)
                };
                fishValues.Add(maliatsahmkarmand);
            }



            #endregion

            #region محاسبه اقساط و ریختن در فیش ولیو
            var tamamaqsat = Fishuti.CalculateAqsat(Year, Month, UserID);
            List<FishValue> aqsatformohasebekasriha = new List<FishValue>();
            foreach (var item in tamamaqsat)
            {
                FishValue aqsat = new FishValue
                {
                    Title = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item.Item2).md_Title,
                    Value = item.Item1
                };
                aqsatformohasebekasriha.Add(aqsat);
                fishValues.Add(aqsat);
            }
            //TODO:Felan badan bayad pakshavad
            FishValue felan = new FishValue
            {
                Title = "صندوق خیریه شرکت",
                Value = 200000
            };
            aqsatformohasebekasriha.Add(felan);
            ///
            Result.FishValuesAqsat = aqsatformohasebekasriha;

            #endregion

            #region قسمت های محاسباتی نهایی

            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)")
            {
                Result.ZakhireKarMazadQabl = Value;
            }
            else
            {
                Result.ZakhireKarMazadQabl = Fishuti.CalculateZakhireKarMazadQabl(UserID, Year, Month);

            }
            FishValue ZakhireKarMazadQabl = new FishValue
            {
                Title = "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)",
                Value = Result.ZakhireKarMazadQabl
            };
            fishValues.Add(ZakhireKarMazadQabl);



            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جمع ناخالص حقوق و مزایا (ریال)")
            {
                Result.JamNakhalesHoqoqVaMazaya = Value;
            }
            else
            {

                Result.JamNakhalesHoqoqVaMazaya = Fishuti.CalculateJamNakhalesHoqoqVaMazaya(fishValues, lstmoalefeqarardadfish, Result.ZakhireKarMazadQabl, ezafatValue, nesbatkarbkolmah, contractinfo.usc_ID);
            }
            FishValue JamNakhalesHoqoqVaMazaya = new FishValue
            {
                Title = "جمع ناخالص حقوق و مزایا (ریال)",
                Value = Result.JamNakhalesHoqoqVaMazaya
            };
            fishValues.Add(JamNakhalesHoqoqVaMazaya);




            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "کسری کارکرد (ریال)")
            {
                Result.kasrikarkard = Value;
            }
            else
            {

                Result.kasrikarkard = Fishuti.Calculatekasrikarkard((int)tedadsaatkasrekar, (mozdmabna / 7.3333) * zaribkasrikar);
            }

            FishValue kasrikarkard = new FishValue
            {
                Title = "کسری کارکرد (ریال)",
                Value = Result.kasrikarkard
            };
            fishValues.Add(kasrikarkard);



            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "ذخیره کار مازاد (ریال)")
            {
                Result.ZakhireKarMazad = Value;
            }
            else
            {

                Result.ZakhireKarMazad = Fishuti.CalculateZakhireKarMazad(Fishuti.CalculatejamkolekosooratbdoonMaxpay(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha), Result.JamNakhalesHoqoqVaMazaya, UserID, Year, Month);
            }

            FishValue ZakhireKarMazad = new FishValue
            {
                Title = "ذخیره کار مازاد (ریال)",
                Value = Result.ZakhireKarMazad
            };
            fishValues.Add(ZakhireKarMazad);




            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "جمع کل کسورات (ریال)")
            {
                Result.jamkolekosoorat = Value;
            }
            else
            {

                Result.jamkolekosoorat = Fishuti.Calculatejamkolekosoorat(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha, Result.ZakhireKarMazad);
            }
            FishValue jamkolekosoorat = new FishValue
            {
                Title = "جمع کل کسورات (ریال)",
                Value = Result.jamkolekosoorat
            };
            fishValues.Add(jamkolekosoorat);



            if (db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == MoalefeID).md_Title == "خالص قابل دریافت (ریال)")
            {
                Result.KhalesQabelDaryaft = Value;
            }
            else
            {

                Result.KhalesQabelDaryaft = Fishuti.CalculateKhalesQabelDaryaft(Result.jamkolekosoorat, Result.JamNakhalesHoqoqVaMazaya);
            }
            FishValue KhalesQabelDaryaft = new FishValue
            {
                Title = "خالص قابل دریافت (ریال)",
                Value = Result.KhalesQabelDaryaft
            };
            fishValues.Add(KhalesQabelDaryaft);
            #endregion

            #region پرکردن جدول مولفه ولیو فیش
            Result.FishValues = fishValues;
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in fishValues)
                    {
                        var fkMoalefe = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == item.Title).md_ID;
                        var existvalue = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (existvalue == null)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                            {
                                FK_Moalefe = fkMoalefe,
                                mlfvlfsh_Submit = false,
                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishRepo.Create(moalefefish);
                        }
                        else if (existvalue.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            existvalue.FK_Moalefe = fkMoalefe;
                            existvalue.mlfvlfsh_Submit = false;
                            existvalue.FK_User = UserID;
                            existvalue.mlfvlfsh_Month = Month;
                            existvalue.mlfvlfsh_Year = Year;
                            existvalue.mlfvlfsh_Value = item.Value;
                            db.SaveChanges();
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception)
                {

                    transaction.Rollback();
                }
            }
            #endregion





            return Result;








        }
    }
}