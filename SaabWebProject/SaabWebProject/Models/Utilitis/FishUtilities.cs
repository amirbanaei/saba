using ExcelLibrary.CompoundDocumentFormat;
using Microsoft.Ajax.Utilities;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Data.Entity.Core.Objects;
using System.Data;
using System.Data.Entity.Infrastructure;
using Stimulsoft.Blockly.Model;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using static SkiaSharp.HarfBuzz.SKShaper;
using OfficeOpenXml.Drawing.Chart;
using System.Web.UI;

namespace SaabWebProject.Models.Utilitis
{
    public class FishUtilities
    {
        SaabEntities db=new SaabEntities(); 
        tbMoalefeValueFishRepositories moalefeValueFishRepo;
        tbMoalefeValueFishTestRepositories moalefeValueFishTestRepo;
        tbStepTaxRepository stepTaxRepository = new tbStepTaxRepository();
        tbUserSalaryDetailRepositories tbusersalarydetailtRepo;
        tbSaleryIsCalculatedRepository saleryiscalcRepo;
        public FishUtilities(SaabEntities Context)
        {
           
            moalefeValueFishRepo = new tbMoalefeValueFishRepositories(db);
            moalefeValueFishTestRepo = new tbMoalefeValueFishTestRepositories(db);
            tbusersalarydetailtRepo = new tbUserSalaryDetailRepositories(db);
            saleryiscalcRepo = new tbSaleryIsCalculatedRepository(db);

        }
        #region فیش قطعی
        public static DateTime ConvertShamsiToMiladi(int year, int month)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            DateTime miladiDate = persianCalendar.ToDateTime(year, month, 29, 0, 0, 0, 0);
            return miladiDate;
        }
        public ManualFishDetail Calculateandinserttotable(int UserID, int Month, int Year, bool qatee, FishHeader fishHeader)
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



            //PersianCalendar persianCalendar = new PersianCalendar();


            //int contractYear = persianCalendar.GetYear(endTime);
            //int contractMonth = persianCalendar.GetMonth(endTime);
        

            bool Karmozdi = false;
            DateTime miladiDate = ConvertShamsiToMiladi(Year, Month);

            var t = db.tbUserContracts.Where(p => p.FK_UserID == UserID && miladiDate >= p.usc_StartTime && miladiDate <= p.usc_EndTime&& p.usc_TypeOfContract == true).FirstOrDefault();
            if (t!=null)
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


            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار") != null)
            {
                Saatezafkar = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").mlfvlfsh_Value;
            }


            double? stelaji = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی") != null)
            {
                stelaji = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی").mlfvlfsh_Value;
            }


            double? roozmamoriat = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت") != null)
            {
                roozmamoriat = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت").mlfvlfsh_Value;
            }

            double? saatkasrkar = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }

            double? roozqybt = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت") != null)
            {
                roozqybt = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت").mlfvlfsh_Value;
            }


            int ayabzohabnadashtehbashwe = 1;
            
            double? ayabzahab = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
            {
                if(db.tbPeymanContracts.Where(s=>s.pec_Title== "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
                {
                    ayabzahab = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                }
                else
                {
                    ayabzahab = 0;

                }
                //ayabzahab = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
                {
                    ayabzahab = CalculateAyabOzahab(UserID, Month, Year);

                }
                else
                {
                    ayabzahab = 0;

                }
                //ayabzahab = CalculateAyabOzahab(UserID, Month, Year);
            }



            double? jome = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه") != null)
            {
                jome = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه").mlfvlfsh_Value;
            }

            var tedadroozmah = TeadaroozMonth(Month, Year);


            double? tatil = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل") != null)
            {
                tatil = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل").mlfvlfsh_Value;
            }

            var moadelkarkar = CalculateMoadelKarkard(UserID, Month, Year);

            if (Month == 1)
            {
                var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == 12 && p.Year == Year - 1).FirstOrDefault();
                if ( ezaf != null)
                {if(ezaf.RemainValue != null)
                    {
                        moadelkarkar += (double)(ezaf.RemainValue / 162);
                    }
                }
                var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                //if (tedadsaatezafkar > exi.Value && exi != null)
                //{


                //    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                //    tedadsaatezafkar = (double)exi.Value;
                //    db.SaveChanges();



                //}
            }
            else
            {
                var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month - 1 && p.Year == Year).FirstOrDefault();
                if ( ezaf != null)
                {if(ezaf.RemainValue != null )
                    {
                        moadelkarkar += (double)(ezaf.RemainValue / 162);

                    }
                }
                //var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                //if (tedadsaatezafkar > exi.Value && exi != null)
                //{


                //    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                //    tedadsaatezafkar = (double)exi.Value;
                //    db.SaveChanges();



                //}
            }
            //int teadadroozkarkard = teadadroozkarkardforfishmain((int)roozqybt, tedadroozmah, (int)stelaji);

            int teadadroozkarkard = teadadroozkarkardforfishmainasly((int)roozqybt, tedadroozmah, (int)stelaji, (double)moadelkarkar);





            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)") != null)
            {
                padash = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                padash = CalculatePadashForkarmozdi(UserID, Month, Year);
            }



            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)") != null)
            {
                jarime = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                jarime = CalculateJarimeForKarmozdi(UserID, Month, Year);
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

                var res = Calculatesaatkasrkarvaezafkarforkarmozdi(teadadroozkarkard, moadelkarkar, tedadroozmah, tedadroozestelaji);
                tedadsaatezafkar = res.Item1;
                #region برای max اضاقه کاری

            






                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        if (Month == 1)
                        {
                            //var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == 12 && p.Year == Year - 1).FirstOrDefault();
                            //if (ezaf.RemainValue != null && ezaf!=null)
                            //{
                            //    tedadsaatezafkar +=(double) ezaf.RemainValue;
                            //}
                            var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
                            if (exi != null)
                            {
                                if (tedadsaatezafkar > exi.Value && exi != null)
                                {


                                    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                                    tedadsaatezafkar = (double)exi.Value;
                                    db.SaveChanges();



                                }
                            }
                         
                        }
                        else
                        {
                            //var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month - 1 && p.Year == Year).FirstOrDefault();
                            //if (ezaf.RemainValue != null && ezaf != null)
                            //{
                            //    tedadsaatezafkar += (double)ezaf.RemainValue;
                            //}
                            var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
                            if (exi != null)
                            {
                                if (tedadsaatezafkar > exi.Value && exi != null)
                                {


                                    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                                    tedadsaatezafkar = (double)exi.Value;
                                    db.SaveChanges();



                                }
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


                tedadroozmamoriat = (double)roozmamoriat;
                //tedadsaatkasrekar = res.Item2;


                MoadelKarkard = moadelkarkar;
                tedadroozQeybat = (double)roozqybt;
                tedadrooztatil = (double)tatil;
                teadaroozjome = (double)jome;

                ayabozahab = (double)ayabzahab;
            }



            var nesbatkarbkolmah = Calculatekarkardtaghsimbarkolemah(teadadroozkarkard, TeadaroozMonth(Month, Year));





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
                var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == item.FKMoalefeGhararDadi && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                if (exist == null)
                {
                    tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                    {
                        FK_Moalefe = item.FKMoalefeGhararDadi,
                        mlfvlfsh_Submit = qatee,
                        FK_User = UserID,
                        mlfvlfsh_Month = Month,
                        mlfvlfsh_Year = Year,
                        Finalaccept = true,

                        mlfvlfsh_Value = item.Value * nesbatkarbkolmah
                    };
                    db.tbMoalefeValueFish.Add(moalefefish);
                    db.SaveChanges();
                }
                else if (exist.mlfvlfsh_Submit != true)
                {
                    exist.FK_Moalefe = item.FKMoalefeGhararDadi;
                    exist.mlfvlfsh_Submit = qatee;
                    exist.FK_User = UserID;
                    exist.Finalaccept = true;

                    exist.mlfvlfsh_Month = Month;
                    exist.mlfvlfsh_Year = Year;
                    exist.mlfvlfsh_Value = item.Value * nesbatkarbkolmah;
                    db.SaveChanges();
                }
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





            double? tedadrooznobatkari1 = 0;
            var x = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری1 ( صبح‏وعصر)");
            if (x != null)
            {
                tedadrooznobatkari1 = x.mlfvlfsh_Value;
            }


            double? tedadrooznobatkari2 = 0;
            var y = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری2 (صبح‏وشب)");
            if (y != null)
            {
                tedadrooznobatkari2 = y.mlfvlfsh_Value;
            }



            double? tedadrooznobatkari3 = 0;
            var z = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری3 (عصروشب)");
            if (z != null)
            {
                tedadrooznobatkari3 = z.mlfvlfsh_Value;
            }



            double? tedadrooznobatkari4 = 0;
            var xx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری4 (صبح‏وعصروشب)");
            if (xx != null)
            {
                tedadrooznobatkari4 = xx.mlfvlfsh_Value;
            }




            double? tedadroozshabkari = 0;
            var xy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد شب های کاری");
            if (tedadroozshabkari == null)
            {
                tedadroozshabkari = xy.mlfvlfsh_Value;
            }


            double? xz = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)") != null)
            {
                xz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)").mlfvlfsh_Value;
            }
            FishValue eslahiyekomakhazineolad = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اولاد(ریال)",
                Value = (double)xz
            };
            fishValues.Add(eslahiyekomakhazineolad);





            double? xxx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)") != null)
            {
                xxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish2 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه مسکن(ریال)",
                Value = (double)xxx
            };
            fishValues.Add(valuefish2);





            double? xxy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)") != null)
            {
                xxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish3 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)",
                Value = (double)xxy
            };
            fishValues.Add(valuefish3);






            if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",
                    Value = ayabozahab
                    //Value = 0

                };
                fishValues.Add(valuefish4);

            }
            else
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",
                    //Value = ayabozahab
                    Value = 0

                };
                fishValues.Add(valuefish4);

            }





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
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)") != null)
            {
                xyz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish7 = new FishValue
            {
                Title = "کمک هزینه ابزار کار(ریال)",
                //Value = (double)xyz
                Value = 0
            };
            fishValues.Add(valuefish7);





            double? yxx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه") != null)
            {
                yxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه").mlfvlfsh_Value;
            }
            FishValue valuefish8 = new FishValue
            {
                Title = "کمک هزینه تبلت و رایانه(ریال)",
                /* Value = (double)yxx *//** moadelkarkar*/
                Value = 0
            };
            fishValues.Add(valuefish8);




            double? yxy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)") != null)
            {
                yxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)").mlfvlfsh_Value;
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

            string inputDate = "1" + "/" + Month + "/" + Year;
            DateTime persianDate = PersianDateToDateTime(inputDate);
            double ayabzahab2 = 0;
            var ayabzahab2000 = db.tbMoadelPadashJarimeAyab.Where(p=>p.Month==Month&&p.Year==Year&&p.UserID==UserID).FirstOrDefault();
            if (ayabzahab2000 != null)
            {
                ayabzahab2= db.tbMoadelPadashJarimeAyab.Where(p => p.Month == Month && p.Year == Year && p.UserID == UserID).Select(s=>s.AyabOZahab).FirstOrDefault();

            }
             var tbfkfinancial = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();
            var fin = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5001).FirstOrDefault();
            var fin2 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5002).FirstOrDefault();
            var find3= tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5003).FirstOrDefault();
            var find4 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5004).FirstOrDefault();
            var find5 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5005).FirstOrDefault();


            if (fin != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == fin.ID).FirstOrDefault();
                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9).FirstOrDefault();
                if (find != null )
                {
                    if (findsss != null)
                    {

                        find.Creditor = (float)ayabzahab2;
                        db.SaveChanges();
                    }

                }
                else
                {
                    if (findsss != null)
                    {

                      FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = fin.ID;
                    FinancialDocumentss.Creditor = (float)ayabzahab2;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "خودروشخصی";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();
                    }
                  

                }
            }
            else
            {
                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9).FirstOrDefault();

                if (findsss != null)
                {
                    tbfkfinancial fkfinancial = new tbfkfinancial();
                    fkfinancial.Title = "خودروشخصی";
                    fkfinancial.DataDocument = persianDate;
                    fkfinancial.numbershomar = 5001;
                    db.tbfkfinancial.Add(fkfinancial);
                    db.SaveChanges();

                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = fkfinancial.ID;
                    FinancialDocumentss.Creditor = (float)ayabzahab2;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "خودروشخصی";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();
                }

                 


            }
            long tablett = 0;
            float mo = 0;
            var ezafggh = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
            if(ezafggh != null)
            {
                if (ezafggh.Value != null)
                {
                   var f= (ezafggh.Value / 162)+1;
                    if (moadelkarkar > f)
                    {
                        mo = (float)f;
                    }
                    else
                    {
                        mo =(float) moadelkarkar;
                    }
                }
            }
            float moh = 0;

            if (fin2 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == fin2.ID ).FirstOrDefault();
                if (db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 2).FirstOrDefault() != null)
                {
                    tablett = 5000000;
                }
                else if (db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 1).FirstOrDefault() != null)
                {
                    tablett = 2000000;
                }
                moh=(float)(tablett * mo);
                if (find != null)
                {
                    find.Creditor = (float)(tablett* mo);
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = fin2.ID;
                    FinancialDocumentss.Creditor = (float)(tablett * mo);
                        FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره تیلت";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره تیلت";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5002;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)(tablett * mo);
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره تیلت";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }



            var py = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == UserID && p.Status == true).FirstOrDefault();

            long mashincol = 0;
            long mashincolsogt = 0;

            long mashincolegareh = 0;
            long mashincouniform = 0;

            int mon = 0;
            int yr = 0;
            if (Month == 1)
            {
                mon = 12;
                yr = Year - 1;
            }
            else
            {
                mon = Month-1;
                yr = Year;
            }
            var mashin = db.tbReffrenceSave.Where(p => p.FK_PeymanID == py.tbPeymanContracts.pec_ID && p.IsFor == 3).FirstOrDefault();
            if (mashin != null)
            {
                var mashin2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == mashin.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                if (mashin2 != null) { 

                    var mashin3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == mashin2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == mon && s.Year == yr)).FirstOrDefault();
                if (mashin3 != null)
                {
                    var mashin4 = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == mashin3.ID && p.CountDays != 0&&p.Fk_user==UserID).ToList();
                    foreach (var item6 in mashin4)
                    {
                        if (item6.CountDays != 0&&item6.CountDays!=null)
                        {
                            var rx = db.tbEquipments.Where(p => p.ID == item6.FK_Equipment && p.FK_Peyman== py.tbPeymanContracts.pec_ID).FirstOrDefault();
                            if (rx != null)
                            {
                                mashincolsogt += (long)(item6.CountDays * (rx.Soght * 12 / 366));
                                mashincolegareh += (long)(item6.CountDays * (rx.egareh * 12 / 366));
                                if(rx.uniform!=null&& rx.uniform != 0)
                                {
                                    mashincouniform += (long)(item6.CountDays * (rx.uniform * 12 / 366));

                                }
                              
                                var c55c = rx.EachValue * 12 / 366;
                            }
                        }

                    }
                }
                }
            }
            long mohas = mashincouniform + mashincolegareh + mashincolsogt;
            mashincol += mohas;

            //var kasr2 = db.KasriMashin.Where(p => p.Fk_pymn == item.pec_ID && p.number_sorat == number_sorat).FirstOrDefault();
            //if (kasr2 != null && kasr2.Value != null)
            //{
            //    mashincol += (long)kasr2.Value;
            //}

            long Abzarncol = 0;
            var Abzar = db.tbReffrenceSave.Where(p => p.FK_PeymanID == py.tbPeymanContracts.pec_ID && p.IsFor == 4).FirstOrDefault();
            if (Abzar != null)
            {
                var Abzar2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == Abzar.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                if (Abzar2 != null)
                {
                    var Abzar3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == Abzar2.ID /*&& p.Final_Accept == true*/ && p.tbEquipmentMoalefeValue.Any(s => s.Month == mon && s.Year == yr)).FirstOrDefault();
                    if (Abzar3 != null)
                    {
                        var Abzar4 = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == Abzar3.ID && p.CountDays != 0 && p.Fk_user == UserID).ToList();
                        foreach (var item10 in Abzar4)
                        {
                            if (item10.CountDays != null)
                            {
                                var rx = db.tbEquipments.Where(p => p.ID == item10.FK_Equipment && p.FK_Peyman == py.tbPeymanContracts.pec_ID).FirstOrDefault();
                                if (rx != null)
                                {
                                    Abzarncol += (long)(item10.CountDays * Math.Round((double)(rx.EachValue ?? 0) * 12 / 366, 2));

                                    //Abzarncol += (long)(item10.CountDays * Math.Round((double)(rx.EachValue ?? 0) * 12 / 365, 2));
                                    //var c55c = Math.Round((double)(rx.EachValue ?? 0) * 12 / 365, 2);
                                    //var abz = item10.CountDays * c55c;
                                }
                            }
                           
                        }
                    }
                }

            }













            if (find3 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find3.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Creditor = (float)mashincol;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find3.ID;
                    FinancialDocumentss.Creditor = (float)mashincol;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره خودرو شخصیاجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5003;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)mashincol;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }
            if (find4 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find4.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Creditor = (float)Abzarncol;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find4.ID;
                    FinancialDocumentss.Creditor = (float)Abzarncol;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره ابزار";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره ابزار";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5004;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)Abzarncol;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره ابزار";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }


            float kosor = (float)((Abzarncol + mashincol + ayabzahab2 + moh)*0.1);

            if (find5 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find5.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Debtore = (float)kosor;
                    find.Creditor = 0;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find5.ID;
                    FinancialDocumentss.Debtore = (float)kosor;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Creditor = 0;
                    FinancialDocumentss.Title = "کسورات قانونی";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "کسورات قانونی";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5005;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Debtore = (float)kosor;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Creditor = 0;
                FinancialDocumentss.Title = "کسورات قانونی";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }






            FishValue valuefish12 = new FishValue
            {
                Title = "تعداد ساعات اضافه کار",
                Value = tedadsaatezafkar/*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish12);
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }
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


            //var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 * tedadroozmah);
            //var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 * tedadroozmah);
            //var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 * tedadroozmah);
            //var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 * tedadroozmah);
            var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 * teadadroozkarkard);
            var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 * teadadroozkarkard);
            var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 * teadadroozkarkard);
            var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 * teadadroozkarkard);
            var shabkari = (mozdmabna * 35 / 100) * (tedadroozshabkari *tedadroozmah);
            var eydivapadash = mozdmabna * 60/365* (teadadroozkarkard+ tedadroozestelaji);
            var sanavatkhdmat = eydivapadash /2;
            //var  malf = db.tbMoalefeValueFish.ToList();
            //foreach (var m in malf)
            //{
            //    var ex = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.FK_Moalefe == 1944).FirstOrDefault();
            //    if (ex != null)
            //    {
            //        ex.mlfvlfsh_Value = eydivapadash;
            //        db.SaveChanges();
            //    }
            //}

            if (contractinfo != null)
            {
                // Result.FishValues = new List<FishValue>();
                FishValue fv = new FishValue

                {
                    Title = "مزد سنوات (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane)
                };

                fishValues.Add(fv);

                FishValue nobatkari = new FishValue

                {
                    Title = "نوبتکاری (ریال)",
                    Value = (double)(nobatkari1 + nobatkari2 + nobatkari3 + nobatkari4 + shabkari)
                };

                fishValues.Add(nobatkari);

                FishValue fv85 = new FishValue

                {
                    Title = "مزد سنوات (ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane) * teadadroozkarkard
                };

                fishValues.Add(fv85);
                FishValue fv2 = new FishValue
                {
                    Title = "مزد شغل (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdshoqlroozane)
                };
                fishValues.Add(fv2);
                FishValue fv200 = new FishValue
                {
                    Title = "مزد گروه (شغل)",
                    Value = Convert.ToDouble(mozdshoqlroozane) * teadadroozkarkard
                };
                fishValues.Add(fv200);
                FishValue fv3 = new FishValue
                {
                    Title = "مزد سایر ( روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsayer)
                };
                fishValues.Add(fv3);
                FishValue fv4 = new FishValue
                {
                    Title = "جمع مزد مبنا ( روزانه-ریال)",
                    Value = mozdmabna
                };
                fishValues.Add(fv4);
                FishValue fv5 = new FishValue
                {
                    Title = "اضافه کار ( هرساعت-ریال)",
                    Value = Math.Round(ezafekar)
                };
                fishValues.Add(fv5);
                FishValue fv6 = new FishValue
                {
                    Title = "کارکرد جمعه (روزانه-ریال)",
                    Value = karkardjome
                };
                fishValues.Add(fv6);
                FishValue fv7 = new FishValue
                {
                    Title = "کارکرد روز تعطیل (روزانه-ریال)",
                    Value = karkarrooztatil
                };
                fishValues.Add(fv7);
                FishValue fv8 = new FishValue
                {
                    Title = "نوبتکاری1 ( صبح وعصر-روزانه-ریال)",
                    Value = (double)nobatkari1
                };
                fishValues.Add(fv8);
                FishValue fv9 = new FishValue
                {
                    Title = "نوبتکاری2 (صبح وشب-روزانه-ریال)",
                    Value = (double)nobatkari2
                };
                fishValues.Add(fv9);
                FishValue fv10 = new FishValue
                {
                    Title = "نوبتکاری3 (عصروشب-روزانه-ریال)",
                    Value = (double)nobatkari3
                };
                fishValues.Add(fv10);
                FishValue fv11 = new FishValue
                {
                    Title = "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)",
                    Value = (double)nobatkari4
                };
                fishValues.Add(fv11);
                FishValue fv12 = new FishValue
                {
                    Title = "شبکاری (هرشب-ریال)",
                    Value = (double)shabkari
                };
                fishValues.Add(fv12);
                FishValue fv13 = new FishValue
                {
                    Title = "عیدی و پاداش ( ماهانه-ریال)",
                    Value = eydivapadash /** nesbatkarbkolmah*/
                };
                fishValues.Add(fv13);
                FishValue fv14 = new FishValue
                {
                    Title = "سنوات خدمت (ماهانه-ریال)",
                    Value = sanavatkhdmat /** nesbatkarbkolmah*/
                };
                fishValues.Add(fv14);
                FishValue fv15 = new FishValue
                {
                    Title = "حق مرخصی (ریال)",
                    Value = haghmorkhasi * nesbatkarbkolmah

                };
                fishValues.Add(fv15);
                FishValue fv16 = new FishValue
                {
                    Title = "کمک هزینه مسکن (ریال)",
                    Value = Convert.ToDouble(haghmaskanbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv16);
                FishValue f17 = new FishValue
                {
                    Title = "کمک هزینه اقلام مصرفی خانوار(ریال)",
                    Value = Convert.ToDouble(aghlammasrafikhanevarbase) * nesbatkarbkolmah
                };
                fishValues.Add(f17);
                FishValue fv18 = new FishValue
                {
                    Title = "کمک هزینه اولاد(ریال)",
                    Value = Convert.ToDouble(hagholadbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv18);

                FishValue fv19 = new FishValue
                {
                    Title = "کسر کار (هرساعت-ریال)",
                    Value = (mozdmabna / 7.3333) * zaribkasrikar
                };
                fishValues.Add(fv19);
                FishValue fv20 = new FishValue
                {
                    Title = "ماموریت ( روزانه-ریال)",
                    Value = (mozdmabna * zaribmamoriat)
                };
                fishValues.Add(fv20);
                FishValue fv21 = new FishValue
                {
                    Title = "غیبت (روزانه-ریال)",
                    Value = (mozdmabna * zaribqeybat)
                };
                fishValues.Add(fv21);


            }
            #endregion
            #region پر کردن هدر فیش

            Result.FishHeader = fishHeader;
            #endregion



            FishValue Mamoriat = new FishValue
            {
                Title = "ماموریت (ریال)",
                Value = CalculateMamoriat((mozdmabna * zaribmamoriat), (double)roozmamoriat)
            };
            fishValues.Add(Mamoriat);
            FishValue ezafkari = new FishValue
            {
                Title = "اضافه کاری (ریال)",
                Value = CalculateEzafkar(ezafekar, tedadsaatezafkar)
            };
            fishValues.Add(ezafkari);

            FishValue jomehkari = new FishValue
            {
                Title = "جمعه کاری (ریال)",
                Value = Calculatejomehkar(karkardjome, teadaroozjome)
            };
            //fishValues.Add(jomehkari);
            var findfkmoalfeh = db.tbContractMoalefeDastmozdi.Where(s => s.md_Title == "جمعه کاری (ریال)").FirstOrDefault();
            if (findfkmoalfeh != null)
            {
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == findfkmoalfeh.md_ID && p.Finalaccept == true);
                if (varObject == null)
                {
                    tbMoalefeValueFish tbMoalefeValueFish12233 = new tbMoalefeValueFish();
                    tbMoalefeValueFish12233.FK_User = UserID;
                    tbMoalefeValueFish12233.mlfvlfsh_Month = Month;
                    tbMoalefeValueFish12233.mlfvlfsh_Year = Year;
                    tbMoalefeValueFish12233.Finalaccept = true;
                    tbMoalefeValueFish12233.FK_Moalefe = findfkmoalfeh.md_ID;

                    // رند کردن مقدار
                    tbMoalefeValueFish12233.mlfvlfsh_Value = Math.Round(jomehkari.Value);

                    db.tbMoalefeValueFish.Add(tbMoalefeValueFish12233);
                    db.SaveChanges();

                }
                else
                {
                    varObject.mlfvlfsh_Value = Math.Round(jomehkari.Value);
                    db.SaveChanges();

                }

            }

            #region پر کردن اضافات و ریختن در فیش ولیو
            var ezafat = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
            List<FishValue> ezafatValue = new List<FishValue>();

            foreach (var item in ezafat)
            {
                //sistan
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID&&p.Finalaccept==true);
               
                    FishValue ezafvalue = new FishValue
                    {
                        Title = item.md_Title,
                        //Value = (double)db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID).mlfvlfsh_Value
                        Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
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
                var var = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID).FirstOrDefault();
                FishValue kasrvalue = new FishValue
                {
                    Title = item.md_Title,
                    Value = var != null ? (double)var.mlfvlfsh_Value : 0
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
                        var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (exist == null)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                            {
                                FK_Moalefe = fkMoalefe,
                                mlfvlfsh_Submit = qatee,
                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishRepo.Create(moalefefish);
                        }
                        else if (exist.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            exist.FK_Moalefe = fkMoalefe;
                            exist.mlfvlfsh_Submit = qatee;
                            exist.FK_User = UserID;
                            exist.mlfvlfsh_Month = Month;
                            exist.mlfvlfsh_Year = Year;
                            exist.mlfvlfsh_Value = item.Value;
                            db.SaveChanges();
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception e)
                {

                    transaction.Rollback();
                }
            }


            #endregion




            #region محاسبه بیمه
            var bimelist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 1 || p.included == 3).ToList();
            double jamebime = 0;
            foreach (var item in bimelist)
            {
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID);

                double valueToAdd = varObject != null ? (double)varObject.mlfvlfsh_Value : 0;
                jamebime += valueToAdd;

            }
            FishValue valuejamebime = new FishValue
            {
                Title = "جمع کل مشمول بیمه (ریال)",
                Value = Math.Round(jamebime)
            };
            fishValues.Add(valuejamebime);

            FishValue bimesahmkarmand = new FishValue
            {
                Title = "حق بیمه سهم کارمند (ریال)",
                Value = Math.Round(CalculateBime(jamebime, Year))
            };
            fishValues.Add(bimesahmkarmand);
            #endregion



            #region محاسبه مالیات

            var eydivapadashsalane = 0;
            int countday1 = 0;
            int startmonth = contractinfo.FirstMonthFish+1??0; //contractinfo.ShamsiStartTime_month;
            var Moalefeeydivapadash = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "عیدی و پاداش ( ماهانه-ریال)");
            int moalefeeydivapadashID = 0;
            if (Moalefeeydivapadash != null)
            {
                moalefeeydivapadashID = Moalefeeydivapadash.md_ID;
                if (Month == 12)
                {
                    if (Month == 12)
                    {
                        var rozaneh = Convert.ToInt32(mozdsanavatroozane);
                        var shoglroza = Convert.ToInt32(mozdshoqlroozane);
                        eydivapadashsalane = (int)shoglroza + rozaneh;


                        for (int i = 1; i <= 11; i++)
                        {



                            var countday = db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == i && s.FK_User == UserID && s.mlfvlfsh_Year == Year && s.FK_Moalefe == 1960).FirstOrDefault();
                            if (countday != null)
                            {
                                countday1 += (int)countday.mlfvlfsh_Value;

                            }
                            var countdaystalagi = db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == i && s.FK_User == UserID && s.mlfvlfsh_Year == Year && s.FK_Moalefe == 1964).FirstOrDefault();
                            if (countdaystalagi != null)
                            {
                                countday1 += (int)countdaystalagi.mlfvlfsh_Value;

                            }


                        }
                        countday1 += teadadroozkarkard;
                        countday1 +=(int) tedadroozestelaji;

                        long countnumber = 0;
                        eydivapadashsalane = eydivapadashsalane * 60;


                        countnumber = (long)eydivapadashsalane / 365;
                        countnumber = countnumber * countday1;

                        eydivapadashsalane = (int)countnumber;


                        //for (int i = 0 /*startmonth*/; i < startmonth/*13*/; i++)
                        //{

                        //    eydivapadashsalane += Convert.ToInt32(db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == moalefeeydivapadashID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == i && p.FK_User == UserID).mlfvlfsh_Value);
                        //}
                    }
                }
            }
            var maliatlist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 2 || p.included == 3).ToList();
            double jammaliat = 0;
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
            jammaliat = jammaliat - ((7 / 7) * bimesahmkarmand.Value) - getBimetakmiliValue(Year, Month, UserID);

            FishValue eydivapadashsaliyaneh = new FishValue
            {
                Title = "عیدی و پاداش (سالیانه - ریال)",
                Value = eydivapadashsalane
            };
            fishValues.Add(eydivapadashsaliyaneh);


            FishValue valuejammaliat = new FishValue
            {
                Title = "جمع کل مشمول مالیات (ریال)",
                Value = Math.Round(jammaliat + eydivapadashsalane)
            };
            fishValues.Add(valuejammaliat);




            FishValue maliatsahmkarmand = new FishValue
            {
                Title = "مالیات سهم کارمند (ریال)",
                Value = CalculateTaxTajamoe(jammaliat, Year, Month, UserID, eydivapadashsalane)
            };
            fishValues.Add(maliatsahmkarmand);
            #endregion

            #region محاسبه اقساط و ریختن در فیش ولیو
            var tamamaqsat = CalculateAqsat(Year, Month, UserID);
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
            //TODO: Felan badan bayad pakshavad
            //FishValue felan = new FishValue
            //{
            //    Title = "صندوق خیریه شرکت",
            //    Value = 200000
            //};
            //aqsatformohasebekasriha.Add(felan);

            Result.FishValuesAqsat = aqsatformohasebekasriha;

            #endregion

            #region قسمت های محاسباتی سمت ویو
            Result.ZakhireKarMazadQabl = CalculateZakhireKarMazadQabl(UserID, Year, Month);
            FishValue ZakhireKarMazadQabl = new FishValue
            {
                Title = "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)",
                Value = Result.ZakhireKarMazadQabl
            };
            fishValues.Add(ZakhireKarMazadQabl);

            Result.JamNakhalesHoqoqVaMazaya = CalculateJamNakhalesHoqoqVaMazaya(fishValues, lstmoalefeqarardadfish, Result.ZakhireKarMazadQabl, ezafatValue, nesbatkarbkolmah, contractinfo.usc_ID);
            Result.JamNakhalesHoqoqVaMazaya += eydivapadashsalane;
           FishValue JamNakhalesHoqoqVaMazaya = new FishValue
            {
                Title = "جمع ناخالص حقوق و مزایا (ریال)",
                Value = Result.JamNakhalesHoqoqVaMazaya
            };
            fishValues.Add(JamNakhalesHoqoqVaMazaya);
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }
            Result.kasrikarkard = Calculatekasrikarkard((int)tedadsaatkasrekar, (mozdmabna / 7.3333) * zaribkasrikar);
            FishValue kasrikarkard = new FishValue
            {
                Title = "کسری کارکرد (ریال)",
                Value = Result.kasrikarkard
            };
            fishValues.Add(kasrikarkard);
            Result.ZakhireKarMazad = CalculateZakhireKarMazad(CalculatejamkolekosooratbdoonMaxpay(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha), Result.JamNakhalesHoqoqVaMazaya, UserID, Year, Month);
            FishValue ZakhireKarMazad = new FishValue
            {
                Title = "ذخیره کار مازاد (ریال)",
                Value = Result.ZakhireKarMazad
            };
            fishValues.Add(ZakhireKarMazad);
            Result.jamkolekosoorat = Calculatejamkolekosoorat(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha, Result.ZakhireKarMazad);
            FishValue jamkolekosoorat = new FishValue
            {
                Title = "جمع کل کسورات (ریال)",
                Value = Result.jamkolekosoorat
            };
            fishValues.Add(jamkolekosoorat);
            Result.KhalesQabelDaryaft = CalculateKhalesQabelDaryaft(Result.jamkolekosoorat, Result.JamNakhalesHoqoqVaMazaya);
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
                        var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (exist == null)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                            {
                                FK_Moalefe = fkMoalefe,
                                mlfvlfsh_Submit = qatee,
                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishRepo.Create(moalefefish);
                        }
                        else if (exist.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            exist.FK_Moalefe = fkMoalefe;
                            exist.mlfvlfsh_Submit = qatee;
                            exist.FK_User = UserID;
                            exist.mlfvlfsh_Month = Month;
                            exist.mlfvlfsh_Year = Year;
                            exist.mlfvlfsh_Value = item.Value;
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



            tbSaleryIsCalculated entity = new tbSaleryIsCalculated
            {
                FKUser = UserID,
                slrcal_Month = Month,
                slrcal_Year = Year
            };
            saleryiscalcRepo.Create(entity);


            return Result;








        }
        public List<Tuple<int, int>> CalculateAqsat(int year, int month, int UserID)
        {
            var installmentList = db.tbinstallments.Where(p => p.User_ID == UserID&&p.tb_Subset_of_installments.Any(s=>s.Month==month&&s.Year==year)).ToList();
            List<Tuple<int, int>> result = new List<Tuple<int, int>>();
            foreach (var item in installmentList)
            {

                var value = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == item.installment_ID && p.Year == year && p.Month == month).Select(b=>b.value).FirstOrDefault();
                var title = Convert.ToInt32(item.Fk_molfe);
                Tuple<int, int> tuple = new Tuple<int, int>(value, title);
                result.Add(tuple);
            }




            return result;
        }

        public int TeadaroozMonth(int Month, int year)
        {
            PersianCalendar pc = new PersianCalendar();

            if (Month < 7)
            {
                return 31;
            }
            else if (Month < 12)
            {

                return 30;
            }
            else
            {
                if (pc.IsLeapYear(year))
                {
                    return 30;
                }
                else
                {
                    return 29;
                }
            }
        }
        public string MonthName(int Month)
        {
            switch (Month)
            {
                case 1:
                    return "فروردین";

                case 2:
                    return "اردیبهشت";

                case 3:
                    return "خرداد";

                case 4:
                    return "تیر";

                case 5:
                    return "مرداد";

                case 6:
                    return "شهریور";

                case 7:
                    return "مهر";

                case 8:
                    return "آبان";

                case 9:
                    return "آذر";

                case 10:
                    return "دی";

                case 11:
                    return "بهمن";

                case 12:
                    return "اسفند";

                default:
                    return "ناشناخته";

            }
        }

        public static DateTime PersianDateToDateTime2(string persianDate)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            string[] parts = persianDate.Split('/');
            int year = int.Parse(parts[2]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[0]);
            return persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        }

        public static DateTime PersianDateToDateTime(string persianDate)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            string[] parts = persianDate.Split('/');
            int year = int.Parse(parts[2]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[0]);
            return persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        }


        public float CalculateTax(double mashmoolmaliyat, int year)

        {

            var liststeptaxes = stepTaxRepository.List().Where(p => p.year == year).OrderBy(p => p.number).ToList();
            float min = 0;
            float max = 0;
            int step = 0;
            float result = 0;
            foreach (var item in liststeptaxes)
            {
                min = (float)item.Fromamount;
                max = (float)item.Toamount;
                if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                {
                    step = (int)item.number;
                    break;
                }
            }
            for (int i = 1; i < step; i++)
            {
                result = (float)((liststeptaxes[i - 1].Toamount - liststeptaxes[i - 1].Fromamount) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
            }
            result = (float)((mashmoolmaliyat - liststeptaxes[step - 1].Fromamount) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;
            return result;


        }

        public double CalculateBime(double mashmoolbime, int year)
        {
            return mashmoolbime * 0.07;
        }

        public double CalculateJamNakhalesHoqoqVaMazaya(List<FishValue> fshvlu, List<FishValue> lstmoalefeqaradad, double ZakhireKarMazadQabl, List<FishValue> ezafat,double nesbatkarbkolmah,int usc_ID)//TODO:badan inro bayad faqat ezafat bashe
        {
            var x = fshvlu.FirstOrDefault(p => p.Title == "مزد گروه (شغل)");
            double mozdshoqlroozane = 0;
            if (x != null)
            {
                mozdshoqlroozane = x.Value;
            }

            var y = fshvlu.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)");
            double mozdsanavat = 0;
            if (y != null)
            {
                mozdsanavat = y.Value;
            }
            var z = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)");
            double komakhazinemaskan = 0;
            if (z != null)
            {
                komakhazinemaskan = z.Value;
            }

            var xx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)");
            double aghlammasrafi = 0;
            if (xx != null)
            {
                aghlammasrafi = xx.Value;
            }

            var xy = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)");
            double komakhazineolad = 0;
            if (xy != null)
            {
                komakhazineolad = xy.Value;
            }

            var xz = fshvlu.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)");
            double ezafekari = 0;
            if (xz != null)
            {
                ezafekari = xz.Value;
            }


            var xxx = fshvlu.FirstOrDefault(p => p.Title == "ماموریت (ریال)");
            double mamoriat = 0;
            if (xxx != null)
            {
                mamoriat = xxx.Value;
            }


            var xxy = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)");
            double eslahiyemaskan = 0;
            if (xxy != null)
            {
                eslahiyemaskan = xxy.Value;
            }


            var xxz = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)");
            double bimetakmildarman = 0;
            if (xxz != null)
            {
                bimetakmildarman = xxz.Value;
            }
            var yxx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)");
            double ayabozahab = 0;

            if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null /*&& db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null*/)
            {
                if (yxx != null)
                {
                    ayabozahab = yxx.Value;
                    //ayabozahab = 0;

                }
            }
            else
            {
                if (yxx != null)
                {
                    //ayabozahab = yxx.Value;
                    ayabozahab = 0;

                }
            }

            var yxy = fshvlu.FirstOrDefault(p => p.Title == "پاداش(ریال)");
            double padash = 0;
            if (yxy != null)
            {
                padash = yxy.Value;
            }


            var yxz = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)");
            double eslahiyeolad = 0;
            if (yxz != null)
            {
                eslahiyeolad = yxz.Value;
            }

            var yyx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)");
            double abzarkar = 0;
            if (yyx != null)
            {
                //abzarkar = yyx.Value;
                abzarkar = 0;
            }


            var yyz = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه تبلت و رایانه(ریال)");
            double tablet = 0;
            if (yyz != null)
            {
                //tablet = yyz.Value;
        tablet = 0;

    }






    var zxx = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)");
            double eslahiyekhanevar = 0;
            if (zxx != null)
            {
                eslahiyekhanevar = zxx.Value;
            }
            var zxy = fshvlu.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)");
            double nobatkari = 0;
            if (zxy != null)
            {
                nobatkari = zxy.Value;
            }
            double result = Convert.ToInt32(Math.Round(mozdshoqlroozane + mozdsanavat + komakhazinemaskan + aghlammasrafi + komakhazineolad + ezafekari + mamoriat + eslahiyemaskan + bimetakmildarman + ayabozahab + padash + eslahiyeolad + abzarkar + tablet + eslahiyekhanevar + nobatkari)) + ZakhireKarMazadQabl;
            foreach (var item in lstmoalefeqaradad)
            {
                result += Convert.ToInt32(item.Value) * nesbatkarbkolmah;
            }
            foreach (var item in ezafat)
            {
                var tt=db.tbUserContractsAndMoalefeGhararDadi.Where(p=> p.FKContractID == usc_ID&&p.tbContractMoalefeDastmozdi.md_Title.Contains(item.Title)).FirstOrDefault();
                if (tt == null)
                {
                    result += Convert.ToInt32(item.Value);

                }
            }
            return result;
        }
        public double CalculateJamNakhalesHoqoqVaMazayaFishTest(List<FishValue> fshvlu, List<FishValue> lstmoalefeqaradad, double ZakhireKarMazadQabl, double ezafat)
        {
            double x = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "مزد گروه (شغل)") != null)
            {
                x = fshvlu.FirstOrDefault(p => p.Title == "مزد گروه (شغل)").Value;
            }
            double y = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)") != null)
            {
                y = fshvlu.FirstOrDefault(p => p.Title == "مزد سنوات (ریال)").Value;
            }

            double xx = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)") != null)
            {
                xx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه مسکن (ریال)").Value;
            }
            double xy = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)") != null)
            {
                xy = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اقلام مصرفی خانوار(ریال)").Value;
            }

            double yx = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)") != null)
            {
                yx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه اولاد(ریال)").Value;
            }

            double yz = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)") != null)
            {
                yz = fshvlu.FirstOrDefault(p => p.Title == "اضافه کاری (ریال)").Value;
            }
            double zx = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "ماموریت (ریال)") != null)
            {
                zx = fshvlu.FirstOrDefault(p => p.Title == "ماموریت (ریال)").Value;
            }
            double zy = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)") != null)
            {
                zy = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه مسکن(ریال)").Value;
            }
            double zz = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)") != null)
            {
                zz = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه بیمه تکمیل درمان (ریال)").Value;
            }
            double xxx = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
            {
                xxx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ایاب و ذهاب(ریال)").Value;
            }
            double xxy = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "پاداش(ریال)") != null)
            {
                xxy = fshvlu.FirstOrDefault(p => p.Title == "پاداش(ریال)").Value;
            }
            double xxz = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)") != null)
            {
                xxz = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اولاد(ریال)").Value;
            }
            double xyx = 0;
            if (fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)") != null)
            {
                xyx = fshvlu.FirstOrDefault(p => p.Title == "کمک هزینه ابزار کار(ریال)").Value;
            }
            var zxx = fshvlu.FirstOrDefault(p => p.Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)");
            double eslahiyekhanevar = 0;
            if (zxx != null)
            {
                eslahiyekhanevar = zxx.Value;
            }

            var zxy = fshvlu.FirstOrDefault(p => p.Title == "نوبتکاری (ریال)");
            double nobatkari = 0;
            if (zxy != null)
            {
                nobatkari = zxy.Value;
            }
            double result = Convert.ToInt32(Math.Round(x + y + xx + xy + yx + yz + zx + zy + zz + xxx + xxy + xxz + xyx + eslahiyekhanevar) + ZakhireKarMazadQabl + ezafat);
            foreach (var item in lstmoalefeqaradad)
            {
                result += Convert.ToInt32(item.Value);
            }
            return result;
        }
        public int Calculatekasrikarkard(double tedadsaatkasrkar, double mablaqpayekasrekar)
        {
            int result = Convert.ToInt32(tedadsaatkasrkar * mablaqpayekasrekar);/* Convert.ToInt32(fshvlu.FirstOrDefault(p => p.Title == "کسر کار (هرساعت-ریال)").Value * fshvlu.FirstOrDefault(p => p.Title == "تعداد ساعات کسر کار").Value)*/;
            return result;
        }

        public double Calculatejamkolekosoorat(List<FishValue> fshvlu, List<FishValue> lstFishValuesKosoorat, double kasrkarkard, List<FishValue> aqsat, double zakhirekarmazad)
        {
            double result = Convert.ToInt32(fshvlu.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value + fshvlu.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value) + kasrkarkard + zakhirekarmazad;
            foreach (var item in lstFishValuesKosoorat)
            {
                result += Convert.ToInt32(item.Value);
            }
            foreach (var item in aqsat)
            {
                result += Convert.ToInt32(item.Value);
            }
            return result;
        }
        public double CalculatejamkolekosooratbdoonMaxpay(List<FishValue> fshvlu, List<FishValue> lstFishValuesKosoorat, double kasrkarkard, List<FishValue> aqsat)
        {
            double result = Convert.ToInt32(fshvlu.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value + fshvlu.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value) + kasrkarkard;
            foreach (var item in lstFishValuesKosoorat)
            {
                result += Convert.ToInt32(item.Value);
            }
            foreach (var item in aqsat)
            {
                result += Convert.ToInt32(item.Value);
            }
            return result;
        }
        public double CalculateZakhireKarMazad(double jamkosooratbdonmaxpay, double jamnakhalesHoqoqvamazaya, int userID, int year, int month)
        {
            double result = 0;
            var x = db.tbMaxPayForMonth.FirstOrDefault(p => p.FKUser == userID && p.Year == year && p.Month == month);
            if (x != null)
            {
                var maxpay = x.Value;

                result = jamnakhalesHoqoqvamazaya - jamkosooratbdonmaxpay - maxpay;
                if(result<0 ){
                    result = 0;
                }

            }
            FillRemainValueFortbMaxPayForMonth(userID, year, month, result);
            return result;
        }
        public double CalculateKhalesQabelDaryaft(double Jamkolekosorat, double JamNakhalesHoqoqVaMazaya)
        {
            double result = JamNakhalesHoqoqVaMazaya - Jamkolekosorat;
            return result;
        }

        public int CalculateZakhireKarMazadQabl(int UserID, int Year, int Month)
        {
            int result = 0;
            if (Month != 1)
            {
                var x = db.tbMaxPayForMonth.FirstOrDefault(p => p.FKUser == UserID && p.Year == Year && p.Month == Month - 1);
                if (x != null)
                {

                    result = Convert.ToInt32(x.RemainValue);
                }
            }
            else
            {
                var x = db.tbMaxPayForMonth.FirstOrDefault(p => p.FKUser == UserID && p.Year == Year - 1 && p.Month == 12);
                if (x != null)
                {
                    result = Convert.ToInt32(x.RemainValue);
                }

            }
            return result;
        }

        public void FillRemainValueFortbMaxPayForMonth(int UserID, int Year, int Month, double Zakhirekarmazad)
        {
            var tbMaxpay = db.tbMaxPayForMonth.FirstOrDefault(p => p.FKUser == UserID && p.Year == Year && p.Month == Month);
            if (tbMaxpay != null)
            {
                tbMaxpay.RemainValue = Convert.ToInt32(Zakhirekarmazad);

                db.SaveChanges();
            }

        }
        #endregion
        #region فیش آزمایشی
        public ManualFishDetail CalculateandinserttotableTest(int UserID, bool ISRoozMozd, RoozMozdParam roozMozdParam, KarMozdParam karMozdParam, int Year, int Month, FishHeader fishHeader)
        {
            List<FishValue> fishValues = new List<FishValue>();
            ManualFishDetail Result = new ManualFishDetail();
            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);
            var mozdsayer = "0";
            var haghmaskanbase = "0";
            var aghlammasrafikhanevarbase = "0";
            var hagholadbase = "0";
            float zaribkasrikar = 1;
            float zaribmamoriat = 1;
            float zaribqeybat = 1;

            double tedadroozkarkard = 0;
            double tedadroozestelaji = 0;
            double tedadsaatezafkar = 0;
            double tedadroozmamoriat = 0;
            double tedadsaatkasrekar = 0;
            double jamezafat = 0;
            double jamkasrekar = 0;
            double MoadelKarkard = 0;
            double tedadroozQeybat = 0;
            double tedadrooztatil = 0;
            double teadaroozjome = 0;
            double qestbimetakmili = 0;
            double ayabozahab = 0;

            var tedadroozmah = TeadaroozMonth(Month, Year);
            if (ISRoozMozd)
            {
                tedadroozkarkard = roozMozdParam.tedadrooz;
                tedadroozestelaji = roozMozdParam.tedadroozStelaji;
                tedadsaatezafkar = roozMozdParam.tedadsaatEzafeKar;
                tedadroozmamoriat = roozMozdParam.tedadroozMamoriat;
                tedadsaatkasrekar = roozMozdParam.teadadsaatKasrekar;
                jamezafat = roozMozdParam.jamezafat;
                jamkasrekar = roozMozdParam.jambedehiha;
                tedadroozQeybat = roozMozdParam.tedadroozQeybat;
                qestbimetakmili = roozMozdParam.qestbimetakmili;
                ayabozahab = roozMozdParam.ayabozahab;
            }
            else
            {
                tedadroozestelaji = karMozdParam.tedadroozStelaji;

                var res = Calculatesaatkasrkarvaezafkarforkarmozdi(karMozdParam.tedadroozkarkard, karMozdParam.MoadelKarkard, tedadroozmah, tedadroozestelaji);
                tedadroozkarkard = karMozdParam.tedadroozkarkard;
                tedadsaatezafkar = res.Item1;
                tedadroozmamoriat = karMozdParam.tedadroozMamoriat;
                tedadsaatkasrekar = res.Item2;
                jamezafat = karMozdParam.jamezafat;
                jamkasrekar = karMozdParam.jambedehiha;
                MoadelKarkard = karMozdParam.MoadelKarkard;
                tedadroozQeybat = karMozdParam.tedadroozQeybat;
                tedadrooztatil = karMozdParam.tedadrooztatil;
                teadaroozjome = karMozdParam.teadaroozjome;
                qestbimetakmili = karMozdParam.qestbimetakmili;
                ayabozahab = karMozdParam.ayabozahab;
            }
            var PeymanID = db.Link_User_And_Peyman.FirstOrDefault(p => p.FK_User_ID == UserID && p.Status == true).FK_Peyman_ID;
            var zaribpeyman = db.tbPeymanZaribForFish.FirstOrDefault(p => p.FKPeyman == PeymanID);

            if (zaribpeyman != null)
            {
                zaribkasrikar = (float)zaribpeyman.pymnzrbfsh_zaribkasrekar;
                zaribqeybat = (float)zaribpeyman.pymnzrbfsh_zaribqeybat;
                zaribmamoriat = (float)zaribpeyman.pymnzrbfsh_zaribmamoriat;

            }
            
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
            if (contractinfo.jobgroup_HagheMaskan != "")
            {
                haghmaskanbase = Regex.Replace(contractinfo.jobgroup_HagheMaskan, ",", "");
                if (haghmaskanbase == null)
                {
                    haghmaskanbase = "0";
                }

            }
            if (contractinfo.jobgroup_HagheOlad != "")
            {

                hagholadbase = Regex.Replace(contractinfo.jobgroup_HagheOlad, ",", "");
                if (hagholadbase == null)
                {
                    hagholadbase = "0";
                }
            }
            if (contractinfo.jobgroup_KharoBar != "")
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
            var haghmorkhasi = Math.Round(((((Convert.ToDouble(haghmaskanbase) + Convert.ToDouble(hagholadbase) + Convert.ToDouble(aghlammasrafikhanevarbase)) / 30) + mozdmabna) * 9) / 12);
            var ezafekar = (mozdmabna / 7.3333) * 1.4;
            var karkardjome = ezafekar * 1.4;
            var karkarrooztatil = ezafekar;



            //TODO:این قسمت از کارکرد مامور مشخص خواهد شد



            var moalefehayeexcel = db.tbMoalefeValeFishTest.Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID).ToList();

            //int teadadroozkarkard = teadadroozkarkardforfishmain((int)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت").mlfvlfsh_Value, tedadroozmah, (int)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی").mlfvlfsh_Value);

            FishValue valuefish11 = new FishValue
            {
                Title = "تعداد روز کارکرد",
                Value = tedadroozkarkard /*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز کارکرد").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish11);



            var nesbatkarbkolmah = Calculatekarkardtaghsimbarkolemah(tedadroozkarkard, TeadaroozMonth(Month, Year));

            double? tedadrooznobatkari1 = 0;
            double? tedadrooznobatkari2 = 0;
            double? tedadrooznobatkari3 = 0;
            double? tedadrooznobatkari4 = 0;
            double? tedadroozshabkari = 0;

            var x = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری1 ( صبح‏وعصر)");
            if (x == null)
            {
                tedadrooznobatkari1 = 0;
            }
            else
            {
                tedadrooznobatkari1 = x.mlfvlfsh_Value;
            }

            var y = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری2 (صبح‏وشب)");
            if (y == null)
            {
                tedadrooznobatkari2 = 0;
            }
            else
            {
                tedadrooznobatkari2 = y.mlfvlfsh_Value;
            }
            var z = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری3 (عصروشب)");
            if (z == null)
            {
                tedadrooznobatkari3 = 0;
            }
            else
            {
                tedadrooznobatkari3 = z.mlfvlfsh_Value;
            }
            var xx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری4 (صبح‏وعصروشب)");
            if (xx == null)
            {
                tedadrooznobatkari4 = 0;
            }
            else
            {
                tedadrooznobatkari4 = xx.mlfvlfsh_Value;
            }
            var xy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد شب های کاری");
            if (xy == null)
            {
                tedadroozshabkari = 0;
            }
            else
            {
                tedadroozshabkari = xy.mlfvlfsh_Value;
            }


            double? xz = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)") != null)
            {
                xz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)").mlfvlfsh_Value;
            }

            FishValue valuefish = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اولاد(ریال)",
                Value = (double)xz
            };
            fishValues.Add(valuefish);


            double? yx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)") != null)
            {
                yx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish2 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه مسکن(ریال)",
                Value = (double)yx
            };
            fishValues.Add(valuefish2);




            double? yy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)") != null)
            {
                yy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish3 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)",
                Value = (double)yy
            };
            fishValues.Add(valuefish3);



            double? zx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
            {
                zx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
            }
            if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",

                    Value = ayabozahab * (double)zx //felan az view migirim
                    //Value = 0
                };
                fishValues.Add(valuefish4);
            }
            else
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",

                    //Value = ayabozahab * (double)zx //felan az view migirim
                    Value = 0
                };
                fishValues.Add(valuefish4);
            }


            double? zy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)") != null)
            {
                zy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish5 = new FishValue
            {
                Title = "پاداش(ریال)",
                Value = (double)zy
            };
            fishValues.Add(valuefish5);


            double? zz = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)") != null)
            {
                zz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish6 = new FishValue
            {
                Title = "جریمه(ریال)",
                Value = (double)zz
            };
            fishValues.Add(valuefish6);



            double? xxx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)") != null)
            {
                xxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish7 = new FishValue
            {
                Title = "کمک هزینه ابزار کار(ریال)",
                //Value = (double)xxx
                Value = 0

            };
            fishValues.Add(valuefish7);


            double? xxy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه") != null)
            {
                xxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه").mlfvlfsh_Value;
            }
            FishValue valuefish8 = new FishValue
            {
                Title = "کمک هزینه تبلت و رایانه(ریال)",
                /*Value = (double)xxy*/ /** MoadelKarkard*/
                Value=0
            };
            fishValues.Add(valuefish8);


            double? xxz = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)") != null)
            {
                xxz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish9 = new FishValue
            {
                Title = "کمک هزینه بیمه تکمیل درمان (ریال)",
                Value = (double)xxz
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
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }
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
            fishValues.Add(valuefish18);
            #endregion
            #region سایر مولفه ها


            //var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 / tedadroozmah);
            //var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 / tedadroozmah);
            //var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 / tedadroozmah);
            //var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 / tedadroozmah);
            var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 * tedadroozkarkard);
            var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 * tedadroozkarkard);
            var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 * tedadroozkarkard);
            var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 * tedadroozkarkard);
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



                FishValue nobatkari = new FishValue

                {
                    Title = "نوبتکاری (ریال)",
                    Value = (double)(nobatkari1 + nobatkari2 + nobatkari3 + nobatkari4 + shabkari)
                };
                fishValues.Add(nobatkari);


                FishValue fv85 = new FishValue

                {
                    Title = "مزد سنوات (ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane) * tedadroozkarkard
                };

                fishValues.Add(fv85);

                FishValue fv2 = new FishValue
                {
                    Title = "مزد شغل (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdshoqlroozane)
                };
                fishValues.Add(fv2);
                FishValue fv200 = new FishValue
                {
                    Title = "مزد گروه (شغل)",
                    Value = Convert.ToDouble(mozdshoqlroozane) * tedadroozkarkard
                };
                fishValues.Add(fv200);
                FishValue fv3 = new FishValue
                {
                    Title = "مزد سایر ( روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsayer)
                };
                fishValues.Add(fv3);
                FishValue fv4 = new FishValue
                {
                    Title = "جمع مزد مبنا ( روزانه-ریال)",
                    Value = mozdmabna
                };
                fishValues.Add(fv4);
                FishValue fv5 = new FishValue
                {
                    Title = "اضافه کار ( هرساعت-ریال)",
                    Value = Math.Round(ezafekar)
                };
                fishValues.Add(fv5);
                FishValue fv6 = new FishValue
                {
                    Title = "کارکرد جمعه (روزانه-ریال)",
                    Value = karkardjome
                };
                fishValues.Add(fv6);
                FishValue fv7 = new FishValue
                {
                    Title = "کارکرد روز تعطیل (روزانه-ریال)",
                    Value = karkarrooztatil
                };
                fishValues.Add(fv7);
                FishValue fv8 = new FishValue
                {
                    Title = "نوبتکاری1 ( صبح وعصر-روزانه-ریال)",
                    Value = (double)nobatkari1
                };
                fishValues.Add(fv8);
                FishValue fv9 = new FishValue
                {
                    Title = "نوبتکاری2 (صبح وشب-روزانه-ریال)",
                    Value = (double)nobatkari2
                };
                fishValues.Add(fv9);
                FishValue fv10 = new FishValue
                {
                    Title = "نوبتکاری3 (عصروشب-روزانه-ریال)",
                    Value = (double)nobatkari3
                };
                fishValues.Add(fv10);
                FishValue fv11 = new FishValue
                {
                    Title = "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)",
                    Value = (double)nobatkari4
                };
                fishValues.Add(fv11);
                FishValue fv12 = new FishValue
                {
                    Title = "شبکاری (هرشب-ریال)",
                    Value = (double)shabkari
                };
                fishValues.Add(fv12);
                FishValue fv13 = new FishValue
                {
                    Title = "عیدی و پاداش ( ماهانه-ریال)",
                    Value = eydivapadash * nesbatkarbkolmah
                };
                fishValues.Add(fv13);
                FishValue fv14 = new FishValue
                {
                    Title = "سنوات خدمت (ماهانه-ریال)",
                    Value = 0//sanavatkhdmat * tedadroozkarkard
                };
                fishValues.Add(fv14);
                FishValue fv15 = new FishValue
                {
                    Title = "حق مرخصی (ریال)",
                    Value = haghmorkhasi * nesbatkarbkolmah

                };
                fishValues.Add(fv15);
                FishValue fv16 = new FishValue
                {
                    Title = "کمک هزینه مسکن (ریال)",
                    Value = Convert.ToDouble(haghmaskanbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv16);
                FishValue f17 = new FishValue
                {
                    Title = "کمک هزینه اقلام مصرفی خانوار(ریال)",
                    Value = Convert.ToDouble(aghlammasrafikhanevarbase) * nesbatkarbkolmah
                };
                fishValues.Add(f17);
                FishValue fv18 = new FishValue
                {
                    Title = "کمک هزینه اولاد(ریال)",
                    Value = Convert.ToDouble(hagholadbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv18);

                FishValue fv19 = new FishValue
                {
                    Title = "کسر کار (هرساعت-ریال)",
                    Value = (mozdmabna / 7.3333) * zaribkasrikar
                };
                fishValues.Add(fv19);
                FishValue fv20 = new FishValue
                {
                    Title = "ماموریت ( روزانه-ریال)",
                    Value = (mozdmabna * zaribmamoriat)
                };
                fishValues.Add(fv20);
                FishValue fv21 = new FishValue
                {
                    Title = "غیبت (روزانه-ریال)",
                    Value = (mozdmabna * zaribqeybat)
                };
                fishValues.Add(fv21);


            }
            #endregion
            #region پر کردن هدر فیش

            Result.FishHeader = fishHeader;
            #endregion


            FishValue Mamoriat = new FishValue
            {
                Title = "ماموریت (ریال)",
                Value = CalculateMamoriat((mozdmabna * zaribmamoriat), tedadroozmamoriat)
            };
            fishValues.Add(Mamoriat);
            FishValue ezafkari = new FishValue
            {
                Title = "اضافه کاری (ریال)",
                Value = Math.Round(CalculateEzafkar(ezafekar, tedadsaatezafkar))
            };
            fishValues.Add(ezafkari);
            FishValue jomehkari = new FishValue
            {
                Title = "جمعه کاری (ریال)",
                Value = Calculatejomehkar(karkardjome, teadaroozjome)
            };
            //fishValues.Add(jomehkari);
            var findfkmoalfeh = db.tbContractMoalefeDastmozdi.Where(s => s.md_Title == "جمعه کاری (ریال)").FirstOrDefault();
            if (findfkmoalfeh != null)
            {
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == findfkmoalfeh.md_ID && p.Finalaccept == true);
                if (varObject == null)
                {
                    tbMoalefeValueFish tbMoalefeValueFish12233 = new tbMoalefeValueFish();
                    tbMoalefeValueFish12233.FK_User = UserID;
                    tbMoalefeValueFish12233.mlfvlfsh_Month = Month;
                    tbMoalefeValueFish12233.mlfvlfsh_Year = Year;
                    tbMoalefeValueFish12233.Finalaccept = true;
                    tbMoalefeValueFish12233.FK_Moalefe = findfkmoalfeh.md_ID;

                    // رند کردن مقدار
                    tbMoalefeValueFish12233.mlfvlfsh_Value = Math.Round(jomehkari.Value);

                    db.tbMoalefeValueFish.Add(tbMoalefeValueFish12233);
                    db.SaveChanges();

                }
                else
                {
                    varObject.mlfvlfsh_Value = Math.Round(jomehkari.Value);
                    db.SaveChanges();

                }

            }

            #region پر کردن جدول فیش ولیو برای محاسبه بیمه و مالیات
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in fishValues)
                    {
                        db.Database.CommandTimeout = 400;

                        var fkMoalefe = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == item.Title).md_ID;
                        var exist = db.tbMoalefeValeFishTest.Any(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (!exist)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValeFishTest moalefefish = new tbMoalefeValeFishTest
                            {
                                FK_Moalefe = fkMoalefe,

                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishTestRepo.Create(moalefefish);
                        }
                        else//اگر داشت ویرایش کن
                        {
                            var entity = db.tbMoalefeValeFishTest.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                            entity.FK_Moalefe = fkMoalefe;
                            entity.FK_User = UserID;
                            entity.mlfvlfsh_Month = Month;
                            entity.mlfvlfsh_Year = Year;
                            entity.mlfvlfsh_Value = item.Value;
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
            var bimelist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 1 || p.included == 3).ToList();
            double jamebime = 0;
            foreach (var item in bimelist)
            {

                var varObject = db.tbMoalefeValeFishTest.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID);

                if (varObject != null)
                {
                    jamebime += (double)varObject.mlfvlfsh_Value;
                }
                else
                {
                    jamebime += 0;

                }
            }
            FishValue valuejamebime = new FishValue
            {
                Title = "جمع کل مشمول بیمه (ریال)",
                Value = jamebime
            };
            fishValues.Add(valuejamebime);
            FishValue bimesahmkarmand = new FishValue
            {
                Title = "حق بیمه سهم کارمند (ریال)",
                Value = Math.Round(CalculateBime(jamebime, Year))
            };
            fishValues.Add(bimesahmkarmand);


            #endregion

            #region محاسبه مالیات
            var maliatlist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 2 || p.included == 3).ToList();
            double jammaliat = 0;
            foreach (var item in maliatlist)
            {

                var varObject = db.tbMoalefeValeFishTest.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID);

                if (varObject != null)
                {
                    jammaliat += (double)varObject.mlfvlfsh_Value;
                }
                else
                {
                    jammaliat += 0;
                }
            }
            jammaliat += jamezafat;
            jammaliat = jammaliat - (((7 / 7) * bimesahmkarmand.Value) + qestbimetakmili); /*getBimetakmiliValue(Year, Month, UserID) */ // felan az view miyad
            FishValue valuejammaliat = new FishValue
            {
                Title = "جمع کل مشمول مالیات (ریال)",
                Value = Math.Round(jammaliat)
            };
            fishValues.Add(valuejammaliat);

            FishValue eydivapadashsaliyaneh = new FishValue
            {
                Title = "عیدی و پاداش (سالیانه - ریال)",
                Value = 0
            };
            fishValues.Add(eydivapadashsaliyaneh);

            FishValue maliatsahmkarmand = new FishValue
            {
                Title = "مالیات سهم کارمند (ریال)",
                Value = CalculateTax(jammaliat, Year)
            };
            fishValues.Add(maliatsahmkarmand);

            #endregion

            #region قسمت های محاسباتی سمت ویو
            Result.ZakhireKarMazadQabl = CalculateZakhireKarMazadQabl(UserID, Year, Month);
            FishValue ZakhireKarMazadQabl = new FishValue
            {
                Title = "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)",
                Value = Result.ZakhireKarMazadQabl
            };
            fishValues.Add(ZakhireKarMazadQabl);

            Result.JamNakhalesHoqoqVaMazaya = CalculateJamNakhalesHoqoqVaMazayaFishTest(fishValues, lstmoalefeqarardadfish, Result.ZakhireKarMazadQabl, jamezafat);

            FishValue JamNakhalesHoqoqVaMazaya = new FishValue
            {
                Title = "جمع ناخالص حقوق و مزایا (ریال)",
                Value = Result.JamNakhalesHoqoqVaMazaya
            };
            fishValues.Add(JamNakhalesHoqoqVaMazaya);
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }

            Result.kasrikarkard = Calculatekasrikarkard((double)tedadsaatkasrekar, (mozdmabna / 7.3333) * zaribkasrikar);
            FishValue kasrikarkard = new FishValue
            {
                Title = "کسری کارکرد (ریال)",
                Value = Result.kasrikarkard
            };
            fishValues.Add(kasrikarkard);
            Result.ZakhireKarMazad = CalculateZakhireKarMazad(CalculatejamkolekosooratbdoonMaxpayFishTest(fishValues, jamkasrekar, Result.kasrikarkard), Result.JamNakhalesHoqoqVaMazaya, UserID, Year, Month);
            FishValue ZakhireKarMazad = new FishValue
            {
                Title = "ذخیره کار مازاد (ریال)",
                Value = Result.ZakhireKarMazad
            };
            fishValues.Add(ZakhireKarMazad);
            Result.jamkolekosoorat = CalculatejamkolekosooratFishTest(fishValues, jamkasrekar, Result.kasrikarkard, Result.ZakhireKarMazad, qestbimetakmili);
            FishValue jamkolekosoorat = new FishValue
            {
                Title = "جمع کل کسورات (ریال)",
                Value = Result.jamkolekosoorat
            };
            fishValues.Add(jamkolekosoorat);
            Result.KhalesQabelDaryaft = CalculateKhalesQabelDaryaft(Result.jamkolekosoorat, Result.JamNakhalesHoqoqVaMazaya);
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
                        var exist = db.tbMoalefeValeFishTest.Any(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                        if (!exist)//اگر نداشت اضافه کن
                        {
                            tbMoalefeValeFishTest moalefefish = new tbMoalefeValeFishTest
                            {
                                FK_Moalefe = fkMoalefe,

                                FK_User = UserID,
                                mlfvlfsh_Month = Month,
                                mlfvlfsh_Year = Year,
                                mlfvlfsh_Value = item.Value

                            };
                            moalefeValueFishTestRepo.Create(moalefefish);
                        }
                        else//اگر داشت ویرایش کن
                        {
                            var entity = db.tbMoalefeValeFishTest.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID);
                            entity.FK_Moalefe = fkMoalefe;
                            entity.FK_User = UserID;
                            entity.mlfvlfsh_Month = Month;
                            entity.mlfvlfsh_Year = Year;
                            entity.mlfvlfsh_Value = item.Value;
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



            Result.sayerkosoorat = jamkasrekar;
            Result.SayerEzafat = jamezafat;
            Result.BimetakmiliForFishTest = qestbimetakmili;
            return Result;





        }

        #endregion

        public double Calculatekarkardtaghsimbarkolemah(double tedadroozkarkard, double tedadroozmah)
        {
            double result = tedadroozkarkard / tedadroozmah;
            return result;
        }


        public int teadadroozkarkardforfishmain(int qeybat, int tedadroozmah, int tedadstelaji)
        {
            return tedadroozmah - (qeybat + tedadstelaji);
        }
        public int teadadroozkarkardforfishmainasly(int qeybat, int tedadroozmah, int tedadstelaji,double moadelkarkar)
        {
            double moadel = Math.Ceiling((moadelkarkar * tedadroozmah)+ tedadstelaji);
            if (moadel> tedadroozmah)
            {
                moadel = tedadroozmah;
            }
            
            return (int)moadel - (qeybat + tedadstelaji);
        }


        public Tuple<double, double> Calculatesaatkasrkarvaezafkarforkarmozdi(double tedadroozkarkard, double moadelkarkard, double tedadroozmah, double tedadroozestelaji)
        {
            var tedadsaatkarkard = 162 /*Convert.ToInt32(tedadroozkarkard * 7.3333)*/;
            double ezafkar = 0;

            double kasrkar = 0;
            if (moadelkarkard+ (tedadroozestelaji / tedadroozmah) > 1)
            {
                ezafkar = (moadelkarkard + (tedadroozestelaji / tedadroozmah) - (tedadroozmah / tedadroozmah)) * tedadsaatkarkard;
            }
           
            else if (moadelkarkard < 1)
            {
                //kasrkar = ((tedadroozkarkard / tedadroozmah) - moadelkarkard) * tedadsaatkarkard;
                kasrkar =0;

            }

          


                return System.Tuple.Create(ezafkar, kasrkar);
        }


        public double CalculatejamkolekosooratFishTest(List<FishValue> fshvlu, double jambedehiha, double kasrkarkard, double zakhirekarmazad, double qestbimetakmili)
        {
            double result = Convert.ToInt32(fshvlu.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value + fshvlu.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value) + kasrkarkard + zakhirekarmazad + jambedehiha + qestbimetakmili;


            return result;
        }
        public double CalculatejamkolekosooratbdoonMaxpayFishTest(List<FishValue> fshvlu, double jambedehiha, double kasrkarkard)
        {
            double result = Convert.ToInt32(fshvlu.FirstOrDefault(p => p.Title == "حق بیمه سهم کارمند (ریال)").Value + fshvlu.FirstOrDefault(p => p.Title == "مالیات سهم کارمند (ریال)").Value) + kasrkarkard + jambedehiha;


            return result;
        }


        public double CalculateMamoriat(double valuemamoriat, double tedadroozmamoriat)

        {
            return valuemamoriat * tedadroozmamoriat;
        }
        public double CalculateEzafkar(double valueezafkar, double tedadsaatezafkar)
        {
            return valueezafkar * tedadsaatezafkar;
        }
        public double Calculatejomehkar(double valueezafkar, double tedadsaatezafkar)
        {
            return valueezafkar * tedadsaatezafkar;
        }

        public double getBimetakmiliValue(int year, int month, int User)
        {
            double result = 0;
            try
            {

                var MoalefeID = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "بیمه تکمیلی(ریال)").md_ID;
                var x = db.tbinstallments.FirstOrDefault(p => p.User_ID == User && p.Fk_molfe == MoalefeID&&p.tb_Subset_of_installments.Any(s=>s.Month==month&&s.Year==year));
                if (x != null)
                {
                    var installmentID = x.installment_ID;
                    result = db.tb_Subset_of_installments.Where(p => p.Year == year && p.Month == month && p.fk_installment_ID == installmentID).Select(s=>s.value).FirstOrDefault();

                }
                return result;
            }
            catch (Exception)
            {

                return result;
            }

        }





        public double CalculateMoadelKarkard(int UserID, int month, int year)
        {
            var find = db.tbMoadelPadashJarimeAyab.Where(p => p.UserID == UserID && p.Month == month && p.Year == year).FirstOrDefault();
            if (find != null){
                return db.tbMoadelPadashJarimeAyab.FirstOrDefault(p => p.UserID == UserID && p.Month == month && p.Year == year).Moadel;

            }


            else
            {
                return -100000000000;

            }

            //double addorreduc = 0;
            //var xx = db.tbAddReduceMoadelkarkard.FirstOrDefault(p => p.FKUser == UserID && p.addred_Month == month && p.addred_Year == year);

            //if (xx != null)
            //{
            //    addorreduc = (double)xx.addred_Value;

            //}

            //var listmoalefekarkardi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 8).ToList();//لیست تمام مولفه های کارکردی
            //double? result = 0;
            //foreach (var item in listmoalefekarkardi)
            //{
            //    double? taghsim = 0;
            //    double? value = 0;
            //    var Rawvalue = db.tbMoalefeDastmozdiValueFromExcel.FirstOrDefault(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == month && p.MoalfeVal_Year == year && p.MoalfeVal_FKMoalafeDastmozdi == item.md_ID);
            //    if (Rawvalue != null)
            //    {
            //        value = Rawvalue.MoalfeVal_Value;
            //    }
            //    var standard = db.tbCaranSettings.FirstOrDefault(p => p.FK_Moalefe_ID == item.md_ID).CaranStandard;
            //    if (standard != 0)
            //    {
            //        taghsim = (value / standard);
            //    }
            //    result += Math.Round((double)taghsim, 4);
            //}
            //double? zaribshahrestan = 1;
            //var cityID = db.tbUsers.FirstOrDefault(p => p.usr_ID == UserID).usr_City_Dutysystem;
            //if (cityID != null)
            //{
            //    var x = db.tbCities.FirstOrDefault(p => p.ID == cityID);
            //    if (x.ZaribSharestan != null)
            //    {
            //        zaribshahrestan = x.ZaribSharestan;
            //    }
            //}
            //result = result * zaribshahrestan;

            //result += addorreduc;

            //return (double)result;
        }


        public double CalculateAyabOzahab(int UserID, int month, int year)
        {


            var find = db.tbMoadelPadashJarimeAyab.Where(p => p.UserID == UserID && p.Month == month && p.Year == year).FirstOrDefault();
            if (find != null)
            {
                return db.tbMoadelPadashJarimeAyab.FirstOrDefault(p => p.UserID == UserID && p.Month == month && p.Year == year).AyabOZahab;

            }


            else
            {
                return 0;

            }
            //var listmoalefekarkardi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 8).ToList();//لیست تمام مولفه های کارکردی
            //double? result = 0;
            //foreach (var item in listmoalefekarkardi)
            //{
            //    double? value = 0;
            //    var Rawvalue = db.tbMoalefeDastmozdiValueFromExcel.FirstOrDefault(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == month && p.MoalfeVal_Year == year && p.MoalfeVal_FKMoalafeDastmozdi == item.md_ID);
            //    if (Rawvalue != null)
            //    {
            //        value = Rawvalue.MoalfeVal_Value;
            //    }
            //    var tbcaran = db.tbCaranSettings.FirstOrDefault(p => p.FK_Moalefe_ID == item.md_ID);
            //    double ayabozahab = (double)tbcaran.CaranAyabOZahab;
            //    if (value != 0 && ayabozahab != 0)
            //    {

            //    }
            //    result += (value * ayabozahab);
            //}
            //double? zaribshahrestan = 1;
            //var cityID = db.tbUsers.FirstOrDefault(p => p.usr_ID == UserID).usr_City_Dutysystem;
            //if (cityID != null)
            //{

            //    var x = db.tbCities.FirstOrDefault(p => p.ID == cityID);
            //    if (x.ZaribSharestan != null)
            //    {
            //        zaribshahrestan = x.ZaribSharestan;
            //    }
            //}
            //result = result * zaribshahrestan;
            //return (double)result;
        }



        public double CalculatePadashForkarmozdi(int UserID, int Month, int Year)
        {
            return 0;



            //return db.tbMoadelPadashJarimeAyab.FirstOrDefault(p => p.UserID == UserID && p.Month == Month && p.Year == Year).Padash;
            //var listmoalefekarkardi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 8).ToList();//لیست تمام مولفه های کارکردی
            //double? result = 0;

            //foreach (var item in listmoalefekarkardi)
            //{

            //    double? value = 0;
            //    var Rawvalue = db.tbMoalefeDastmozdiValueFromExcel.FirstOrDefault(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year && p.MoalfeVal_FKMoalafeDastmozdi == item.md_ID);
            //    if (Rawvalue != null)
            //    {
            //        value = Rawvalue.MoalfeVal_Value;
            //    }
            //    var tbcaran = db.tbCaranSettings.FirstOrDefault(p => p.FK_Moalefe_ID == item.md_ID);
            //    int shoropadash = (int)tbcaran.CaranAstaneShoroPadash;
            //    if (value > shoropadash)
            //    {
            //        result += (tbcaran.CaranSaranePadash) * (value - shoropadash);
            //    }


            //}
            //return (double)result;
        }


        public double CalculateJarimeForKarmozdi(int UserID, int Month, int Year)
        {


            var find = db.tbMoadelPadashJarimeAyab.Where(p => p.UserID == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
            if (find != null)
            {
                return 0;
                //return db.tbMoadelPadashJarimeAyab.FirstOrDefault(p => p.UserID == UserID && p.Month == Month && p.Year == Year).Jarime;

            }


            else
            {
                return 0;

            }
            //var listmoalefekarkardi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Type == 8).ToList();//لیست تمام مولفه های کارکردی
            //double? result = 0;

            //foreach (var item in listmoalefekarkardi)
            //{

            //    double? value = 0;
            //    var Rawvalue = db.tbMoalefeDastmozdiValueFromExcel.FirstOrDefault(p => p.MoalfeVal_FKUser == UserID && p.MoalfeVal_Month == Month && p.MoalfeVal_Year == Year && p.MoalfeVal_FKMoalafeDastmozdi == item.md_ID);
            //    if (Rawvalue != null)
            //    {
            //        value = Rawvalue.MoalfeVal_Value;
            //    }
            //    var tbcaran = db.tbCaranSettings.FirstOrDefault(p => p.FK_Moalefe_ID == item.md_ID);
            //    int shorojarime = (int)tbcaran.CaranAstaneShoroJarime;
            //    if (value < shorojarime && value != 0)
            //    {
            //        result += (tbcaran.CaranSaraneJarime) * (shorojarime - value);
            //    }


            //}
            //return (double)result;
        }
        public double CalculateTaxTajamoetest2(double mashmoolmaliyat, int year, int Month, int UserID, int eydivapadash)
        {

            #region بدست آوردن مشمول مالیات ها از ماه شروع فرد تا به حال

            float result = 0;
            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(year, Month);
            var MoalefeID = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "جمع کل مشمول مالیات (ریال)").md_ID;

            var havecontract = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID && p.usc_StartTime <= myDatetime && p.usc_EndTime >= myDatetime);

            if (havecontract != null)
            {
                int startMonth = havecontract.ShamsiStartTime_month;
                for (int i = startMonth; i < Month; i++)
                {
                    var x111 = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeID &&p.FK_EXCel==null&& p.mlfvlfsh_Submit == true);
                    if (x111 != null)
                    {
                        mashmoolmaliyat += (double)x111.mlfvlfsh_Value;

                    }
                }
                #endregion


                #region محاسبه مالیات تجمعی بر اساس مشمول مالیات فرد تا به الان

                //double teadadmah = havecontract.FirstMonthFish+1??0;// Month - startMonth + 1;//بدست آوردن تعداد ماه تا به الان
                var MoalefeID22 = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                var t = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == year && p.FK_User == UserID && p.mlfvlfsh_Submit == true && p.FK_EXCel == null && p.FK_Moalefe == MoalefeID22 && p.mlfvlfsh_Value > 0).ToList();
                double teadadmah = t.Count + 1;
                if (Month == 12)//اگر ماه دوازده بود عیدی و پاداش اضافه میشه و همچنین یدونه معافیت مالیاتی بهش اضافه میشه یعنی تعداد ماه یدونه اضافه میشه
                {
                    mashmoolmaliyat += eydivapadash;
                    teadadmah += (teadadmah / 12);
                }


                var liststeptaxes = stepTaxRepository.List().Where(p => p.year == year).OrderBy(p => p.number).ToList();
                float min = 0;
                float max = 0;
                int step = 0;
                int ii = 0;
                double x = 0;

                foreach (var item in liststeptaxes)
                {
                    if (Month == 12)
                    {
                        if (ii == 1)
                        {
                            min = (float)(item.Fromamount * (teadadmah));
                            max = (float)(item.Toamount * (teadadmah - 1));
                            max += (float)x;
                            //min += (float)x;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                            ii = 2;
                        }
                        else if (ii != 1 && ii != 0)
                        {
                            min = (float)(item.Fromamount * (teadadmah - 1));
                            max = (float)(item.Toamount * (teadadmah - 1));
                            max += (float)x;
                            min += (float)x;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                        }

                        else
                        {
                            min = (float)(item.Fromamount * teadadmah);
                            max = (float)(item.Toamount * teadadmah);
                            x = (double)item.Toamount;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                            ii = 1;
                        }
                    }


                    else
                    {
                        min = (float)(item.Fromamount * teadadmah);
                        max = (float)(item.Toamount * teadadmah);
                        if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                        {
                            step = (int)item.number;
                            break;
                        }

                    }


                }
                double x11 = 0;
                ii = 0;

                for (int i = 1; i < step; i++)
                {

                    if (Month == 12)
                    {
                        if (ii == 1)
                        {
                            ii = 2;
                            result = (float)((((liststeptaxes[i - 1].Toamount) * (teadadmah - 1)) + x11 - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                        }
                        else if (ii != 1 && ii != 0)
                        {

                            result = (float)(((((liststeptaxes[i - 1].Toamount) * (teadadmah - 1)) + x11) - ((liststeptaxes[i - 1].Fromamount * (teadadmah - 1)) + x11)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                        }
                        else
                        {
                            result = (float)((((liststeptaxes[i - 1].Toamount) * teadadmah) - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                            ii = 1;
                            x11 = (double)(((liststeptaxes[i - 1].Toamount)));
                        }

                    }
                    else
                    {
                        result = (float)((((liststeptaxes[i - 1].Toamount) * teadadmah) - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;

                    }


                }
                if (Month == 12)
                {
                    result = (float)((mashmoolmaliyat - ((liststeptaxes[step - 1].Fromamount * (teadadmah - 1)) + x11)) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;

                }

                else
                {
                    result = (float)((mashmoolmaliyat - (liststeptaxes[step - 1].Fromamount * teadadmah)) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;

                }
                #endregion

                #region کم کردن میزان مالیاتی که تا الان داده از مالیاتی که این ماه براش حساب کردیم

                var MoalefeIDmaliyat = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "مالیات سهم کارمند (ریال)").md_ID;
                int LastMonth = 1;
                if (Month != 1)
                {
                    LastMonth = Month - 1;
                }
                int jammaliyatqabl = 0;

                for (int i = LastMonth; i >= 1; i--)
                {
                    var moalefefish = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.FK_EXCel == null && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeIDmaliyat && p.mlfvlfsh_Submit == true);
                    if (moalefefish != null)
                    {
                        jammaliyatqabl += Convert.ToInt32(moalefefish.mlfvlfsh_Value);
                    }
                }
                //var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                //if (cit.usr_City_Dutysystem != null)
                //{
                //    var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                //    if (city.Zaribmalyat != null)
                //    {
                //        result = (float)city.Zaribmalyat / 100 * result;

                //    }
                //}
                if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null /*&& db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null*/)
                {
                    var lonktbusr = db.Link_User_And_Peyman.Where(s => s.FK_User_ID == UserID && s.Status == true).FirstOrDefault();
                    if (lonktbusr != null)
                    {
                        var findusrlinkcit = db.tbpeymancities.Where(s => s.FK_PYMN == lonktbusr.FK_Peyman_ID).FirstOrDefault();
                        if (findusrlinkcit != null)
                        {
                            var city = db.tbCities.Where(p => p.ID == findusrlinkcit.FK_City).FirstOrDefault();
                            if (city.Zaribmalyat != null)
                            {
                                result = (float)city.Zaribmalyat / 100 * result;

                            }
                            else
                            {
                                //var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                                //if (cit.usr_City_Dutysystem != null)
                                //{
                                //    var city2 = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                                //    if (city2.Zaribmalyat != null)
                                //    {
                                //        result = (float)city2.Zaribmalyat / 100 * result;

                                //    }
                                //}
                            }
                        }
                        else
                        {
                            var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                            if (cit.usr_City_Dutysystem != null)
                            {
                                var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                                if (city.Zaribmalyat != null)
                                {
                                    result = (float)city.Zaribmalyat / 100 * result;

                                }
                            }
                        }
                    }
                    else
                    {
                        var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                        if (cit.usr_City_Dutysystem != null)
                        {
                            var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                            if (city.Zaribmalyat != null)
                            {
                                result = (float)city.Zaribmalyat / 100 * result;

                            }
                        }

                    }
                }
                else
                {
                    var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                    if (cit.usr_City_Dutysystem != null)
                    {
                        var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                        if (city.Zaribmalyat != null)
                        {
                            result = (float)city.Zaribmalyat / 100 * result;

                        }
                    }
                }
                if (result - jammaliyatqabl < 0)
                {
                    result = 0;
                }
                else
                {
                    result = result - jammaliyatqabl;
                }



                #endregion

            }

            return result;
        }


        public double CalculateTaxTajamoe(double mashmoolmaliyat, int year, int Month, int UserID, int eydivapadash)
        {

            #region بدست آوردن مشمول مالیات ها از ماه شروع فرد تا به حال

            float result = 0;
            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(year, Month);
            var MoalefeID = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "جمع کل مشمول مالیات (ریال)").md_ID;

            var havecontract = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID && p.usc_StartTime <= myDatetime && p.usc_EndTime >= myDatetime);

            if (havecontract != null)
            {
                int startMonth = havecontract.ShamsiStartTime_month;
                for (int i = startMonth; i < Month; i++)
                {
                    var x111 = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeID && p.mlfvlfsh_Submit == true);
                    if (x111 != null)
                    {
                        mashmoolmaliyat += (double)x111.mlfvlfsh_Value;

                    }
                }
                #endregion


                #region محاسبه مالیات تجمعی بر اساس مشمول مالیات فرد تا به الان

                //double teadadmah = havecontract.FirstMonthFish+1??0;// Month - startMonth + 1;//بدست آوردن تعداد ماه تا به الان
                var MoalefeID22 = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                var t = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == year && p.FK_User == UserID && p.mlfvlfsh_Submit == true && p.FK_Moalefe==MoalefeID22 && p.mlfvlfsh_Value>0).ToList();
                double teadadmah = t.Count + 1;
                if (Month == 12)//اگر ماه دوازده بود عیدی و پاداش اضافه میشه و همچنین یدونه معافیت مالیاتی بهش اضافه میشه یعنی تعداد ماه یدونه اضافه میشه
                {
                    mashmoolmaliyat += eydivapadash;
                    teadadmah += (teadadmah / 12);
                }


                var liststeptaxes = stepTaxRepository.List().Where(p => p.year == year).OrderBy(p => p.number).ToList();
                float min = 0;
                float max = 0;
                int step = 0;
                int ii = 0;
                double x = 0;

                foreach (var item in liststeptaxes)
                {
                    if (Month == 12)
                    {
                        if (ii == 1)
                        {
                            min = (float)(item.Fromamount * (teadadmah));
                            max = (float)(item.Toamount * (teadadmah - 1));
                            max += (float)x;
                            //min += (float)x;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                            ii = 2;
                        }
                        else if (ii != 1 && ii != 0)
                        {
                            min = (float)(item.Fromamount * (teadadmah - 1));
                            max = (float)(item.Toamount * (teadadmah - 1));
                            max += (float)x;
                            min += (float)x;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                        }

                        else
                        {
                            min = (float)(item.Fromamount * teadadmah);
                            max = (float)(item.Toamount * teadadmah);
                            x = (double)item.Toamount;
                            if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                            {
                                step = (int)item.number;
                                break;
                            }
                            ii = 1;
                        }
                    }
                    
                    
                    else
                    {
                        min = (float)(item.Fromamount * teadadmah);
                        max = (float)(item.Toamount * teadadmah);
                        if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                        {
                            step = (int)item.number;
                            break;
                        }

                    }

                    
                }
                double x11 = 0;
                 ii = 0;

                for (int i = 1; i < step; i++)
                {

                    if (Month == 12)
                    {
                        if (ii == 1)
                        {
                            ii = 2;
                            result = (float)((((liststeptaxes[i - 1].Toamount) * (teadadmah-1))+ x11 - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                        }
                        else if (ii != 1&&ii!=0)
                        {

                            result = (float)(((((liststeptaxes[i - 1].Toamount) * (teadadmah-1))+ x11) -((liststeptaxes[i - 1].Fromamount * (teadadmah - 1)) + x11)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                        }
                        else
                        {
                            result = (float)((((liststeptaxes[i - 1].Toamount) * teadadmah) - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                            ii = 1;
                            x11 = (double)(((liststeptaxes[i - 1].Toamount)));
            }

                    }
                    else
                    {
                        result = (float)((((liststeptaxes[i - 1].Toamount) * teadadmah) - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;

                    }


                }
                if (Month == 12)
                {
                    result = (float)((mashmoolmaliyat - ((liststeptaxes[step - 1].Fromamount * (teadadmah-1))+x11)) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;

                }

                else
                {
                    result = (float)((mashmoolmaliyat - (liststeptaxes[step - 1].Fromamount * teadadmah)) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;

                }
                #endregion

                #region کم کردن میزان مالیاتی که تا الان داده از مالیاتی که این ماه براش حساب کردیم

                var MoalefeIDmaliyat = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "مالیات سهم کارمند (ریال)").md_ID;
                int LastMonth = 1;
                if (Month != 1)
                {
                    LastMonth = Month - 1;
                }
                int jammaliyatqabl = 0;

                for (int i = LastMonth; i >=1; i--)
                {
                    var moalefefish = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeIDmaliyat && p.mlfvlfsh_Submit == true);
                    if (moalefefish != null)
                    {
                        jammaliyatqabl += Convert.ToInt32(moalefefish.mlfvlfsh_Value);
                    }
                }
                var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                if (cit.usr_City_Dutysystem != null)
                {
                    var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                    if (city.Zaribmalyat != null)
                    {
                        result = (float)city.Zaribmalyat / 100 * result;

                    }
                }

                if (result - jammaliyatqabl < 0)
                {
                    result = 0;
                }
                else
                {
                    result = result - jammaliyatqabl;
                }
             
             

                #endregion

            }

            return result;
        }

        public double CalculateTaxTajamoetsttt(double mashmoolmaliyat, int year, int Month, int UserID, int eydivapadash)
        {

            #region بدست آوردن مشمول مالیات ها از ماه شروع فرد تا به حال

            float result = 0;
            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(year, Month);
            var MoalefeID = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "جمع کل مشمول مالیات (ریال)").md_ID;

            var havecontract = db.tbUserContracts.FirstOrDefault(p => p.FK_UserID == UserID && p.usc_StartTime <= myDatetime && p.usc_EndTime >= myDatetime);

            if (havecontract != null)
            {
                int startMonth = havecontract.ShamsiStartTime_month;
                for (int i = startMonth; i < Month; i++)
                {
                    var x = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeID && p.mlfvlfsh_Submit == true);
                    if (x != null)
                    {
                        mashmoolmaliyat += (double)x.mlfvlfsh_Value;

                    }
                }
                #endregion


                #region محاسبه مالیات تجمعی بر اساس مشمول مالیات فرد تا به الان

                //double teadadmah = havecontract.FirstMonthFish+1??0;// Month - startMonth + 1;//بدست آوردن تعداد ماه تا به الان
                var MoalefeID22 = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "خالص قابل دریافت (ریال)").md_ID;
                var t = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == year && p.FK_User == UserID && p.mlfvlfsh_Submit == true && p.FK_Moalefe == MoalefeID22 && p.mlfvlfsh_Value > 0).ToList();
                double teadadmah = t.Count + 1;
                if (Month == 12)//اگر ماه دوازده بود عیدی و پاداش اضافه میشه و همچنین یدونه معافیت مالیاتی بهش اضافه میشه یعنی تعداد ماه یدونه اضافه میشه
                {
                    mashmoolmaliyat += eydivapadash;
                    teadadmah += (teadadmah / 12);
                }


                var liststeptaxes = stepTaxRepository.List().Where(p => p.year == year).OrderBy(p => p.number).ToList();
                float min = 0;
                float max = 0;
                int step = 0;
                foreach (var item in liststeptaxes)
                {
                    min = (float)(item.Fromamount * teadadmah);
                    max = (float)(item.Toamount * teadadmah);
                    if (mashmoolmaliyat >= min && mashmoolmaliyat <= max)
                    {
                        step = (int)item.number;
                        break;
                    }
                }
                for (int i = 1; i < step; i++)
                {
                    result = (float)((((liststeptaxes[i - 1].Toamount) * teadadmah) - (liststeptaxes[i - 1].Fromamount * teadadmah)) * ((liststeptaxes[i - 1].Percent_amount) / 100)) + result;
                }
                result = (float)((mashmoolmaliyat - (liststeptaxes[step - 1].Fromamount * teadadmah)) * ((liststeptaxes[step - 1].Percent_amount) / 100)) + result;
                #endregion

                #region کم کردن میزان مالیاتی که تا الان داده از مالیاتی که این ماه براش حساب کردیم

                var MoalefeIDmaliyat = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "مالیات سهم کارمند (ریال)").md_ID;
                int LastMonth = 1;
                if (Month != 1)
                {
                    LastMonth = Month - 1;
                }
                int jammaliyatqabl = 0;

                for (int i = LastMonth; i >= 1; i--)
                {
                    var moalefefish = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Year == year && p.mlfvlfsh_Month == i && p.FK_Moalefe == MoalefeIDmaliyat && p.mlfvlfsh_Submit == true);
                    if (moalefefish != null)
                    {
                        jammaliyatqabl += Convert.ToInt32(moalefefish.mlfvlfsh_Value);
                    }
                }
                //var lonktbusr = db.Link_User_And_Peyman.Where(s => s.FK_User_ID == UserID&&s.Status==true).FirstOrDefault();
                //if (lonktbusr != null)
                //{
                // var findusrlinkcit=   db.tbpeymancities.Where(s => s.FK_PYMN == lonktbusr.FK_Peyman_ID).FirstOrDefault();
                //    if (findusrlinkcit!=null)
                //    {
                //        var city = db.tbCities.Where(p => p.ID == findusrlinkcit.FK_City).FirstOrDefault();
                //        if (city.Zaribmalyat != null)
                //        {
                //            result = (float)city.Zaribmalyat / 100 * result;

                //        }
                //        else
                //        {
                //            //var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                //            //if (cit.usr_City_Dutysystem != null)
                //            //{
                //            //    var city2 = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                //            //    if (city2.Zaribmalyat != null)
                //            //    {
                //            //        result = (float)city2.Zaribmalyat / 100 * result;

                //            //    }
                //            //}
                //        }
                //    }
                //    else
                //    {
                //        var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                //        if (cit.usr_City_Dutysystem != null)
                //        {
                //            var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                //            if (city.Zaribmalyat != null)
                //            {
                //                result = (float)city.Zaribmalyat / 100 * result;

                //            }
                //        }
                //    }
                //}
                //else
                //{
                //    var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                //    if (cit.usr_City_Dutysystem != null)
                //    {
                //        var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                //        if (city.Zaribmalyat != null)
                //        {
                //            result = (float)city.Zaribmalyat / 100 * result;

                //        }
                //    }

                //}
                var cit = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
                if (cit.usr_City_Dutysystem != null)
                {
                    var city = db.tbCities.Where(p => p.ID == cit.usr_City_Dutysystem).FirstOrDefault();
                    if (city.Zaribmalyat != null)
                    {
                        result = (float)city.Zaribmalyat / 100 * result;

                    }
                }

                if (result - jammaliyatqabl < 0)
                {
                    result = 0;
                }
                else
                {
                    result = result - jammaliyatqabl;
                }



                #endregion

            }

            return result;
        }




        public int FillMohasebemaheqabl(int Month, int year)
        {
            try
            {
                var YearandMonth = MonthforUsersalarDetail(Month, year);
                var exist = db.tbUserSalaryDetail.FirstOrDefault(p => p.usrslrydtl_Year == YearandMonth.Item1 && p.usrslrydtl_Month == YearandMonth.Item2);
                if (exist != null)
                {
                    exist.usrslrydtl_LastMonthCalculate = true;
                    db.SaveChanges();
                }
                else
                {
                    tbUserSalaryDetail salardetail = new tbUserSalaryDetail
                    {
                        usrslrydtl_LastMonthCalculate = true,
                        usrslrydtl_Consolidation = false,
                        usrslrydtl_PerformanceSave = false,
                        usrslrydtl_PerformanceAccept = false,
                        usrslrydtl_Month = YearandMonth.Item2,
                        usrslrydtl_Year = YearandMonth.Item1

                    };

                    if (tbusersalarydetailtRepo.Create(salardetail) != "True")
                    {
                        return 0;
                    }
                }
                return 1;

            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public int FillQateemaheqabl(int Month, int Year)
        {
            try
            {
                var YearandMonth = MonthforUsersalarDetail(Month, Year);
                var exist = db.tbUserSalaryDetail.FirstOrDefault(p => p.usrslrydtl_Year == YearandMonth.Item1 && p.usrslrydtl_Month == YearandMonth.Item2);
                if (exist != null)
                {
                    exist.usrslrydtl_Consolidation = true;
                    db.SaveChanges();
                }
                else
                {
                    tbUserSalaryDetail salardetail = new tbUserSalaryDetail
                    {
                        usrslrydtl_LastMonthCalculate = false,
                        usrslrydtl_Consolidation = true,
                        usrslrydtl_PerformanceSave = false,
                        usrslrydtl_PerformanceAccept = false,
                        usrslrydtl_Year = YearandMonth.Item1,
                        usrslrydtl_Month = YearandMonth.Item2

                    };

                    if (tbusersalarydetailtRepo.Create(salardetail) != "true")
                    {
                        return 0;
                    }
                }
                return 1;

            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public Tuple<int, int> MonthforUsersalarDetail(int Month, int Year)
        {
            if (Month == 12)
            {
                return System.Tuple.Create(Year - 1, 1);
            }
            return System.Tuple.Create(Year, Month + 1);
        }
            


        //public bool CalculateFromSQL(int Month,int Year,int UserID)
        //{
        //    try
        //    {
        //        db.Database.CommandTimeout = 400;
            
        //        db.CalculateJarimePadashAyabMaodel(UserID, Month, Year);

        //        return true;
        //    }
        //    catch (Exception e)
        //    {

        //        return false;
        //    }     
        //}


        //-------------------------------------------------------testtttt
        public double Calculatekarkardtaghsimbarkolemahtstt(double tedadroozkarkard, double tedadroozmah)
        {
            double result = tedadroozkarkard / tedadroozmah;
            return result;
        }

        public int teadadroozkarkardforfishmainaslytest(int qeybat, int tedadroozmah, int tedadstelaji, double moadelkarkar)
        {
            if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null /*&& db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null*/)
            {
                //moadelkarkar = 1;
                double moadel = Math.Ceiling((moadelkarkar * tedadroozmah));
                if (moadel > tedadroozmah)
                {
                    moadel = tedadroozmah;
                }
                if (tedadstelaji > 3)
                {
                    return (int)moadel - (qeybat + tedadstelaji);

                }
                else
                {
                    return (int)moadel - (qeybat);

                }
            }
            else
            {
                //moadelkarkar = 1;
                double moadel = Math.Ceiling((moadelkarkar * tedadroozmah) + tedadstelaji);
                if (moadel > tedadroozmah)
                {
                    moadel = tedadroozmah;
                }
               
                    return (int)moadel - (qeybat + tedadstelaji);

                
               
            }
             
        }
        public (int qybat, double moadelkarkarOut) rozgaybatkasri(int UserID, int Month, int Year, double moadelkarkar)
        {
            double ragamkol = 0;
            var mozdsayer = "0";
            var hagholadbase = "0";
            var aghlammasrafikhanevarbase = "0";
            var tedadroozmah = TeadaroozMonth(Month, Year);
            var haghmaskanbase = "0";
            var myDatetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);
            var contractinfo = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_EndTime >= myDatetime && p.usc_StartTime <= myDatetime).FirstOrDefault();
            if (contractinfo != null)
            {
                var listmoalefeqarardadi = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKContractID == contractinfo.usc_ID).ToList();
                foreach(var it in listmoalefeqarardadi)
                {
                    ragamkol += (double)it.Value;
                }
                var mozdsanavatroozane = Regex.Replace(contractinfo.jobgroup_ValueSanavat, ",", "");
                if (mozdsanavatroozane == null)
                {
                    mozdsanavatroozane = "0";
                }
                double mozdmahaneh = Convert.ToDouble(mozdsanavatroozane) * tedadroozmah;

                var mozdshoqlroozane = Regex.Replace(contractinfo.jobgroup_ValueMozdGroup, ",", "");
                if (mozdshoqlroozane == null)
                {
                    mozdshoqlroozane = "0";
                }
                double mozdmahanehshogl = Convert.ToDouble(mozdshoqlroozane) * tedadroozmah;

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
                ragamkol= ragamkol+ Convert.ToDouble(aghlammasrafikhanevarbase)+ Convert.ToDouble(hagholadbase) + Convert.ToDouble(haghmaskanbase) + Convert.ToDouble(mozdmahaneh) + Convert.ToDouble(mozdmahanehshogl) ;
            }
            double moadelkarkarasli = (double)(1 - moadelkarkar);

            double hesab = (double)(moadelkarkarasli * ragamkol);
            var dbolavaitkasri1 = db.dbolavaitkasri1.OrderBy(s => s.olaviat).ToList();
            double ezafat = 0;
            foreach (var it in dbolavaitkasri1)
            {
                if (hesab > 0)
                {
                    double tosavedat = 0;
                    var find = db.tbMoalefeValueFish.Where(s => s.FK_User == UserID && s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_Moalefe == it.FK_moalfeh && s.Finalaccept == true).FirstOrDefault();
                    if (find != null)
                    {
                        ezafat += (double)(find.mlfvlfsh_Value);

                        if (find.mlfvlfsh_Value >= hesab)
                        {

                            tosavedat = (double)(find.mlfvlfsh_Value - hesab);
                            hesab = 0; ;

                        }
                        else
                        {
                            tosavedat = 0;
                            hesab -= (double)(find.mlfvlfsh_Value);
                        }
                        tbmoalfefishexcel tbmoalfefishexcel = new tbmoalfefishexcel();
                        tbMoalefeValueFish tbMoalefeValueFish = new tbMoalefeValueFish();

                        var findtbmoalfefishexcel = db.tbmoalfefishexcel.Where(s => s.tbMoalefeValueFish.Any(t => t.mlfvlfsh_Month == Month && t.mlfvlfsh_Year == Year && t.Finalaccept == true) && s.Fk_Pymn == null && s.usr_sabt == null).FirstOrDefault();
                        if (findtbmoalfefishexcel == null)
                        {
                            tbmoalfefishexcel.submitt = true;
                            tbmoalfefishexcel.FK_control = true;
                            db.tbmoalfefishexcel.Add(tbmoalfefishexcel);
                            db.SaveChanges();
                        }
                        else
                        {
                            tbmoalfefishexcel = findtbmoalfefishexcel;
                        }
                        var find2 = db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == Month && s.FK_User == UserID && s.mlfvlfsh_Year == Year && s.FK_EXCel == tbmoalfefishexcel.ID && s.FK_Moalefe == it.FK_moalfeh && s.Finalaccept == true).FirstOrDefault();
                        if (find2 == null)
                        {
                            var findkol = db.tbMoalefeValueFish.Where(s => s.FK_User == UserID && s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_Moalefe == it.FK_moalfeh && s.Finalaccept == true).ToList();
                            foreach (var it1 in findkol)
                            {
                                it1.Finalaccept = null;
                                db.SaveChanges();
                            }
                            tbMoalefeValueFish.mlfvlfsh_Month = Month;
                            tbMoalefeValueFish.mlfvlfsh_Year = Year;
                            tbMoalefeValueFish.mlfvlfsh_Value = tosavedat;
                            tbMoalefeValueFish.FK_EXCel = tbmoalfefishexcel.ID;
                            tbMoalefeValueFish.FK_User = UserID;
                            tbMoalefeValueFish.FK_Moalefe = it.FK_moalfeh;
                            tbMoalefeValueFish.Finalaccept = true;
                            db.tbMoalefeValueFish.Add(tbMoalefeValueFish);
                            db.SaveChanges();
                        }
                        else
                        {
                            var findkol = db.tbMoalefeValueFish.Where(s => s.FK_User == UserID && s.mlfvlfsh_Month == Month && s.mlfvlfsh_Year == Year && s.FK_Moalefe == it.FK_moalfeh && s.Finalaccept == true).ToList();
                            foreach (var it1 in findkol)
                            {
                                it1.Finalaccept = null;
                                db.SaveChanges();
                            }
                            find2.mlfvlfsh_Value = tosavedat;
                            find2.Finalaccept = true;

                            db.SaveChanges();
                        }
                     
                    
                    }
                }
               

            }
            int qybat = 0;
            double moadelkoklkasri = (double)(hesab / ragamkol);
            qybat = (int)(moadelkoklkasri * tedadroozmah);
            moadelkarkar = (1 - 0);
            return (qybat, moadelkarkar);

        }
        public ManualFishDetail Calculateandinserttotabletestbanaei(int UserID, int Month, int Year, bool qatee, FishHeader fishHeader,bool contikar)
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



            //PersianCalendar persianCalendar = new PersianCalendar();


            //int contractYear = persianCalendar.GetYear(endTime);
            //int contractMonth = persianCalendar.GetMonth(endTime);


            bool Karmozdi = false;
            DateTime miladiDate = ConvertShamsiToMiladi(Year, Month);

            var t = db.tbUserContracts.Where(p => p.FK_UserID == UserID && miladiDate >= p.usc_StartTime && miladiDate <= p.usc_EndTime && p.usc_TypeOfContract == true).FirstOrDefault();
            if (t != null)
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
            var moalefehayeexcelfkexcel = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID && p.Finalaccept == true&&p.FK_EXCel==null).ToList();

            double? Saatezafkar = 0;


            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار") != null)
            {
                Saatezafkar = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات اضافه کار").mlfvlfsh_Value;
            }


            double? stelaji = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی") != null)
            {
                stelaji = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز استعلاجی").mlfvlfsh_Value;
            }


            double? roozmamoriat = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت") != null)
            {
                roozmamoriat = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز ماموریت").mlfvlfsh_Value;
            }

            double? saatkasrkar = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar =(double) moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }

            double? roozqybt = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت") != null)
            {
                roozqybt = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز غیبت").mlfvlfsh_Value;
            }

            var moadelkarkar = CalculateMoadelKarkard(UserID, Month, Year);
            //محاسبه کسری کار کرد 
            if (contikar == true)
            {
                if (moadelkarkar > 0)
                {
                    if (moadelkarkar < 1)
                    {
                        var result = rozgaybatkasri(UserID, Month, Year, moadelkarkar);
                        roozqybt += result.qybat;
                        moadelkarkar = result.moadelkarkarOut;
                        moalefehayeexcel = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID && p.Finalaccept == true).ToList();
                         moalefehayeexcelfkexcel = db.tbMoalefeValueFish.Where(p => p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID && p.Finalaccept == true && p.FK_EXCel == null).ToList();

                    }
                }
            }
           
            double? ayabzahab = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
            {
                if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
                {
                    ayabzahab = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                }
                else
                {
                    ayabzahab = 0;

                }
                //ayabzahab = 0;

                //ayabzahab = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
                {
                    ayabzahab = CalculateAyabOzahab(UserID, Month, Year);

                }
                else
                {
                    ayabzahab = 0;

                }
                //ayabzahab = 0;

                //ayabzahab = CalculateAyabOzahab(UserID, Month, Year);
            }



            double? jome = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه") != null)
            {
                jome = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد جمعه").mlfvlfsh_Value;
            }

            var tedadroozmah = TeadaroozMonth(Month, Year);


            double? tatil = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل") != null)
            {
                tatil = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کارکرد روز تعطیل").mlfvlfsh_Value;
            }

           
           
                if (Month == 1)
            {
                var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == 12 && p.Year == Year - 1).FirstOrDefault();
                if (ezaf != null)
                {
                    if (ezaf.RemainValue != null)
                    {
                        moadelkarkar += (double)(ezaf.RemainValue / 162);
                    }
                }
                var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                //if (tedadsaatezafkar > exi.Value && exi != null)
                //{


                //    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                //    tedadsaatezafkar = (double)exi.Value;
                //    db.SaveChanges();



                //}
            }
            else
            {
                var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month - 1 && p.Year == Year).FirstOrDefault();
                if (ezaf != null)
                {
                    if (ezaf.RemainValue != null)
                    {
                        moadelkarkar += (double)(ezaf.RemainValue / 162);

                    }
                }
                //var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                //if (tedadsaatezafkar > exi.Value && exi != null)
                //{


                //    exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                //    tedadsaatezafkar = (double)exi.Value;
                //    db.SaveChanges();



                //}
            }
            //int teadadroozkarkard = teadadroozkarkardforfishmain((int)roozqybt, tedadroozmah, (int)stelaji);
            if (moadelkarkar == -100000000000)
            {
                moadelkarkar = 1;
                if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_ID == 1960) != null)
                {
                    double moadeltest = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_ID == 1960).mlfvlfsh_Value;

                    moadelkarkar = (double)(moadeltest / tedadroozmah);

                }
                else
                {
                    moadelkarkar = (double)(tedadroozmah / tedadroozmah);

                }
            }
          
            int teadadroozkarkard = teadadroozkarkardforfishmainaslytest((int)roozqybt, tedadroozmah, (int)stelaji, (double)moadelkarkar);





            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)") != null)
            {
                padash = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "پاداش(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                padash = CalculatePadashForkarmozdi(UserID, Month, Year);
            }



            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)") != null)
            {
                jarime = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "جریمه(ریال)").mlfvlfsh_Value;
            }
            else if (Karmozdi)//اگر کارمزدی بود و برای آن ایاب و ذهاب دستی وارد نشده بود
            {
                jarime = CalculateJarimeForKarmozdi(UserID, Month, Year);
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

                var res = Calculatesaatkasrkarvaezafkarforkarmozdi(teadadroozkarkard, moadelkarkar, tedadroozmah, tedadroozestelaji);
                //tedadsaatezafkar = res.Item1;
                #region برای max اضاقه کاری








                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        if (Month == 1)
                        {
                            //var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == 12 && p.Year == Year - 1).FirstOrDefault();
                            //if (ezaf.RemainValue != null && ezaf!=null)
                            //{
                            //    tedadsaatezafkar +=(double) ezaf.RemainValue;
                            //}
                            var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                            if (tedadsaatezafkar > exi.Value && exi != null)
                            {


                                exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                                tedadsaatezafkar = (double)exi.Value;
                                db.SaveChanges();



                            }
                        }
                        else
                        {
                            //var ezaf = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month - 1 && p.Year == Year).FirstOrDefault();
                            //if (ezaf.RemainValue != null && ezaf != null)
                            //{
                            //    tedadsaatezafkar += (double)ezaf.RemainValue;
                            //}
                            var exi = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();

                            if (tedadsaatezafkar > exi.Value && exi != null)
                            {


                                exi.RemainValue = (double)tedadsaatezafkar - exi.Value;

                                tedadsaatezafkar = (double)exi.Value;
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
                tedadsaatezafkar = (double)Saatezafkar;


                tedadroozmamoriat = (double)roozmamoriat;
                ////tedadsaatkasrekar = res.Item2;


                MoadelKarkard = moadelkarkar;
                tedadroozQeybat = (double)roozqybt;
                tedadrooztatil = (double)tatil;
                teadaroozjome = (double)jome;

                ayabozahab = (double)ayabzahab;
            }



            var nesbatkarbkolmah = Calculatekarkardtaghsimbarkolemahtstt(teadadroozkarkard, TeadaroozMonth(Month, Year));





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
                var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == item.FKMoalefeGhararDadi && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID&&p.FK_EXCel==null);
                if (exist == null)
                {
                    tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                    {
                        FK_Moalefe = item.FKMoalefeGhararDadi,
                        mlfvlfsh_Submit = qatee,
                        FK_User = UserID,
                        mlfvlfsh_Month = Month,
                        mlfvlfsh_Year = Year,
                        Finalaccept = true,

                        mlfvlfsh_Value = item.Value * nesbatkarbkolmah
                    };
                    db.tbMoalefeValueFish.Add(moalefefish);
                    db.SaveChanges();
                }
                else if (exist.mlfvlfsh_Submit != true)
                {
                    exist.FK_Moalefe = item.FKMoalefeGhararDadi;
                    exist.mlfvlfsh_Submit = qatee;
                    exist.FK_User = UserID;
                    exist.mlfvlfsh_Month = Month;
                    exist.mlfvlfsh_Year = Year;
                    exist. Finalaccept = true;
                    exist.mlfvlfsh_Value = item.Value * nesbatkarbkolmah;
                    db.SaveChanges();
                }
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



            //test
            FishValue valuefish11 = new FishValue
            {
                Title = "تعداد روز کارکرد",
                Value = teadadroozkarkard /*(double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز کارکرد").mlfvlfsh_Value*/
            };
            fishValues.Add(valuefish11);





            double? tedadrooznobatkari1 = 0;
            var x = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری1 ( صبح‏وعصر)");
            if (x != null)
            {
                tedadrooznobatkari1 = x.mlfvlfsh_Value;
            }


            double? tedadrooznobatkari2 = 0;
            var y = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری2 (صبح‏وشب)");
            if (y != null)
            {
                tedadrooznobatkari2 = y.mlfvlfsh_Value;
            }



            double? tedadrooznobatkari3 = 0;
            var z = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری3 (عصروشب)");
            if (z != null)
            {
                tedadrooznobatkari3 = z.mlfvlfsh_Value;
            }



            double? tedadrooznobatkari4 = 0;
            var xx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد روز نوبتکاری4 (صبح‏وعصروشب)");
            if (xx != null)
            {
                tedadrooznobatkari4 = xx.mlfvlfsh_Value;
            }




            double? tedadroozshabkari = 0;
            var xy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد شب های کاری");
            if (tedadroozshabkari == null)
            {
                tedadroozshabkari = xy.mlfvlfsh_Value;
            }


            double? xz = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)") != null)
            {
                xz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اولاد(ریال)").mlfvlfsh_Value;
            }
            FishValue eslahiyekomakhazineolad = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اولاد(ریال)",
                Value = (double)xz
            };
            fishValues.Add(eslahiyekomakhazineolad);





            double? xxx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)") != null)
            {
                xxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه مسکن(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish2 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه مسکن(ریال)",
                Value = (double)xxx
            };
            fishValues.Add(valuefish2);





            double? xxy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)") != null)
            {
                xxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish3 = new FishValue
            {
                Title = "اصلاحیه کمک هزینه اقلام مصرفی خانوار(ریال)",
                Value = (double)xxy
            };
            fishValues.Add(valuefish3);






            if (db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault() != null && db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 9).FirstOrDefault() != null)
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",
                    Value = ayabozahab
                    //Value = 0

                };
                fishValues.Add(valuefish4);

            }
            else
            {
                FishValue valuefish4 = new FishValue
                {
                    Title = "کمک هزینه ایاب و ذهاب(ریال)",
                    //Value = ayabozahab
                    Value = 0

                };
                fishValues.Add(valuefish4);

            }





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
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)") != null)
            {
                xyz = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ابزار کار(ریال)").mlfvlfsh_Value;
            }
            FishValue valuefish7 = new FishValue
            {
                Title = "کمک هزینه ابزار کار(ریال)",
                //Value = (double)xyz
                Value = 0
            };
            fishValues.Add(valuefish7);





            double? yxx = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه") != null)
            {
                yxx = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه تبلت و رایانه پایه").mlfvlfsh_Value;
            }
            FishValue valuefish8 = new FishValue
            {
                Title = "کمک هزینه تبلت و رایانه(ریال)",
                /* Value = (double)yxx *//** moadelkarkar*/
                Value = 0
            };
            fishValues.Add(valuefish8);




            double? yxy = 0;
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)") != null)
            {
                yxy = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه بیمه تکمیل درمان (ریال)").mlfvlfsh_Value;
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

            string inputDate = "1" + "/" + Month + "/" + Year;
            DateTime persianDate = PersianDateToDateTime(inputDate);
            double ayabzahab2 = 0;
            var ayabzahab2000 = db.tbMoadelPadashJarimeAyab.Where(p => p.Month == Month && p.Year == Year && p.UserID == UserID).FirstOrDefault();
            //if (ayabzahab2000 != null)
            //{
            //    ayabzahab2 = db.tbMoadelPadashJarimeAyab.Where(p => p.Month == Month && p.Year == Year && p.UserID == UserID).Select(s => s.AyabOZahab).FirstOrDefault();

            //}
            var tbfkfinancial = db.tbfkfinancial.ToList();
            var FinancialDocuments = db.FinancialDocuments.ToList();
            var fin = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5001).FirstOrDefault();
            var fin2 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5002).FirstOrDefault();
            var find3 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5003).FirstOrDefault();
            var find4 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5004).FirstOrDefault();
            var find5 = tbfkfinancial.Where(p => p.DataDocument == persianDate && p.numbershomar == 5005).FirstOrDefault();
            int ayabzohabnadashtehbashwe = 1;
            var findpymn = db.tbPeymanContracts.Where(s => s.pec_Title == "کارگزاری خدمات توزیع نیروی برق بیرجند").FirstOrDefault();


            if (fin != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == fin.ID).FirstOrDefault();
                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9&&p.usr_ID== UserID).FirstOrDefault();
                if (find != null )
                {
                    if (findsss != null && findpymn != null)
                    {
                        if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                        {
                            find.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                            //ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                        }
                        else
                        {
                            find.Creditor = 0;

                            ayabzahab2 = 0;

                        }

                        //find.Creditor = (float)ayabzahab2;
                        db.SaveChanges();
                    }
                    else if(findpymn == null)
                    {
                        //find.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                        if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                        {
                            find.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                            ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                        }
                        else
                        {
                            find.Creditor = 0;

                            ayabzahab2 = 0;

                        }
                        //find.Creditor = (float)ayabzahab2;
                        db.SaveChanges();
                    }

                }
                else
                {
                    if (findsss != null && findpymn != null)
                    {

                        FinancialDocuments FinancialDocumentss = new FinancialDocuments();
                        if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                        {
                            FinancialDocumentss.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                            //ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                        }
                        else
                        {
                            FinancialDocumentss.Creditor = 0;

                            ayabzahab2 = 0;

                        }
                        FinancialDocumentss.FK_final = fin.ID;

                        FinancialDocumentss.User_ID = UserID;
                        FinancialDocumentss.DataDocument = persianDate;
                        FinancialDocumentss.Debtore = 0;
                        FinancialDocumentss.Title = "خودروشخصی";
                        db.FinancialDocuments.Add(FinancialDocumentss);

                        db.SaveChanges();
                    }
                    else if (findpymn == null)
                    {
                        FinancialDocuments FinancialDocumentss = new FinancialDocuments();
                        if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                        {
                            FinancialDocumentss.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;
                            ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                        }
                        else
                        {
                            FinancialDocumentss.Creditor = 0;

                            ayabzahab2 = 0;

                        }
                        FinancialDocumentss.FK_final = fin.ID;

                        FinancialDocumentss.User_ID = UserID;
                        FinancialDocumentss.DataDocument = persianDate;
                        FinancialDocumentss.Debtore = 0;
                        FinancialDocumentss.Title = "خودروشخصی";
                        db.FinancialDocuments.Add(FinancialDocumentss);
                        db.SaveChanges();
                    }
                    else
                    {

                    }


                }
            }
            else
            {
                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9 && p.usr_ID == UserID).FirstOrDefault();

                if (findsss != null&& findpymn!=null)
                {
                    tbfkfinancial fkfinancial = new tbfkfinancial();
                    fkfinancial.Title = "خودروشخصی";
                    fkfinancial.DataDocument = persianDate;
                    fkfinancial.numbershomar = 5001;
                    db.tbfkfinancial.Add(fkfinancial);
                    db.SaveChanges();

                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();
                    if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                    {
                        FinancialDocumentss.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value; ;
                        //ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                    }
                    else
                    {
                        FinancialDocumentss.Creditor = 0;

                        ayabzahab2 = 0;

                    }
                    FinancialDocumentss.FK_final = fkfinancial.ID;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "خودروشخصی";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    ayabzohabnadashtehbashwe = 2;

                    db.SaveChanges();
                }
                else if (findpymn == null)
                {
                    tbfkfinancial fkfinancial = new tbfkfinancial();
                    fkfinancial.Title = "خودروشخصی";
                    fkfinancial.DataDocument = persianDate;
                    fkfinancial.numbershomar = 5001;
                    db.tbfkfinancial.Add(fkfinancial);
                    db.SaveChanges();

                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();
                    if(moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)") != null)
                    {
                        FinancialDocumentss.Creditor = moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value; ;
                        ayabzahab2 = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "کمک هزینه ایاب و ذهاب(ریال)").mlfvlfsh_Value;

                    }
                    else
                    {
                        FinancialDocumentss.Creditor = 0;

                        ayabzahab2 = 0;

                    }
                    FinancialDocumentss.FK_final = fkfinancial.ID;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;

                    FinancialDocumentss.Title = "خودروشخصی";
                    db.FinancialDocuments.Add(FinancialDocumentss);
                    ayabzohabnadashtehbashwe = 2;

                    db.SaveChanges();
                }
                


            }
            long tablett = 0;
            float mo = 0;
            var ezafggh = db.tbMaxkarkardMonth.Where(p => p.FKUser == UserID && p.Month == Month && p.Year == Year).FirstOrDefault();
            if (ezafggh != null)
            {
                if (ezafggh.Value != null)
                {
                    var f = (ezafggh.Value / 162) + 1;
                    if (moadelkarkar > f)
                    {
                        mo = (float)f;
                    }
                    else
                    {
                        mo = (float)moadelkarkar;
                    }
                }
            }
            float moh = 0;

            if (fin2 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == fin2.ID).FirstOrDefault();
                if (db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 2).FirstOrDefault() != null)
                {
                    tablett = 5000000;
                }
                else if (db.tbUsers.Where(p => p.usr_ID == UserID && p.usr_typeshoghl == 1).FirstOrDefault() != null)
                {
                    tablett = 2000000;
                }
                moh = (float)(tablett * mo);
                if (find != null)
                {
                    find.Creditor = (float)(tablett * mo);
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = fin2.ID;
                    FinancialDocumentss.Creditor = (float)(tablett * mo);
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره تیلت";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره تیلت";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5002;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)(tablett * mo);
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره تیلت";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }



            var py = db.Link_User_And_Peyman.Where(p => p.FK_User_ID == UserID && p.Status == true).FirstOrDefault();

            long mashincol = 0;
            long mashincolsogt = 0;

            long mashincolegareh = 0;
            long mashincouniform = 0;

            int mon = 0;
            int yr = 0;
            if (Month == 1)
            {
                mon = 12;
                yr = Year - 1;
            }
            else
            {
                mon = Month - 1;
                yr = Year;
            }
            var mashin = db.tbReffrenceSave.Where(p => p.FK_PeymanID == py.tbPeymanContracts.pec_ID && p.IsFor == 3).FirstOrDefault();
            if (mashin != null)
            {
                var mashin2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == mashin.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                if (mashin2 != null)
                {

                    var mashin3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == mashin2.ID && p.tbEquipmentMoalefeValue.Any(s => s.Month == mon && s.Year == yr)).FirstOrDefault();
                    if (mashin3 != null)
                    {
                        var mashin4 = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == mashin3.ID && p.CountDays != 0 && p.Fk_user == UserID).ToList();
                        foreach (var item6 in mashin4)
                        {
                            if (item6.CountDays != 0 && item6.CountDays != null)
                            {
                                var rx = db.tbEquipments.Where(p => p.ID == item6.FK_Equipment && p.FK_Peyman == py.tbPeymanContracts.pec_ID).FirstOrDefault();
                                if (rx != null)
                                {
                                    mashincolsogt += (long)(item6.CountDays * (rx.Soght * 12 / 366));
                                    mashincolegareh += (long)(item6.CountDays * (rx.egareh * 12 / 366));
                                    if (rx.uniform != null && rx.uniform != 0)
                                    {
                                        mashincouniform += (long)(item6.CountDays * (rx.uniform * 12 / 366));

                                    }

                                    var c55c = rx.EachValue * 12 / 366;
                                }
                            }

                        }
                    }
                }
            }
            long mohas = mashincouniform + mashincolegareh + mashincolsogt;
            mashincol += mohas;

            //var kasr2 = db.KasriMashin.Where(p => p.Fk_pymn == item.pec_ID && p.number_sorat == number_sorat).FirstOrDefault();
            //if (kasr2 != null && kasr2.Value != null)
            //{
            //    mashincol += (long)kasr2.Value;
            //}

            long Abzarncol = 0;
            var Abzar = db.tbReffrenceSave.Where(p => p.FK_PeymanID == py.tbPeymanContracts.pec_ID && p.IsFor == 4).FirstOrDefault();
            if (Abzar != null)
            {
                var Abzar2 = db.tbReffrenceSaveLevel.Where(p => p.FK_RRSave == Abzar.ID && p.Deleted != true).OrderByDescending(s => s.ID).FirstOrDefault();
                if (Abzar2 != null)
                {
                    var Abzar3 = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == Abzar2.ID /*&& p.Final_Accept == true*/ && p.tbEquipmentMoalefeValue.Any(s => s.Month == mon && s.Year == yr)).FirstOrDefault();
                    if (Abzar3 != null)
                    {
                        var Abzar4 = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == Abzar3.ID && p.CountDays != 0 && p.Fk_user == UserID).ToList();
                        foreach (var item10 in Abzar4)
                        {
                            if (item10.CountDays != null)
                            {
                                var rx = db.tbEquipments.Where(p => p.ID == item10.FK_Equipment && p.FK_Peyman == py.tbPeymanContracts.pec_ID).FirstOrDefault();
                                if (rx != null)
                                {
                                    Abzarncol += (long)(item10.CountDays * Math.Round((double)(rx.EachValue ?? 0) * 12 / 366, 2));

                                    //Abzarncol += (long)(item10.CountDays * Math.Round((double)(rx.EachValue ?? 0) * 12 / 365, 2));
                                    //var c55c = Math.Round((double)(rx.EachValue ?? 0) * 12 / 365, 2);
                                    //var abz = item10.CountDays * c55c;
                                }
                            }

                        }
                    }
                }

            }













            if (find3 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find3.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Creditor = (float)mashincol;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find3.ID;
                    FinancialDocumentss.Creditor = (float)mashincol;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره خودرو شخصیاجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5003;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)mashincol;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره خودرو شخصی(اجاره،سوخت و یونیفرم)";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }
            if (find4 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find4.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Creditor = (float)Abzarncol;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find4.ID;
                    FinancialDocumentss.Creditor = (float)Abzarncol;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Debtore = 0;
                    FinancialDocumentss.Title = "اجاره ابزار";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "اجاره ابزار";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5004;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Creditor = (float)Abzarncol;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Debtore = 0;
                FinancialDocumentss.Title = "اجاره ابزار";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }

            float kosor = 0;
            if (findpymn == null)
            {
                kosor = (float)((Abzarncol + mashincol + ayabzahab2 + moh) * 0.1);

            }
            kosor = 0;
            if (find5 != null)
            {
                var find = FinancialDocuments.Where(p => p.User_ID == UserID && p.FK_final == find5.ID).FirstOrDefault();
                if (find != null)
                {
                    find.Debtore = (float)kosor;
                    find.Creditor = 0;
                    db.SaveChanges();
                }
                else
                {
                    FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                    FinancialDocumentss.FK_final = find5.ID;
                    FinancialDocumentss.Debtore = (float)kosor;
                    FinancialDocumentss.User_ID = UserID;
                    FinancialDocumentss.DataDocument = persianDate;
                    FinancialDocumentss.Creditor = 0;
                    FinancialDocumentss.Title = "کسورات قانونی";
                    db.FinancialDocuments.Add(FinancialDocumentss);

                    db.SaveChanges();

                }
            }
            else
            {
                tbfkfinancial fkfinancial = new tbfkfinancial();
                fkfinancial.Title = "کسورات قانونی";
                fkfinancial.DataDocument = persianDate;
                fkfinancial.numbershomar = 5005;
                db.tbfkfinancial.Add(fkfinancial);

                db.SaveChanges();
                FinancialDocuments FinancialDocumentss = new FinancialDocuments();

                FinancialDocumentss.FK_final = fkfinancial.ID;
                FinancialDocumentss.Debtore = (float)kosor;
                FinancialDocumentss.User_ID = UserID;
                FinancialDocumentss.DataDocument = persianDate;
                FinancialDocumentss.Creditor = 0;
                FinancialDocumentss.Title = "کسورات قانونی";
                db.FinancialDocuments.Add(FinancialDocumentss);

                db.SaveChanges();


            }






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


            //var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 * tedadroozmah);
            //var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 * tedadroozmah);
            //var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 * tedadroozmah);
            //var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 * tedadroozmah);
            var nobatkari1 = (mozdmabna * 0.1) * (tedadrooznobatkari1 * teadadroozkarkard);
            var nobatkari2 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari2 * teadadroozkarkard);
            var nobatkari3 = (mozdmabna * 22.5 / 100) * (tedadrooznobatkari3 * teadadroozkarkard);
            var nobatkari4 = (mozdmabna * 15 / 100) * (tedadrooznobatkari4 * teadadroozkarkard);
            var shabkari = (mozdmabna * 35 / 100) * (tedadroozshabkari * tedadroozmah);
            var eydivapadash = mozdmabna * 60 / 365 * (teadadroozkarkard + tedadroozestelaji);
            var sanavatkhdmat = eydivapadash / 2;
            //var  malf = db.tbMoalefeValueFish.ToList();
            //foreach (var m in malf)
            //{
            //    var ex = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.FK_Moalefe == 1944).FirstOrDefault();
            //    if (ex != null)
            //    {
            //        ex.mlfvlfsh_Value = eydivapadash;
            //        db.SaveChanges();
            //    }
            //}

            if (contractinfo != null)
            {
                // Result.FishValues = new List<FishValue>();
                FishValue fv = new FishValue

                {
                    Title = "مزد سنوات (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane)
                };

                fishValues.Add(fv);

                FishValue nobatkari = new FishValue

                {
                    Title = "نوبتکاری (ریال)",
                    Value = (double)(nobatkari1 + nobatkari2 + nobatkari3 + nobatkari4 + shabkari)
                };

                fishValues.Add(nobatkari);

                FishValue fv85 = new FishValue

                {
                    Title = "مزد سنوات (ریال)",
                    Value = Convert.ToDouble(mozdsanavatroozane) * teadadroozkarkard
                };

                fishValues.Add(fv85);
                FishValue fv2 = new FishValue
                {
                    Title = "مزد شغل (روزانه-ریال)",
                    Value = Convert.ToDouble(mozdshoqlroozane)
                };
                fishValues.Add(fv2);
                FishValue fv200 = new FishValue
                {
                    Title = "مزد گروه (شغل)",
                    Value = Convert.ToDouble(mozdshoqlroozane) * teadadroozkarkard
                };
                fishValues.Add(fv200);
                FishValue fv3 = new FishValue
                {
                    Title = "مزد سایر ( روزانه-ریال)",
                    Value = Convert.ToDouble(mozdsayer)
                };
                fishValues.Add(fv3);
                FishValue fv4 = new FishValue
                {
                    Title = "جمع مزد مبنا ( روزانه-ریال)",
                    Value = mozdmabna
                };
                fishValues.Add(fv4);
                FishValue fv5 = new FishValue
                {
                    Title = "اضافه کار ( هرساعت-ریال)",
                    Value = Math.Round(ezafekar)
                };
                fishValues.Add(fv5);
                FishValue fv6 = new FishValue
                {
                    Title = "کارکرد جمعه (روزانه-ریال)",
                    Value = karkardjome
                };
                fishValues.Add(fv6);
                FishValue fv7 = new FishValue
                {
                    Title = "کارکرد روز تعطیل (روزانه-ریال)",
                    Value = karkarrooztatil
                };
                fishValues.Add(fv7);
                FishValue fv8 = new FishValue
                {
                    Title = "نوبتکاری1 ( صبح وعصر-روزانه-ریال)",
                    Value = (double)nobatkari1
                };
                fishValues.Add(fv8);
                FishValue fv9 = new FishValue
                {
                    Title = "نوبتکاری2 (صبح وشب-روزانه-ریال)",
                    Value = (double)nobatkari2
                };
                fishValues.Add(fv9);
                FishValue fv10 = new FishValue
                {
                    Title = "نوبتکاری3 (عصروشب-روزانه-ریال)",
                    Value = (double)nobatkari3
                };
                fishValues.Add(fv10);
                FishValue fv11 = new FishValue
                {
                    Title = "نوبتکاری4 (صبح وعصروشب-روزانه-ریال)",
                    Value = (double)nobatkari4
                };
                fishValues.Add(fv11);
                FishValue fv12 = new FishValue
                {
                    Title = "شبکاری (هرشب-ریال)",
                    Value = (double)shabkari
                };
                fishValues.Add(fv12);
                FishValue fv13 = new FishValue
                {
                    Title = "عیدی و پاداش ( ماهانه-ریال)",
                    Value = eydivapadash /** nesbatkarbkolmah*/
                };
                fishValues.Add(fv13);
                FishValue fv14 = new FishValue
                {
                    Title = "سنوات خدمت (ماهانه-ریال)",
                    Value = sanavatkhdmat /** nesbatkarbkolmah*/
                };
                fishValues.Add(fv14);
                FishValue fv15 = new FishValue
                {
                    Title = "حق مرخصی (ریال)",
                    Value = haghmorkhasi * nesbatkarbkolmah

                };
                fishValues.Add(fv15);
                FishValue fv16 = new FishValue
                {
                    Title = "کمک هزینه مسکن (ریال)",
                    Value = Convert.ToDouble(haghmaskanbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv16);
                FishValue f17 = new FishValue
                {
                    Title = "کمک هزینه اقلام مصرفی خانوار(ریال)",
                    Value = Convert.ToDouble(aghlammasrafikhanevarbase) * nesbatkarbkolmah
                };
                fishValues.Add(f17);
                FishValue fv18 = new FishValue
                {
                    Title = "کمک هزینه اولاد(ریال)",
                    Value = Convert.ToDouble(hagholadbase) * nesbatkarbkolmah
                };
                fishValues.Add(fv18);

                FishValue fv19 = new FishValue
                {
                    Title = "کسر کار (هرساعت-ریال)",
                    Value = (mozdmabna / 7.3333) * zaribkasrikar
                };
                fishValues.Add(fv19);
                FishValue fv20 = new FishValue
                {
                    Title = "ماموریت ( روزانه-ریال)",
                    Value = (mozdmabna * zaribmamoriat)
                };
                fishValues.Add(fv20);
                FishValue fv21 = new FishValue
                {
                    Title = "غیبت (روزانه-ریال)",
                    Value = (mozdmabna * zaribqeybat)
                };
                fishValues.Add(fv21);


            }
            #endregion
            #region پر کردن هدر فیش

            Result.FishHeader = fishHeader;
            #endregion



            FishValue Mamoriat = new FishValue
            {
                Title = "ماموریت (ریال)",
                Value = CalculateMamoriat((mozdmabna * zaribmamoriat), (double)roozmamoriat)
            };
            fishValues.Add(Mamoriat);
            FishValue ezafkari = new FishValue
            {
                Title = "اضافه کاری (ریال)",
                Value = CalculateEzafkar(ezafekar, tedadsaatezafkar)
            };
            fishValues.Add(ezafkari);

            FishValue jomehkari = new FishValue
            {
                Title = "جمعه کاری (ریال)",
                Value = Calculatejomehkar(karkardjome, teadaroozjome)
            };
            //fishValues.Add(jomehkari);
            var findfkmoalfeh = db.tbContractMoalefeDastmozdi.Where(s => s.md_Title == "جمعه کاری (ریال)").FirstOrDefault();
            if (findfkmoalfeh != null)
            {
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == findfkmoalfeh.md_ID && p.Finalaccept == true);
                if (varObject == null)
                {
                    tbMoalefeValueFish tbMoalefeValueFish12233 = new tbMoalefeValueFish();
                    tbMoalefeValueFish12233.FK_User = UserID;
                    tbMoalefeValueFish12233.mlfvlfsh_Month = Month;
                    tbMoalefeValueFish12233.mlfvlfsh_Year = Year;
                    tbMoalefeValueFish12233.Finalaccept = true;
                    tbMoalefeValueFish12233.FK_Moalefe = findfkmoalfeh.md_ID;

                    // رند کردن مقدار
                    tbMoalefeValueFish12233.mlfvlfsh_Value = Math.Round(jomehkari.Value);

                    db.tbMoalefeValueFish.Add(tbMoalefeValueFish12233);
                    db.SaveChanges();

                }
                else
                {
                    varObject.mlfvlfsh_Value = Math.Round(jomehkari.Value);
                    db.SaveChanges();

                }

            }




            #region پر کردن اضافات و ریختن در فیش ولیو
            var ezafat = db.tbContractMoalefeDastmozdi.Where(p => p.Deduction_or_addition == 2 && p.IsMain != true).ToList();
            List<FishValue> ezafatValue = new List<FishValue>();

            foreach (var item in ezafat)
            {
                if(item.md_ID== 3244)
                {

                }
                if (item.md_ID == 3389)
                {

                }
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID&&p.Finalaccept==true);
                
                FishValue ezafvalue = new FishValue
                {
                    Title = item.md_Title,
                    //Value = (double)db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID).mlfvlfsh_Value
                    Value = varObject != null ? (double)varObject.mlfvlfsh_Value : 0
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
                if (item.md_ID == 3244)
                {

                }
                if (item.md_ID == 3389)
                {

                }
                var var = db.tbMoalefeValueFish.Where(p => p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.mlfvlfsh_Year == Year && p.FK_Moalefe == item.md_ID&&p.Finalaccept==true).FirstOrDefault();
                FishValue kasrvalue = new FishValue
                {
                    Title = item.md_Title,
                    Value = var != null ? (double)var.mlfvlfsh_Value : 0
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
                        var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID&&p.FK_EXCel==null);
                        if (exist == null)//اگر نداشت اضافه کن
                        {
                            if(item.Title== "کمک هزینه ایاب و ذهاب(ریال)")
                            {
                                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9 && p.usr_ID == UserID).FirstOrDefault();
                                if (findsss == null)
                                {
                                    tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                                    {
                                        FK_Moalefe = fkMoalefe,
                                        mlfvlfsh_Submit = qatee,
                                        FK_User = UserID,
                                        mlfvlfsh_Month = Month,
                                        mlfvlfsh_Year = Year,
                                        mlfvlfsh_Value = item.Value

                                    };
                                    moalefeValueFishRepo.Create(moalefefish);

                                }
                            }
                            else
                            {
                                tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                                {
                                    FK_Moalefe = fkMoalefe,
                                    mlfvlfsh_Submit = qatee,
                                    FK_User = UserID,
                                    mlfvlfsh_Month = Month,
                                    mlfvlfsh_Year = Year,
                                    mlfvlfsh_Value = item.Value

                                };
                                moalefeValueFishRepo.Create(moalefefish);
                            }
                        
                        }
                        else if (exist.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            if (item.Title == "کمک هزینه ایاب و ذهاب(ریال)")
                            {
                                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9 && p.usr_ID == UserID).FirstOrDefault();
                                if (findsss == null)
                                {
                                    exist.FK_Moalefe = fkMoalefe;
                                    exist.mlfvlfsh_Submit = qatee;
                                    exist.FK_User = UserID;
                                    exist.mlfvlfsh_Month = Month;
                                    exist.mlfvlfsh_Year = Year;
                                    exist.mlfvlfsh_Value = item.Value;
                                    db.SaveChanges();
                                }

                            }
                            else
                            {
                                exist.FK_Moalefe = fkMoalefe;
                                exist.mlfvlfsh_Submit = qatee;
                                exist.FK_User = UserID;
                                exist.mlfvlfsh_Month = Month;
                                exist.mlfvlfsh_Year = Year;
                                exist.mlfvlfsh_Value = item.Value;
                                db.SaveChanges();
                            }
                              
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception e)
                {

                    transaction.Rollback();
                }
            }


            #endregion




            #region محاسبه بیمه
            var bimelist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 1 || p.included == 3).ToList();
            double jamebime = 0;
            foreach (var item in bimelist)
            {
                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID&&p.FK_EXCel==null);

                double valueToAdd = varObject != null ? (double)varObject.mlfvlfsh_Value : 0;
                jamebime += valueToAdd;

            }
            FishValue valuejamebime = new FishValue
            {
                Title = "جمع کل مشمول بیمه (ریال)",
                Value = Math.Round(jamebime)
            };
            fishValues.Add(valuejamebime);

            FishValue bimesahmkarmand = new FishValue
            {
                Title = "حق بیمه سهم کارمند (ریال)",
                Value = Math.Round(CalculateBime(jamebime, Year))
            };
            fishValues.Add(bimesahmkarmand);
            #endregion



            #region محاسبه مالیات

            int eydivapadashsalane = 0;
            int startmonth = contractinfo.FirstMonthFish + 1 ?? 0; //contractinfo.ShamsiStartTime_month;
            var Moalefeeydivapadash = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == "عیدی و پاداش ( ماهانه-ریال)");
            int moalefeeydivapadashID = 0;
            int countday1 = 0;
            if (Moalefeeydivapadash != null)
            {
                moalefeeydivapadashID = Moalefeeydivapadash.md_ID;
                if (Month == 12)
                {
                    var rozaneh = Convert.ToInt32(mozdsanavatroozane);
                    var shoglroza = Convert.ToInt32(mozdshoqlroozane);
                    eydivapadashsalane =(int) shoglroza + rozaneh;


                    for (int i = 1; i <= 11; i++)
                    {



                        var countday = db.tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == i && s.FK_User == UserID && s.mlfvlfsh_Year == Year && s.FK_Moalefe == 1960 &&s.FK_EXCel==null).FirstOrDefault();
                        if (countday != null)
                        {
                            countday1 += (int)countday.mlfvlfsh_Value;

                        }
                        var countdaystalagi = db. tbMoalefeValueFish.Where(s => s.mlfvlfsh_Month == i && s.FK_User == UserID && s.mlfvlfsh_Year == Year && s.FK_Moalefe == 1964 && s.FK_EXCel == null).FirstOrDefault();
                        if (countdaystalagi != null)
                        {
                            countday1 += (int)countdaystalagi.mlfvlfsh_Value;

                        }
                        

                    }
                    countday1 += teadadroozkarkard;
                    countday1 += (int)tedadroozestelaji;

                    eydivapadashsalane = eydivapadashsalane * 60;
                    eydivapadashsalane = (int)eydivapadashsalane / 365;

                    eydivapadashsalane = eydivapadashsalane * countday1;



                    //for (int i = 0 /*startmonth*/; i < startmonth/*13*/; i++)
                    //{

                    //    eydivapadashsalane += Convert.ToInt32(db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == moalefeeydivapadashID && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == i && p.FK_User == UserID).mlfvlfsh_Value);
                    //}
                }
            }
            var maliatlist = db.tbContractMoalefeDastmozdi.Where(p => p.included == 2 || p.included == 3).ToList();
            double jammaliat = 0;
            foreach (var item in maliatlist)
            {

                var varObject = db.tbMoalefeValueFish.FirstOrDefault(p => p.mlfvlfsh_Year == Year && p.FK_User == UserID && p.mlfvlfsh_Month == Month && p.FK_Moalefe == item.md_ID&&p.FK_EXCel==null);

                if (varObject != null)
                {
                    jammaliat += (double)varObject.mlfvlfsh_Value;
                }
                else
                {
                    jammaliat += 0;

                }
            }
            jammaliat = jammaliat - ((7 / 7) * bimesahmkarmand.Value) - getBimetakmiliValue(Year, Month, UserID);

            FishValue eydivapadashsaliyaneh = new FishValue
            {
                Title = "عیدی و پاداش (سالیانه - ریال)",
                Value = eydivapadashsalane
            };
            fishValues.Add(eydivapadashsaliyaneh);

            
            FishValue valuejammaliat = new FishValue
            {
                Title = "جمع کل مشمول مالیات (ریال)",
                Value = Math.Round(jammaliat + eydivapadashsalane)
            };
            fishValues.Add(valuejammaliat);




            FishValue maliatsahmkarmand = new FishValue
            {
                Title = "مالیات سهم کارمند (ریال)",
                Value = CalculateTaxTajamoetest2(jammaliat, Year, Month, UserID, eydivapadashsalane)
            };
            fishValues.Add(maliatsahmkarmand);
            #endregion

            #region محاسبه اقساط و ریختن در فیش ولیو
            var tamamaqsat = CalculateAqsat(Year, Month, UserID);
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
            //TODO: Felan badan bayad pakshavad
            //FishValue felan = new FishValue
            //{
            //    Title = "صندوق خیریه شرکت",
            //    Value = 200000
            //};
            //aqsatformohasebekasriha.Add(felan);

            Result.FishValuesAqsat = aqsatformohasebekasriha;

            #endregion

            #region قسمت های محاسباتی سمت ویو
            Result.ZakhireKarMazadQabl = CalculateZakhireKarMazadQabl(UserID, Year, Month);
            FishValue ZakhireKarMazadQabl = new FishValue
            {
                Title = "ذخیره کارمازاد قبل (بیمه و مالیات ناپذیر)",
                Value = Result.ZakhireKarMazadQabl
            };
            fishValues.Add(ZakhireKarMazadQabl);
            
            Result.JamNakhalesHoqoqVaMazaya = CalculateJamNakhalesHoqoqVaMazaya(fishValues, lstmoalefeqarardadfish, Result.ZakhireKarMazadQabl, ezafatValue, nesbatkarbkolmah, contractinfo.usc_ID);
            Result.JamNakhalesHoqoqVaMazaya += eydivapadashsalane;

            FishValue JamNakhalesHoqoqVaMazaya = new FishValue
            {
                Title = "جمع ناخالص حقوق و مزایا (ریال)",
                Value = Result.JamNakhalesHoqoqVaMazaya
            };
            fishValues.Add(JamNakhalesHoqoqVaMazaya);
            if (moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار") != null)
            {
                tedadsaatkasrekar = (double)moalefehayeexcel.FirstOrDefault(p => p.tbContractMoalefeDastmozdi.md_Title == "تعداد ساعات کسر کار").mlfvlfsh_Value;
            }
            Result.kasrikarkard = Calculatekasrikarkard((double)tedadsaatkasrekar, (mozdmabna / 7.3333) * zaribkasrikar);
            FishValue kasrikarkard = new FishValue
            {
                Title = "کسری کارکرد (ریال)",
                Value = Result.kasrikarkard
            };
            fishValues.Add(kasrikarkard);
            Result.ZakhireKarMazad = CalculateZakhireKarMazad(CalculatejamkolekosooratbdoonMaxpay(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha), Result.JamNakhalesHoqoqVaMazaya, UserID, Year, Month);
            FishValue ZakhireKarMazad = new FishValue
            {
                Title = "ذخیره کار مازاد (ریال)",
                Value = Result.ZakhireKarMazad
            };
            fishValues.Add(ZakhireKarMazad);
            Result.jamkolekosoorat = Calculatejamkolekosoorat(fishValues, koosratvalue, Result.kasrikarkard, aqsatformohasebekasriha, Result.ZakhireKarMazad);
            FishValue jamkolekosoorat = new FishValue
            {
                Title = "جمع کل کسورات (ریال)",
                Value = Result.jamkolekosoorat
            };
            fishValues.Add(jamkolekosoorat);
            Result.KhalesQabelDaryaft = CalculateKhalesQabelDaryaft(Result.jamkolekosoorat, Result.JamNakhalesHoqoqVaMazaya);
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
                        var exist = db.tbMoalefeValueFish.FirstOrDefault(p => p.FK_Moalefe == fkMoalefe && p.mlfvlfsh_Year == Year && p.mlfvlfsh_Month == Month && p.FK_User == UserID&&p.FK_EXCel==null);
                        if (exist == null)//اگر نداشت اضافه کن
                        {
                            if (item.Title == "کمک هزینه ایاب و ذهاب(ریال)")
                            {
                                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9 && p.usr_ID == UserID).FirstOrDefault();
                                if (findsss == null)
                                {

                                    tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                                    {
                                        FK_Moalefe = fkMoalefe,
                                        mlfvlfsh_Submit = qatee,
                                        FK_User = UserID,
                                        mlfvlfsh_Month = Month,
                                        mlfvlfsh_Year = Year,
                                        mlfvlfsh_Value = item.Value

                                    };
                                    moalefeValueFishRepo.Create(moalefefish);
                                }
                                }
                            else
                            {
                                tbMoalefeValueFish moalefefish = new tbMoalefeValueFish
                                {
                                    FK_Moalefe = fkMoalefe,
                                    mlfvlfsh_Submit = qatee,
                                    FK_User = UserID,
                                    mlfvlfsh_Month = Month,
                                    mlfvlfsh_Year = Year,
                                    mlfvlfsh_Value = item.Value

                                };
                                moalefeValueFishRepo.Create(moalefefish);
                            }
                            
                        }
                        else if (exist.mlfvlfsh_Submit != true)//اگر داشت و فیش قطعی نشده بود ویرایش کن
                        {
                            if (item.Title == "کمک هزینه ایاب و ذهاب(ریال)")
                            {
                                var findsss = db.tbUsers.Where(p => p.usr_typeshoghl != 9 && p.usr_ID == UserID).FirstOrDefault();
                                if (findsss == null)
                                {
                                    exist.FK_Moalefe = fkMoalefe;
                                    exist.mlfvlfsh_Submit = qatee;
                                    exist.FK_User = UserID;
                                    exist.mlfvlfsh_Month = Month;
                                    exist.mlfvlfsh_Year = Year;
                                    exist.mlfvlfsh_Value = item.Value;
                                }
                                }
                            else
                            {
                                exist.FK_Moalefe = fkMoalefe;
                                exist.mlfvlfsh_Submit = qatee;
                                exist.FK_User = UserID;
                                exist.mlfvlfsh_Month = Month;
                                exist.mlfvlfsh_Year = Year;
                                exist.mlfvlfsh_Value = item.Value;

                            }
                               
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



            tbSaleryIsCalculated entity = new tbSaleryIsCalculated
            {
                FKUser = UserID,
                slrcal_Month = Month,
                slrcal_Year = Year
            };
            saleryiscalcRepo.Create(entity);


            return Result;








        }



    }



    public class SQLResponse
    {
        
        public float @ResultPadash { get; set; }
        public float @ResultJarime { get; set; }
        public float @ResultAyabOZahab { get; set; }
        public float @ResultMoadel { get; set; }
       
    }













}
