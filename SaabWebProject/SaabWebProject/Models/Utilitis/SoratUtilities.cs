using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.ViewModels.Salaries.Formula;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace SaabWebProject.Models.Utilitis
{
    public class SoratUtilities
    {
        SaabEntities db;
        tbMoalefeValueSoratRepositories moalefeValueSorathRepo;
        public SoratUtilities(SaabEntities Context)
        {
            db = Context;
            moalefeValueSorathRepo = new tbMoalefeValueSoratRepositories(db);

        }

        public ManualSoratDetail Calculateandinserttotable(int Pyman_ID, int Month, int Year, bool qatee, SoratHeader SoratHeader)
        {
             

            ManualSoratDetail Result2 = new ManualSoratDetail();
            var contractinfo = db.tbPeymanContracts.Where(p => p.pec_ID == Pyman_ID).FirstOrDefault();
            var Avalyenar3 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID);
            var contarctprice = db.tbPeymanContractPrice.Where(p => p.FKPeymanID == Pyman_ID).FirstOrDefault();
        
            var Elhagh=db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman==Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title== "پیشخوان-پیمان-مبلغ الحاقیه").FirstOrDefault();
            var Saba1 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-حقوق").FirstOrDefault();
            var Saba2 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-تبلت").FirstOrDefault();
            var Saba3 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-ایاب و ذهاب").FirstOrDefault();
            var Saba4 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-سوخت").FirstOrDefault();
            var Saba5 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-یونیفرم").FirstOrDefault();
            var Saba6 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-اجاره خودرو سواری").FirstOrDefault();
            var Saba7 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-اجاره خودرو عملیاتی").FirstOrDefault();
            var Saba8 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-بیمه تامین اجتماعی").FirstOrDefault();
            var Saba9 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-بیمه تکمیلی").FirstOrDefault();
            var Saba10 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-سایر").FirstOrDefault();
            var Saba11 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-اعتبار-مالیات حقوق").FirstOrDefault();
            var Saba13 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-ایاب و ذهاب").FirstOrDefault();
            var Saba12 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-حقوق").FirstOrDefault();
            var Saba14= db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-یونیفرم").FirstOrDefault();
            var Saba15 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-سوخت").FirstOrDefault();
            var Saba16 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-اجاره خودرو عملیاتی").FirstOrDefault();
            var Saba17 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-اجاره خودرو سواری").FirstOrDefault();
            var Saba18 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-تبلت").FirstOrDefault();
            var Saba19 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-بیمه تکمیلی").FirstOrDefault();
            var Saba20 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-بیمه تامین اجتماعی").FirstOrDefault();
            var Saba21 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-مالیات حقوق").FirstOrDefault();
            var Saba22 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مرکز هزینه-تحقق-سایر").FirstOrDefault();
            var Saba23 = db.tbMoalefeValuePishkhan.Where(p => p.mlfval_FKPeyman == Pyman_ID && p.tbContractMoalefeDastmozdi.md_Title == "").FirstOrDefault();









            var AvalwyehEner = contarctprice.Count;
            if (AvalwyehEner == null)
            {
                AvalwyehEner = 0;
            }
            var Avalwyehprice = Regex.Replace(contarctprice.Price, ",", "");
            if (Avalwyehprice == null)
            {
                Avalwyehprice = "0";
            }
            var Endprice = Regex.Replace(contractinfo.pec_Price, ",", "");
            if (Endprice == null)
            {
                Endprice = "0";
            }
            var Elhagheprice = Regex.Replace(Elhagh.mlfval_Value, ",", "");
            if (Endprice == null)
            {
                Endprice = "0";
            }
            //var EndEnr= Regex.Replace(contractinfo., ",", "");
            var hoghogsabad1 = Regex.Replace(Saba1.mlfval_Value, ",", "");
            if (hoghogsabad1 == null)
            {
                hoghogsabad1 = "0";
            }
            var tabletsabad2 = Regex.Replace(Saba2.mlfval_Value, ",", "");
            if (tabletsabad2 == null)
            {
                tabletsabad2 = "0";
            }
            var ayabzohabsabad3 = Regex.Replace(Saba3.mlfval_Value, ",", "");
            if (ayabzohabsabad3 == null)
            {
                ayabzohabsabad3 = "0";
            }
            var soghtsabad4 = Regex.Replace(Saba4.mlfval_Value, ",", "");
            if (soghtsabad4 == null)
            {
                soghtsabad4 = "0";
            }
            var uniformsabad5 = Regex.Replace(Saba5.mlfval_Value, ",", "");
            if (uniformsabad5 == null)
            {
                uniformsabad5 = "0";
            }
            var carssabad6 = Regex.Replace(Saba6.mlfval_Value, ",", "");
            if (carssabad6 == null)
            {
                carssabad6 = "0";
            }
            var carsamaliatsabad7 = Regex.Replace(Saba7.mlfval_Value, ",", "");
            if (carsamaliatsabad7 == null)
            {
                carsamaliatsabad7 = "0";
            }
            var bemetamimnsabad8 = Regex.Replace(Saba8.mlfval_Value, ",", "");
            if (bemetamimnsabad8 == null)
            {
                bemetamimnsabad8 = "0";
            }
            var bimetakmilysabad9 = Regex.Replace(Saba9.mlfval_Value, ",", "");
            if (bimetakmilysabad9 == null)
            {
                bimetakmilysabad9 = "0";
            }
            var Sayersabad10 = Regex.Replace(Saba10.mlfval_Value, ",", "");
            if (Sayersabad10 == null)
            {
                Sayersabad10 = "0";
            }
            var malyatsabad11 = Regex.Replace(Saba11.mlfval_Value, ",", "");
            if (malyatsabad11 == null)
            {
                malyatsabad11 = "0";
            }
            var col = Convert.ToInt64(Sayersabad10) + Convert.ToInt64(bimetakmilysabad9) + Convert.ToInt64(bemetamimnsabad8) + Convert.ToInt64(carsamaliatsabad7) + Convert.ToInt64(carssabad6) + Convert.ToInt64(uniformsabad5) + Convert.ToInt64(soghtsabad4) + Convert.ToInt64(ayabzohabsabad3) + Convert.ToInt64(tabletsabad2) + Convert.ToInt64(hoghogsabad1)+ Convert.ToInt64(malyatsabad11);
          //از این جا برای سبد عملکرد 
            var hoghogsabad12 = Regex.Replace(Saba12.mlfval_Value, ",", "");
            if (hoghogsabad12 == null)
            {
                hoghogsabad12 = "0";
            }
            var tabletsabad13 = Regex.Replace(Saba18.mlfval_Value, ",", "");
            if (tabletsabad13 == null)
            {
                tabletsabad13 = "0";
            }
            var ayabzohabsabad14 = Regex.Replace(Saba13.mlfval_Value, ",", "");
            if (ayabzohabsabad14 == null)
            {
                ayabzohabsabad14 = "0";
            }
            var soghtsabad15 = Regex.Replace(Saba15.mlfval_Value, ",", "");
            if (soghtsabad15 == null)
            {
                soghtsabad15 = "0";
            }
            var uniformsabad16 = Regex.Replace(Saba14.mlfval_Value, ",", "");
            if (uniformsabad16 == null)
            {
                uniformsabad16 = "0";
            }

            var carssabad18 = Regex.Replace(Saba17.mlfval_Value, ",", "");
            if (carssabad18 == null)
            {
                carssabad18 = "0";
            }
            var carsamaliatsabad17 = Regex.Replace(Saba16.mlfval_Value, ",", "");
            if (carsamaliatsabad17 == null)
            {
                carsamaliatsabad17 = "0";
            }
            var bemetamimnsabad19 = Regex.Replace(Saba20.mlfval_Value, ",", "");
            if (bemetamimnsabad19 == null)
            {
                bemetamimnsabad19 = "0";
            }
            var bimetakmilysabad20 = Regex.Replace(Saba19.mlfval_Value, ",", "");
            if (bimetakmilysabad20 == null)
            {
                bimetakmilysabad20 = "0";
            }
            var Sayersabad21 = Regex.Replace(Saba22.mlfval_Value, ",", "");
            if (Sayersabad21 == null)
            {
                Sayersabad21 = "0";
            }

            var malyatsabad22 = Regex.Replace(Saba21.mlfval_Value, ",", "");
            if (malyatsabad22 == null)
            {
                malyatsabad22 = "0";
            }




            if (contractinfo != null)
            {
                Result2.SoratValues = new List<SoratValue>();
                SoratValue sor = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-حقوق",
                    Value = Convert.ToDouble(hoghogsabad1)
                };
                Result2.SoratValues.Add(sor);
                SoratValue sor1 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-تبلت",
                    Value = Convert.ToDouble(tabletsabad2)
                };
                Result2.SoratValues.Add(sor1);
                SoratValue sor2 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-ایاب و ذهاب",
                    Value = Convert.ToDouble(ayabzohabsabad3)
                };
                Result2.SoratValues.Add(sor2);
                SoratValue sor3 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-سوخت",
                    Value = Convert.ToDouble(soghtsabad4)
                };
                Result2.SoratValues.Add(sor3);
                SoratValue sor4 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-یونیفرم",
                    Value = Convert.ToDouble(uniformsabad5)
                };
                Result2.SoratValues.Add(sor4);
                SoratValue sor5 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-اجاره خودرو سواری",
                    Value = Convert.ToDouble(carssabad6)
                };
                Result2.SoratValues.Add(sor5);
                SoratValue sor6 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-اجاره خودرو عملیاتی",
                    Value = Convert.ToDouble(carsamaliatsabad7)
                };
                Result2.SoratValues.Add(sor6);
                SoratValue sor7 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-بیمه تکمیلی",
                    Value = Convert.ToDouble(bimetakmilysabad9)
                };
                Result2.SoratValues.Add(sor7);
                SoratValue sor8 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-بیمه تامین اجتماعی",
                    Value = Convert.ToDouble(bemetamimnsabad8)
                };
                Result2.SoratValues.Add(sor8);
                SoratValue sor9 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-سایر",
                    Value = Convert.ToDouble(Sayersabad10)
                };
                Result2.SoratValues.Add(sor9);
                SoratValue sor10 = new SoratValue
                {
                    Title = "پیشخوان-پیمان-مرکز هزینه-اعتبار-مالیات حقوق",
                    Value = Convert.ToDouble(malyatsabad11)
                };
                Result2.SoratValues.Add(sor10);
                SoratValue sor11 = new SoratValue
                {
                    Title = Saba12.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(hoghogsabad12)
                };
                Result2.SoratValues.Add(sor11);
                SoratValue sor12 = new SoratValue
                {
                    Title = Saba18.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(tabletsabad13)
                };
                Result2.SoratValues.Add(sor12);
                SoratValue sor13 = new SoratValue
                {
                    Title = Saba13.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(ayabzohabsabad14)
                };
                Result2.SoratValues.Add(sor13);
                SoratValue sor14 = new SoratValue
                {
                    Title = Saba15.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(soghtsabad15)
                };
                Result2.SoratValues.Add(sor14);
                SoratValue sor15 = new SoratValue
                {
                    Title = Saba14.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(uniformsabad16)
                };
                Result2.SoratValues.Add(sor15);
                SoratValue sor16 = new SoratValue
                {
                    Title = Saba17.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(carssabad18)
                };
                Result2.SoratValues.Add(sor16);
                SoratValue sor17 = new SoratValue
                {
                    Title = Saba16.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(carsamaliatsabad17)
                };
                Result2.SoratValues.Add(sor17);
                SoratValue sor18 = new SoratValue
                {
                    Title = Saba20.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(bemetamimnsabad19)
                };
                Result2.SoratValues.Add(sor18);
                SoratValue sor19 = new SoratValue
                {
                    Title = Saba19.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(bimetakmilysabad20)
                };
                Result2.SoratValues.Add(sor19);
                SoratValue sor20 = new SoratValue
                {
                    Title = Saba22.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(Sayersabad21)
                };
                Result2.SoratValues.Add(sor20);
                SoratValue sor21 = new SoratValue
                {
                    Title = Saba21.tbContractMoalefeDastmozdi.md_Title,
                    Value = Convert.ToDouble(malyatsabad22)
                };
                Result2.SoratValues.Add(sor21);
            }

            Result2.SoratHeader = SoratHeader;

            foreach (var item in Result2.SoratValues)
            {
                tbMoalefeValueSorat moalefefish = new tbMoalefeValueSorat

                {

                    FK_Moalefe = db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_Title == item.Title ).md_ID,
                    mlfvlSorat_Submit = qatee,
                    FK_Pymn = Pyman_ID,
                    mlfvlSorat_Month = Month,
                    mlfvlSorat_Year = Year,
                    mlfvlSoat_Value = item.Value

                };
                moalefeValueSorathRepo.Create(moalefefish);
            }





            return Result2;
        }

















        public string startdate(int Month)
        {
            switch (Month)
            {
                case 1:
                    return "1/1/";

                case 2:
                    return "1/3/";

                case 3:
                    return "1/5/";

                case 4:
                    return "1/7/";

                case 5:
                    return "1/9/";

                case 6:
                    return "1/11/";


                default:
                    return "0";

            }
        }
        public string Endtdate(int Month)
        {
            switch (Month)
            {
                case 1:
                    return "31 / 2 ";

                case 2:
                    return "31 / 4 ";

                case 3:
                    return "31 / 6 ";

                case 4:
                    return "30 / 8 ";

                case 5:
                    return "30 / 10 /";

                case 6:
                    return "29 / 12 / ";


                default:
                    return "0";

            }
        }
    }
}