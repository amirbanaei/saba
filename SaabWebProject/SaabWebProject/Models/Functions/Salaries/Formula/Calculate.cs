using MathParserTK;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using SaabWebProject.Models.Utilitis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace SaabWebProject.Models.Functions.Salaries.Formula
{
    public class Calculate
    {
        SaabEntities db;
        tbMoalefeValueRepository MoalefeValueRepo;
        MathParser Parser = new MathParser();
       static  DateTime datetime = new DateTime();
        public Calculate()
        {
            db = new SaabEntities();
            MoalefeValueRepo = new tbMoalefeValueRepository(db);

        }

        public double CalculateFormula(string Formula)
        {
            var array = Formula.Split('&');
            Formula = "";
            for (int i = 0; i < array.Length; i++)
            {
                Formula += array[i];
            }
            var result = Parser.Parse(Formula);
            return result;
        }
        // in tabe bayad kamel shavad
        public string GetReadyForCalculate(tbFormula formula, int Month, int Year, int UserID)
        {
            datetime = ConvertDateTimeToShamsi.ConvertShamsiYearToYear(Year, Month);

            var result = formula.Frml_Schema;
            //حذف کردن قسمت سمت چپ فرمول
            var xx = result.Split('&');
            result = "";
            for (int i = 1; i < xx.Length; i++)
            {
                result += xx[i] + '&';
            }
            result = result.Substring(0, result.Length - 1);

            result = SigmaFunction(result, UserID);
            result = CountFunction(result, UserID);
            string[] number = formula.Frml_Variablesid.Split(',');
            if (number.Count()>2)
            {
                var variableId = "";
                for (int i = 1; i < number.Length-1; i++)
                {
                    variableId += number[i] + ",";
                }
                result = Moalefe(result, variableId, UserID, Year, Month);

            }


            if (result.Contains("plus"))
            {
                string pattern = "plus";
                string replacement = "+";

                result = Regex.Replace(result, pattern, replacement);
            }
            if (result.Contains("minus"))
            {
                string pattern = "minus";
                string replacement = "-";

                result = Regex.Replace(result, pattern, replacement);
            }
            if (result.Contains("multiplication"))
            {
                string pattern = "multiplication";
                string replacement = "*";

                result = Regex.Replace(result, pattern, replacement);
            }
            if (result.Contains("divide"))
            {
                string pattern = "divide";
                string replacement = "/";

                result = Regex.Replace(result, pattern, replacement);
            }
            //جای گذاری ویرگول

            if (result.Contains("virgol"))
            {
                string pattern = "virgol";
                string replacement = ",";

                result = Regex.Replace(result, pattern, replacement);
            }
            //day 
            result = dayFunction(result);
            //month
            result = monthFunction(result);
            //year
            result = yearFunction(result);
            //hour
            result = hourFunction(result);
            //minute
            result = minuteFunction(result);

            //RoundUP
            result = RoundUPFunction(result);
            //Round
            result = RoundFunction(result);
            //max
            result = MaxFunction(result);
            //
            //min
            result = MinFunction(result);
            //AVG
            result = AVGFunction(result);
            //IF
            result = IFFunction(result);
            //zoj
            result = zojFunction(result);
            //rounddown
            result = RoundDOWNFunction(result);
            //power
            result = PowerFunction(result);


            //
            // جای گذاری سیگما  
            //از کدام جدول بخواند و ماه و سال باید ویرایش شود
            //if (result.Contains("sigma"))
            //{

            //    string[] number = result.Split('s');
            //    for (int i = 0; i < number.Length; i++)
            //    {
            //        if (number[i].Contains("igma"))
            //        {
            //            double value = 0;
            //            var x = number[i].Substring(5, number[i].Length - 5);
            //            var y = x.Split(',');
            //            var moalefeID = Convert.ToInt32(y[2].Substring(3, y[2].Length - 3));
            //            var startDate = Convert.ToInt32(y[0]);
            //            var enddate = Convert.ToInt32(y[1]);
            //            var valuelst = db.tbMoalefeValue.Where(p => p.mlfval_ID == moalefeID && (p.mlfval_Year <= enddate && p.mlfval_Month <= enddate) && (p.mlfval_Year >= startDate && p.mlfval_Month >= startDate)).ToList();
            //            foreach (var item in valuelst)
            //            {
            //                value = (double)item.mlfval_Value + value;
            //            }
            //            var rep = "s" + number[i];
            //            string pattern = @"(^|\s)" + rep + @"(\s|$)";
            //            string replacement = value.ToString();

            //            result = Regex.Replace(result, pattern, replacement);
            //        }
            //    }
            //}
            //        


            return result;


        }
        public string MaxFunction(string phrase)
        {

            int StartMaxIndex = -1;
            int EndMaxIndex = -1;
            while (phrase.Contains("Max"))
            {
                var array = phrase.Split('&');

                //find start index of max
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "Max")
                    {
                        StartMaxIndex = i;
                        break;
                    }
                }
                //find end index of max
                int counter = 0;
                for (int i = StartMaxIndex + 1; i < array.Length; i++)//TODO:y fkr behalesh bayad bokonam age parantez baz dashtim to max
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndMaxIndex = i;
                        break;
                    }
                }
                if (StartMaxIndex != -1 && EndMaxIndex != -1)
                {
                    string Maxreplace = "";
                    string maxforcheck = "";
                    for (int i = StartMaxIndex; i < EndMaxIndex + 1; i++)
                    {
                        Maxreplace += array[i] + '&';

                    }
                    Maxreplace = Maxreplace.Substring(0, Maxreplace.Length - 1);
                    maxforcheck = Maxreplace;

                    maxforcheck = maxforcheck.Substring(4, maxforcheck.Length - 4);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (maxforcheck.Contains("Max"))//Max
                    {
                        maxforcheck = MaxFunction(maxforcheck);
                    }

                    if (maxforcheck.Contains("Min"))//Min
                    {
                        maxforcheck = MinFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("Ave"))//Ave
                    {
                        maxforcheck = AVGFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("RU"))//RoundUP
                    {
                        maxforcheck = RoundUPFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("RD"))//RoundDOWN
                    {
                        maxforcheck = RoundDOWNFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("rn"))//Round
                    {
                        maxforcheck = RoundFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("IF"))//IF
                    {
                        maxforcheck = IFFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("zoj"))//zoj
                    {
                        maxforcheck = zojFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("^"))//Power
                    {
                        maxforcheck = PowerFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("day"))//day
                    {
                        maxforcheck = dayFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("month"))//month
                    {
                        maxforcheck = monthFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("year"))//year
                    {
                        maxforcheck = yearFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("hour"))//hour
                    {
                        maxforcheck = hourFunction(maxforcheck);
                    }
                    if (maxforcheck.Contains("minute"))//minute
                    {
                        maxforcheck = minuteFunction(maxforcheck);
                    }
                    ////
                    maxforcheck = "Max&" + maxforcheck;



                    //joda kardan ghesmat max
                    string Max = "";
                    var array2 = maxforcheck.Split('&');
                    for (int i = 0; i < maxforcheck.Length; i++)
                    {
                        Max += array2[i];

                    }
                    ///////////////////////
                    string maxphrase = Max.Substring(4, Max.Length - 5);
                    var items = maxphrase.Split(',');
                    List<double> maxlist = new List<double>();
                    for (int i = 0; i < items.Length; i++)
                    {
                        var value = Parser.Parse(items[i], false);//inja bayad dorost shavad
                        maxlist.Add(value);
                    }
                    /////                


                    string pattern = Maxreplace;
                    string replacement = maxlist.Max().ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }



            return phrase;
        }

        public string MinFunction(string phrase)
        {

            int StartMinIndex = -1;
            int EndMinIndex = -1;
            while (phrase.Contains("Min"))
            {
                var array = phrase.Split('&');

                //find start index of min
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "Min")
                    {
                        StartMinIndex = i;
                        break;
                    }
                }
                //find end index of min
                int counter = 0;
                for (int i = StartMinIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndMinIndex = i;
                        break;
                    }
                }

                if (StartMinIndex != -1 && EndMinIndex != -1)
                {
                    string Minreplace = "";
                    string minforcheck = "";
                    for (int i = StartMinIndex; i < EndMinIndex + 1; i++)
                    {
                        Minreplace += array[i] + '&';

                    }
                    Minreplace = Minreplace.Substring(0, Minreplace.Length - 1);
                    minforcheck = Minreplace;

                    minforcheck = minforcheck.Substring(4, minforcheck.Length - 4);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (minforcheck.Contains("Max"))//Max
                    {
                        minforcheck = MaxFunction(minforcheck);
                    }

                    if (minforcheck.Contains("Min"))//Min
                    {
                        minforcheck = MinFunction(minforcheck);
                    }
                    if (minforcheck.Contains("Ave"))//Ave
                    {
                        minforcheck = AVGFunction(minforcheck);
                    }
                    if (minforcheck.Contains("RU"))//RoundUP
                    {
                        minforcheck = RoundUPFunction(minforcheck);
                    }
                    if (minforcheck.Contains("RD"))//RoundDOWN
                    {
                        minforcheck = RoundDOWNFunction(minforcheck);
                    }
                    if (minforcheck.Contains("rn"))//Round
                    {
                        minforcheck = RoundFunction(minforcheck);
                    }
                    if (minforcheck.Contains("IF"))//IF
                    {
                        minforcheck = IFFunction(minforcheck);
                    }
                    if (minforcheck.Contains("zoj"))//zoj
                    {
                        minforcheck = zojFunction(minforcheck);
                    }
                    if (minforcheck.Contains("^"))//Power
                    {
                        minforcheck = PowerFunction(minforcheck);
                    }
                    if (minforcheck.Contains("day"))//day
                    {
                        minforcheck = dayFunction(minforcheck);
                    }
                    if (minforcheck.Contains("month"))//month
                    {
                        minforcheck = monthFunction(minforcheck);
                    }
                    if (minforcheck.Contains("year"))//year
                    {
                        minforcheck = yearFunction(minforcheck);
                    }
                    if (minforcheck.Contains("hour"))//hour
                    {
                        minforcheck = hourFunction(minforcheck);
                    }
                    if (minforcheck.Contains("minute"))//minute
                    {
                        minforcheck = minuteFunction(minforcheck);
                    }
                    ////
                    minforcheck = "Min&" + minforcheck;



                    //joda kardan ghesmat max
                    string Min = "";
                    var array2 = minforcheck.Split('&');
                    for (int i = 0; i < minforcheck.Length; i++)
                    {
                        Min += array2[i];

                    }


                    //جایگذاری مقدار مینیمم در فرمول
                    string minphrase = Min.Substring(4, Min.Length - 5);
                    var items = minphrase.Split(',');
                    List<double> minlist = new List<double>();
                    for (int i = 0; i < items.Length; i++)
                    {
                        var value = Parser.Parse(items[i], false);//inja bayad dorost shavad
                        minlist.Add(value);
                    }
                    /////                
                    //////

                    string pattern = Minreplace;
                    string replacement = minlist.Min().ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }

            return phrase;

        }
        public string AVGFunction(string phrase)
        {

            int StartAveIndex = -1;
            int EndAveIndex = -1;
            while (phrase.Contains("Ave"))
            {
                var array = phrase.Split('&');

                //find start index of Ave
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "Ave")
                    {
                        StartAveIndex = i;
                        break;
                    }
                }
                //find end index of Ave
                int counter = 0;
                for (int i = StartAveIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndAveIndex = i;
                        break;
                    }
                }

                if (StartAveIndex != -1 && EndAveIndex != -1)
                {
                    string Avereplace = "";
                    string Aveforcheck = "";
                    for (int i = StartAveIndex; i < EndAveIndex + 1; i++)
                    {
                        Avereplace += array[i] + '&';

                    }
                    Avereplace = Avereplace.Substring(0, Avereplace.Length - 1);
                    Aveforcheck = Avereplace;

                    Aveforcheck = Aveforcheck.Substring(4, Aveforcheck.Length - 4);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (Aveforcheck.Contains("Max"))//Max
                    {
                        Aveforcheck = MaxFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("Min"))//Min
                    {
                        Aveforcheck = MinFunction(Aveforcheck);
                    }

                    if (Aveforcheck.Contains("Ave"))//Ave
                    {
                        Aveforcheck = AVGFunction(Aveforcheck);
                    }

                    if (Aveforcheck.Contains("RU"))//RoundUP
                    {
                        Aveforcheck = RoundUPFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("RD"))//RoundDOWN
                    {
                        Aveforcheck = RoundDOWNFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("rn"))//Round
                    {
                        Aveforcheck = RoundFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("IF"))//IF
                    {
                        Aveforcheck = IFFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("zoj"))//zoj
                    {
                        Aveforcheck = zojFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("^"))//Power
                    {
                        Aveforcheck = PowerFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("day"))//day
                    {
                        Aveforcheck = dayFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("month"))//month
                    {
                        Aveforcheck = monthFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("year"))//year
                    {
                        Aveforcheck = yearFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("hour"))//hour
                    {
                        Aveforcheck = hourFunction(Aveforcheck);
                    }
                    if (Aveforcheck.Contains("minute"))//minute
                    {
                        Aveforcheck = minuteFunction(Aveforcheck);
                    }
                    ////
                    Aveforcheck = "Ave&" + Aveforcheck;



                    //joda kardan ghesmat max
                    string Ave = "";
                    var array2 = Aveforcheck.Split('&');
                    for (int i = 0; i < Aveforcheck.Length; i++)
                    {
                        Ave += array2[i];

                    }


                    string Avephrase = Ave.Substring(4, Ave.Length - 5);
                    var items = Avephrase.Split(',');
                    List<double> Avelist = new List<double>();
                    for (int i = 0; i < items.Length; i++)
                    {
                        var value = Parser.Parse(items[i], false);//inja bayad dorost shavad
                        Avelist.Add(value);
                    }
                    /////                
                    //////

                    string pattern = Avereplace;
                    string replacement = Avelist.Average().ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }

            return phrase;

        }

        public string RoundUPFunction(string phrase)
        {


            int StartRoundUPIndex = -1;
            int EndRoundUPIndex = -1;
            while (phrase.Contains("RU"))
            {
                var array = phrase.Split('&');

                //find start index of min
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "RU")
                    {
                        StartRoundUPIndex = i;
                        break;
                    }
                }
                //find end index of min
                int counter = 0;
                for (int i = StartRoundUPIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndRoundUPIndex = i;
                        break;
                    }
                }


                if (StartRoundUPIndex != -1 && EndRoundUPIndex != -1)
                {
                    string RoundUPreplace = "";
                    string RoundUPphrase = "";
                    for (int i = StartRoundUPIndex; i < EndRoundUPIndex + 1; i++)
                    {
                        RoundUPreplace += array[i] + '&';

                    }
                    RoundUPreplace = RoundUPreplace.Substring(0, RoundUPreplace.Length - 1);
                    RoundUPphrase = RoundUPreplace;
                    RoundUPphrase = RoundUPphrase.Substring(3, RoundUPphrase.Length - 3);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (RoundUPphrase.Contains("Max"))//Max
                    {
                        RoundUPphrase = MaxFunction(RoundUPphrase);
                    }

                    if (RoundUPphrase.Contains("Min"))//Min
                    {
                        RoundUPphrase = MinFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("Ave"))//Ave
                    {
                        RoundUPphrase = AVGFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("RU"))//RoundUP
                    {
                        RoundUPphrase = RoundUPFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("RD"))//RoundDOWN
                    {
                        RoundUPphrase = RoundDOWNFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("rn"))//Round
                    {
                        RoundUPphrase = RoundFunction(RoundUPphrase);
                    }
                   
                    if (RoundUPphrase.Contains("IF"))//IF
                    {
                        RoundUPphrase = IFFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("zoj"))//zoj
                    {
                        RoundUPphrase = zojFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("^"))//Power
                    {
                        RoundUPphrase = PowerFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("day"))//day
                    {
                        RoundUPphrase = dayFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("month"))//month
                    {
                        RoundUPphrase = monthFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("year"))//year
                    {
                        RoundUPphrase = yearFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("hour"))//hour
                    {
                        RoundUPphrase = hourFunction(RoundUPphrase);
                    }
                    if (RoundUPphrase.Contains("minute"))//minute
                    {
                        RoundUPphrase = minuteFunction(RoundUPphrase);
                    }

                    ////
                    RoundUPphrase = "RU&" + RoundUPphrase;


                    string RoundUP = "";
                    var array2 = RoundUPphrase.Split('&');
                    for (int i = 0; i < array2.Length; i++)
                    {
                        RoundUP += array2[i];

                    }



                    string Rounds = RoundUP.Substring(3, RoundUP.Length - 4);

                    var splitround = Rounds.Split(',');
                    string FirstPartRound = splitround[0];
                    string SecondPartRound = splitround[1];
                    FirstPartRound = Parser.Parse(FirstPartRound).ToString();//mathparser
                    SecondPartRound = Parser.Parse(SecondPartRound).ToString();//mathparser
                    var roundUPresult = Math.Ceiling(Convert.ToDouble(FirstPartRound) * Math.Pow(10, Convert.ToInt32(SecondPartRound))) / Math.Pow(10, Convert.ToInt32(SecondPartRound));
                    /////                
                    //////

                    string pattern = RoundUPreplace;
                    string replacement = roundUPresult.ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }


            return phrase;


        }
        public string RoundDOWNFunction(string phrase)
        {


            int StartRoundDOWNIndex = -1;
            int EndRoundDOWNIndex = -1;
            while (phrase.Contains("RD"))
            {
                var array = phrase.Split('&');

                //find start index of min
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "RD")
                    {
                        StartRoundDOWNIndex = i;
                        break;
                    }
                }
                //find end index of min
                int counter = 0;
                for (int i = StartRoundDOWNIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndRoundDOWNIndex = i;
                        break;
                    }
                }


                if (StartRoundDOWNIndex != -1 && EndRoundDOWNIndex != -1)
                {
                    string RoundDOWNreplace = "";
                    string RoundDOWNphrase = "";
                    for (int i = StartRoundDOWNIndex; i < EndRoundDOWNIndex + 1; i++)
                    {
                        RoundDOWNreplace += array[i] + '&';

                    }
                    RoundDOWNreplace = RoundDOWNreplace.Substring(0, RoundDOWNreplace.Length - 1);
                    RoundDOWNphrase = RoundDOWNreplace;
                    RoundDOWNphrase = RoundDOWNphrase.Substring(3, RoundDOWNphrase.Length - 3);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (RoundDOWNphrase.Contains("Max"))//Max
                    {
                        RoundDOWNphrase = MaxFunction(RoundDOWNphrase);
                    }

                    if (RoundDOWNphrase.Contains("Min"))//Min
                    {
                        RoundDOWNphrase = MinFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("Ave"))//Ave
                    {
                        RoundDOWNphrase = AVGFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("RU"))//RoundUP
                    {
                        RoundDOWNphrase = RoundUPFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("RD"))//RoundDOWN
                    {
                        RoundDOWNphrase = RoundDOWNFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("rn"))//Round
                    {
                        RoundDOWNphrase = RoundFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("IF"))//IF
                    {
                        RoundDOWNphrase = IFFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("zoj"))//zoj
                    {
                        RoundDOWNphrase = zojFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("^"))//Power
                    {
                        RoundDOWNphrase = PowerFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("day"))//day
                    {
                        RoundDOWNphrase = dayFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("month"))//month
                    {
                        RoundDOWNphrase = monthFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("year"))//year
                    {
                        RoundDOWNphrase = yearFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("hour"))//hour
                    {
                        RoundDOWNphrase = hourFunction(RoundDOWNphrase);
                    }
                    if (RoundDOWNphrase.Contains("minute"))//minute
                    {
                        RoundDOWNphrase = minuteFunction(RoundDOWNphrase);
                    }
                    ////
                    RoundDOWNphrase = "RD&" + RoundDOWNphrase;


                    string RoundDOWN = "";
                    var array2 = RoundDOWNphrase.Split('&');
                    for (int i = 0; i < array2.Length; i++)
                    {
                        RoundDOWN += array2[i];

                    }



                    string Rounds = RoundDOWN.Substring(3, RoundDOWN.Length - 4);
                    var roundsplit = Rounds.Split(',');
                    string firstPartRound = roundsplit[0];
                    string secondPartRound = roundsplit[1];
                    firstPartRound = Parser.Parse(firstPartRound).ToString();//mathparser
                    secondPartRound = Parser.Parse(secondPartRound).ToString();//mathparser
                    var roundDOWNresult = Math.Floor(Convert.ToDouble(firstPartRound) * Math.Pow(10, Convert.ToInt32(secondPartRound))) / Math.Pow(10, Convert.ToInt32(secondPartRound));
                    /////                
                    //////

                    string pattern = RoundDOWNreplace;
                    string replacement = roundDOWNresult.ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }


            return phrase;


        } 
        public string RoundFunction(string phrase)
        {


            int StartRoundIndex = -1;
            int EndRoundIndex = -1;
            while (phrase.Contains("rn"))
            {
                var array = phrase.Split('&');

                //find start index of min
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "rn")
                    {
                        StartRoundIndex = i;
                        break;
                    }
                }
                //find end index of min
                int counter = 0;
                for (int i = StartRoundIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndRoundIndex = i;
                        break;
                    }
                }


                if (StartRoundIndex != -1 && EndRoundIndex != -1)
                {
                    string Roundreplace = "";
                    string Roundphrase = "";
                    for (int i = StartRoundIndex; i < EndRoundIndex + 1; i++)
                    {
                        Roundreplace += array[i] + '&';

                    }
                    Roundreplace = Roundreplace.Substring(0, Roundreplace.Length - 1);
                    Roundphrase = Roundreplace;
                    Roundphrase = Roundphrase.Substring(3, Roundphrase.Length - 3);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (Roundphrase.Contains("Max"))//Max
                    {
                        Roundphrase = MaxFunction(Roundphrase);
                    }

                    if (Roundphrase.Contains("Min"))//Min
                    {
                        Roundphrase = MinFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("Ave"))//Ave
                    {
                        Roundphrase = AVGFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("RU"))//RoundUP
                    {
                        Roundphrase = RoundUPFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("RD"))//RoundDOWN
                    {
                        Roundphrase = RoundDOWNFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("rn"))//Round
                    {
                        Roundphrase = RoundFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("IF"))//IF
                    {
                        Roundphrase = IFFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("zoj"))//zoj
                    {
                        Roundphrase = zojFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("^"))//Power
                    {
                        Roundphrase = PowerFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("day"))//day
                    {
                        Roundphrase = dayFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("month"))//month
                    {
                        Roundphrase = monthFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("year"))//year
                    {
                        Roundphrase = yearFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("hour"))//hour
                    {
                        Roundphrase = hourFunction(Roundphrase);
                    }
                    if (Roundphrase.Contains("minute"))//minute
                    {
                        Roundphrase = minuteFunction(Roundphrase);
                    }

                    ////
                    Roundphrase = "rn&" + Roundphrase;


                    string Round = "";
                    var array2 = Roundphrase.Split('&');
                    for (int i = 0; i < array2.Length; i++)
                    {
                        Round += array2[i];

                    }



                    string Rounds = Round.Substring(3, Round.Length - 4);
                    var roundsplit = Rounds.Split(',');
                    string firstpartRound = roundsplit[0];
                    string secondpartRound = roundsplit[1];
                    firstpartRound = Parser.Parse(firstpartRound).ToString();//mathparser
                    secondpartRound = Parser.Parse(secondpartRound).ToString();//mathparser
                    var roundresult = Math.Round(Convert.ToDouble(firstpartRound),Convert.ToInt32(secondpartRound));
                    /////                
                    //////

                    string pattern = Roundreplace;
                    string replacement = roundresult.ToString();
                    phrase = phrase.Replace(pattern, replacement);

                }
            }


            return phrase;


        }

        public string IFFunction(string phrase)
        {
            int StartIfIndex = -1;
            int EndIfIndex = -1;
            while (phrase.Contains("IF"))
            {
                var array = phrase.Split('&');

                //find start index of if
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "IF")
                    {
                        StartIfIndex = i;
                        break;
                    }
                }
                //find end index of if
                int counter = 0;
                for (int i = StartIfIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndIfIndex = i;
                        break;
                    }
                }
                if (StartIfIndex != -1 && EndIfIndex != -1)
                {
                    string IFReplace = "";
                    string IFCheck = "";
                    for (int i = StartIfIndex; i < EndIfIndex + 1; i++)
                    {
                        IFReplace += array[i] + '&';

                    }
                    IFReplace = IFReplace.Substring(0, IFReplace.Length - 1);
                    IFCheck = IFReplace;
                    IFCheck = IFCheck.Substring(3, IFCheck.Length - 3);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (IFCheck.Contains("Max"))//Max
                    {
                        IFCheck = MaxFunction(IFCheck);
                    }

                    if (IFCheck.Contains("Min"))//Min
                    {
                        IFCheck = MinFunction(IFCheck);
                    }
                    if (IFCheck.Contains("Ave"))//Ave
                    {
                        IFCheck = AVGFunction(IFCheck);
                    }
                    if (IFCheck.Contains("RU"))//RoundUP
                    {
                        IFCheck = RoundUPFunction(IFCheck);
                    }
                    if (IFCheck.Contains("RD"))//RoundDOWN
                    {
                        IFCheck = RoundDOWNFunction(IFCheck);
                    }
                    if (IFCheck.Contains("rn"))//Round
                    {
                        IFCheck = RoundFunction(IFCheck);
                    }
                    if (IFCheck.Contains("IF"))//IF
                    {
                        IFCheck = IFFunction(IFCheck);
                    }
                    if (IFCheck.Contains("zoj"))//zoj
                    {
                        IFCheck = zojFunction(IFCheck);
                    }
                    if (IFCheck.Contains("^"))//Power
                    {
                        IFCheck = PowerFunction(IFCheck);
                    }
                    if (IFCheck.Contains("day"))//day
                    {
                        IFCheck = dayFunction(IFCheck);
                    }
                    if (IFCheck.Contains("month"))//month
                    {
                        IFCheck = monthFunction(IFCheck);
                    }
                    if (IFCheck.Contains("year"))//year
                    {
                        IFCheck = yearFunction(IFCheck);
                    }
                    if (IFCheck.Contains("hour"))//hour
                    {
                        IFCheck = hourFunction(IFCheck);
                    }
                    if (IFCheck.Contains("minute"))//minute
                    {
                        IFCheck = minuteFunction(IFCheck);
                    }
                    IFCheck = "IF&" + IFCheck;


                    string IF = "";
                    var array2 = IFCheck.Split('&');
                    for (int i = 0; i < array2.Length; i++)
                    {
                        IF += array2[i];

                    }
                    string IFclause = IF.Substring(3, IF.Length - 4);

                    var spilit = IFclause.Split(',');
                    var FirstPartOfIf = spilit[0];
                    var Tru = spilit[1];
                    var Fals = spilit[2];
                    string[] part = null;

                    if (FirstPartOfIf.Contains("SmallerThan"))
                    {
                        FirstPartOfIf = FirstPartOfIf.Replace("SmallerThan", "<");

                        part = FirstPartOfIf.Split('<');
                        //part[0] and part[1] should use mathparser
                        if (Parser.Parse(part[0]) < Parser.Parse(part[1]))
                        {
                            IFclause = Tru;
                        }
                        else
                        {
                            IFclause = Fals;
                        }

                    }
                    if (FirstPartOfIf.Contains("BiggerThan"))
                    {
                        FirstPartOfIf = FirstPartOfIf.Replace("BiggerThan", ">");

                        part = FirstPartOfIf.Split('>');
                        //part[0] and part[1] should use mathparser

                        if (Parser.Parse(part[0]) > Parser.Parse(part[1]))
                        {
                            IFclause = Tru;
                        }
                        else
                        {
                            IFclause = Fals;
                        }

                    }
                    if (FirstPartOfIf.Contains("Contrary"))
                    {
                        FirstPartOfIf = FirstPartOfIf.Replace("Contrary", "!");


                        part = FirstPartOfIf.Split('!');
                        //part[0] and part[1] should use mathparser

                        if (Parser.Parse(part[0]) != Parser.Parse(part[1]))
                        {
                            IFclause = Tru;
                        }
                        else
                        {
                            IFclause = Fals;
                        }

                    }
                    if (FirstPartOfIf.Contains("Equal"))
                    {

                        FirstPartOfIf = FirstPartOfIf.Replace("Equal", "=");

                        part = FirstPartOfIf.Split('=');
                        //part[0] and part[1] should use mathparser

                        if (Parser.Parse(part[0]) == Parser.Parse(part[1]))
                        {
                            IFclause = Tru;
                        }
                        else
                        {
                            IFclause = Fals;
                        }


                    }

                    IFclause = Parser.Parse(IFclause).ToString();
                    string pattern = IFReplace;
                    string replacement = IFclause;
                    phrase = phrase.Replace(pattern, replacement);


                }

            }
            return phrase;
        }
        public string zojFunction(string phrase)
        {
            int StartzojIndex = -1;
            int EndzojIndex = -1;
            while (phrase.Contains("zoj"))
            {
                var array = phrase.Split('&');

                //find start index of if
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] == "zoj")
                    {
                        StartzojIndex = i;
                        break;
                    }
                }
                //find end index of if
                int counter = 0;
                for (int i = StartzojIndex + 1; i < array.Length; i++)//
                {
                    if (array[i] == "(")
                    {
                        counter++;
                    }
                    if (array[i] == ")")
                    {
                        counter--;
                    }
                    if (counter == 0)
                    {
                        EndzojIndex = i;
                        break;
                    }
                }
                if (StartzojIndex != -1 && EndzojIndex != -1)
                {
                    string zojReplace = "";
                    string zojCheck = "";
                    for (int i = StartzojIndex; i < EndzojIndex + 1; i++)
                    {
                        zojReplace += array[i] + '&';

                    }
                    zojReplace = zojReplace.Substring(0, zojReplace.Length - 1);
                    zojCheck = zojReplace;
                    zojCheck = zojCheck.Substring(4, zojCheck.Length - 4);
                    //اگر فرمول های دیگری به کار رفته باشند
                    if (zojCheck.Contains("Max"))//Max
                    {
                        zojCheck = MaxFunction(zojCheck);
                    }

                    if (zojCheck.Contains("Min"))//Min
                    {
                        zojCheck = MinFunction(zojCheck);
                    }
                    if (zojCheck.Contains("Ave"))//Ave
                    {
                        zojCheck = AVGFunction(zojCheck);
                    }
                    if (zojCheck.Contains("RU"))//RoundUP
                    {
                        zojCheck = RoundUPFunction(zojCheck);
                    }
                    if (zojCheck.Contains("RD"))//RoundDOWN
                    {
                        zojCheck = RoundDOWNFunction(zojCheck);
                    }
                    if (zojCheck.Contains("rn"))//Round
                    {
                        zojCheck = RoundFunction(zojCheck);
                    }
                    if (zojCheck.Contains("zoj"))//zoj
                    {
                        zojCheck = zojFunction(zojCheck);
                    }
                    if (zojCheck.Contains("IF"))//IF
                    {
                        zojCheck = IFFunction(zojCheck);
                    }
                    if (zojCheck.Contains("^"))//Power
                    {
                        zojCheck = PowerFunction(zojCheck);
                    }
                    if (zojCheck.Contains("day"))//day
                    {
                        zojCheck = dayFunction(zojCheck);
                    }
                    if (zojCheck.Contains("month"))//month
                    {
                        zojCheck = monthFunction(zojCheck);
                    }
                    if (zojCheck.Contains("year"))//year
                    {
                        zojCheck = yearFunction(zojCheck);
                    }
                    if (zojCheck.Contains("hour"))//hour
                    {
                        zojCheck = hourFunction(zojCheck);
                    }
                    if (zojCheck.Contains("minute"))//minute
                    {
                        zojCheck = minuteFunction(zojCheck);
                    }
                    zojCheck = "zoj&" + zojCheck;


                    string zoj = "";
                    var array2 = zojCheck.Split('&');
                    for (int i = 0; i < array2.Length; i++)
                    {
                        zoj += array2[i];

                    }
                    string zojclause = zoj.Substring(4, zoj.Length - 5);

                    var spilit = zojclause.Split(',');
                    var FirstPart = spilit[0];
                    var Tru = spilit[1];
                    var Fals = spilit[2];
                    if (Parser.Parse(FirstPart) % 2 == 0)
                    {
                        zojclause = Tru;
                    }
                    else
                    {
                        zojclause = Fals;
                    }


                    zojclause = Parser.Parse(zojclause).ToString();
                    string pattern = zojReplace;
                    string replacement = zojclause;
                    phrase = phrase.Replace(pattern, replacement);


                }

            }
            return phrase;
        }

        public string Moalefe(string phrase, string variablesID, int UserID, int Year, int Month)
        {
            // مقداردهی متغییرهای دینامیک
            if (variablesID != null)
            {
                string[] number = variablesID.Split(',');
                for (int i = 0; i < number.Length; i++)
                {
                    double? value = null;
                    if (number[i].Contains("1_"))
                    {
                        var x = number[i].Split('_');
                        var ID = Convert.ToInt32(x[1]);
                            
                        var y= db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_StartTime>=datetime && p.usc_EndTime<=datetime).FirstOrDefault();
                        if(y!=null)
                        {
                            var fkcontract = y.usc_ID;
                            var z = db.tbUserContractsAndMoalefeGhararDadi.Where(p => p.FKMoalefeGhararDadi == ID && p.FKContractID == fkcontract).FirstOrDefault();
                            if(z!=null)
                            {
                                value = (double)z.Value;

                            }
                            else
                            {
                                value = 0;
                            }
                        }
                    }
                    if (number[i].Contains("4_")|| number[i].Contains("5_") || number[i].Contains("6_") || number[i].Contains("8_"))
                    {
                        var x = number[i].Split('_');
                        var ID = Convert.ToInt32(x[1]);
                        var y= db.tbMoalefeValue.Where(p => p.mlfval_FKMoalafeDastmozdi == ID && p.mlfval_Year == Year && p.mlfval_Month == Month && p.mlfval_FKUser == UserID).FirstOrDefault();
                        if(y!=null)
                        {
                            value = (double)y.mlfval_Value;

                        }
                        else
                        {
                            value = 0;
                        }
                    }
                    else if (number[i].Contains("2_") || number[i].Contains("7_"))
                    {
                        var x = number[i].Split('_');
                        var ID = Convert.ToInt32(x[1]);
                        var y= db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == ID && p.MoalfeVal_Year == Year && p.MoalfeVal_Month == Month && p.MoalfeVal_FKUser == UserID).FirstOrDefault();
                        if(y!=null)
                        {
                            value = (double)y.MoalfeVal_Value;

                        }
                        else
                        {
                            value = 0;
                        }
                    }
                    if(value!=null)
                    {                                        
                        string pattern = number[i];
                        string replacement = value.ToString();

                        phrase = Regex.Replace(phrase, pattern, replacement);
                    }
                 
                }
            }
            #region متغییرهای استاتیک قرارداد

            //مقداردهی متغییرهای استاتیک
            var usc = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.usc_StartTime <= datetime && p.usc_EndTime >= datetime).FirstOrDefault();
            if (phrase.Contains("usc_ContractNumber"))
            {
                string pattern = "usc_ContractNumber";
                string replacement = usc.usc_ContractNumber.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_JobCode"))
            {
                string pattern = "usc_JobCode";
                string replacement = usc.usc_JobCode.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("FK_JobGroup"))
            {
                string pattern = "FK_JobGroup";
                string replacement = usc.FK_JobGroup.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_Jobtitle"))
            {
                string pattern = "usc_Jobtitle";
                string replacement = usc.usc_Jobtitle.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_TypeOfContract"))
            {
                string pattern = "usc_TypeOfContract";
                string replacement = usc.usc_TypeOfContract.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_StartTime"))
            {
                string pattern = "usc_StartTime";
                string replacement = usc.usc_StartTime.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_EndTime"))
            {
                string pattern = "usc_EndTime";
                string replacement = usc.usc_EndTime.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usc_TedadSalSanavat"))
            {
                string pattern = "usc_TedadSalSanavat";
                string replacement = usc.usc_TedadSalSanavat.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("jobgroup_ValueMozdGroup"))
            {
                string pattern = "jobgroup_ValueMozdGroup";
                string replacement = usc.jobgroup_ValueMozdGroup.ToString();
                var split = replacement.Split(',');
                if(split.Count()!=0)
                {
                    replacement = "";
                    for (int i = 0; i < split.Length; i++)
                    {
                        replacement += split[i];
                    }
                }            

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("jobgroup_KharoBar"))
            {
                string pattern = "jobgroup_KharoBar";
                string replacement = usc.jobgroup_KharoBar.ToString();
                var split = replacement.Split(',');
                if (split.Count() != 0)
                {
                    replacement = "";
                    for (int i = 0; i < split.Length; i++)
                    {
                        replacement += split[i];
                    }
                }
                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("jobgroup_ValueSanavat"))
            {
                string pattern = "jobgroup_ValueSanavat";
                string replacement = usc.jobgroup_ValueSanavat.ToString();
                var split = replacement.Split(',');
                if (split.Count() != 0)
                {
                    replacement = "";
                    for (int i = 0; i < split.Length; i++)
                    {
                        replacement += split[i];
                    }
                }
                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("jobgroup_HagheMaskan"))
            {
                string pattern = "jobgroup_HagheMaskan";
                string replacement = usc.jobgroup_HagheMaskan.ToString();
                var split = replacement.Split(',');
                if (split.Count() != 0)
                {
                    replacement = "";
                    for (int i = 0; i < split.Length; i++)
                    {
                        replacement += split[i];
                    }
                }
                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("jobgroup_HagheOlad"))
            {
                string pattern = "jobgroup_HagheOlad";
                string replacement = usc.jobgroup_HagheOlad.ToString();
                var split = replacement.Split(',');
                if (split.Count() != 0)
                {
                    replacement = "";
                    for (int i = 0; i < split.Length; i++)
                    {
                        replacement += split[i];
                    }
                }
                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            #endregion
            #region متغییرهای استاتیک کاربری


            var user = db.tbUsers.Where(p => p.usr_ID == UserID).FirstOrDefault();
            if (phrase.Contains("usr_Name"))
            {
                string pattern = "usr_Name";
                string replacement = user.usr_Name.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Family"))
            {
                string pattern = "usr_Family";
                string replacement = user.usr_Family.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_NationalCode"))
            {
                string pattern = "usr_NationalCode";
                string replacement = user.usr_NationalCode.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_FatherName"))
            {
                string pattern = "usr_FatherName";
                string replacement = user.usr_FatherName.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Gender"))
            {
                string pattern = "usr_Gender";
                string replacement = user.usr_Gender.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_MaritaIStatus"))
            {
                string pattern = "usr_MaritaIStatus";
                string replacement = user.usr_MaritaIStatus.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_DateOfBrith"))
            {
                string pattern = "usr_DateOfBrith";
                string replacement = user.usr_DateOfBrith.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_PlaceOfBrith"))
            {
                string pattern = "usr_PlaceOfBrith";
                string replacement = user.usr_PlaceOfBrith.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_PlaceOfIssue"))
            {
                string pattern = "usr_PlaceOfIssue";
                string replacement = user.usr_PlaceOfIssue.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_PhoneNumber"))
            {
                string pattern = "usr_PhoneNumber";
                string replacement = user.usr_PhoneNumber.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_PostalCode"))
            {
                string pattern = "usr_PostalCode";
                string replacement = user.usr_PostalCode.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Address"))
            {
                string pattern = "usr_Address";
                string replacement = user.usr_Address.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_type"))
            {
                string pattern = "usr_type";
                string replacement = user.usr_type.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_SHCode"))
            {
                string pattern = "usr_SHCode";
                string replacement = user.usr_SHCode.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Dutysystem"))
            {
                string pattern = "usr_Dutysystem";
                string replacement = user.usr_Dutysystem.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Email"))
            {
                string pattern = "usr_Email";
                string replacement = user.usr_Email.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_RegistrationNumber"))
            {
                string pattern = "usr_RegistrationNumber";
                string replacement = user.usr_RegistrationNumber.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_PlaceOfRegister"))
            {
                string pattern = "usr_PlaceOfRegister";
                string replacement = user.usr_PlaceOfRegister.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Child_Allowance"))
            {
                string pattern = "usr_Child_Allowance";
                string replacement = user.usr_Child_Allowance.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Personal_ID"))
            {
                string pattern = "usr_Personal_ID";
                string replacement = user.usr_Personal_ID.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Degree"))
            {
                string pattern = "usr_Degree";
                string replacement = user.usr_Degree.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_Grade"))
            {
                string pattern = "usr_Grade";
                string replacement = user.usr_Grade.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_City_Dutysystem"))
            {
                string pattern = "usr_City_Dutysystem";
                string replacement = user.usr_City_Dutysystem.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            if (phrase.Contains("usr_All_Child_Allowance"))
            {
                string pattern = "usr_All_Child_Allowance";
                string replacement = user.usr_All_Child_Allowance.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            #endregion




            //if (phrase.Contains("usc_DurationTime"))
            //{
            //    string pattern = "usc_DurationTime";
            //    string replacement = usc.usc_DurationTime.ToString();

            //    phrase = Regex.Replace(phrase, pattern, replacement);
            //}
            ///
            return phrase;
        }

        public string SigmaFunction(string phrase, int UserID)
        {
            int StartSigmaIndex = -1;
            int EndSigmaIndex = -1;
            var sigmaphrase = "";
            double value = 0;
            while (phrase.Contains("sigma"))
            {

                var array = phrase.Split('&');
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("sigma("))
                    {
                        sigmaphrase = array[i].Substring(6, array[i].Length - 7);
                        break;
                    }
                }
                var spilit = sigmaphrase.Split(',');
                var start = spilit[0];
                var end = spilit[1];
                var moalefe = spilit[2];
                var x = start.Split('/');
                var StartDay = Convert.ToInt32(x[0]);
                var StartMonth = Convert.ToInt32(x[1]);
                var StartYear = Convert.ToInt32(x[2]);
                var y = end.Split('/');
                var EndDay = Convert.ToInt32(y[0]);
                var EndMonth = Convert.ToInt32(y[1]);
                var EndYear = Convert.ToInt32(y[2]);
                if (moalefe.Contains("4_"))
                {
                    var m = moalefe.Split('_');
                    var ID = Convert.ToInt32(m[1]);
                    var lstvalue = db.tbMoalefeValue.Where(p => p.mlfval_FKMoalafeDastmozdi == ID && p.mlfval_Year <= EndYear && p.mlfval_Month <= EndMonth
                    && p.mlfval_Year >= StartYear && p.mlfval_Month >= StartMonth && p.mlfval_FKUser == UserID).Select(p => p.mlfval_Value).ToList();
                    foreach (var item in lstvalue)
                    {
                        value += (double)item;
                    }
                }
                if (moalefe.Contains("2_") || moalefe.Contains("7_"))
                {
                    var m = moalefe.Split('_');
                    var ID = Convert.ToInt32(m[1]);
                    var lstvalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == ID && p.MoalfeVal_Year <= EndYear && p.MoalfeVal_Month <= EndMonth
                    && p.MoalfeVal_Year >= StartYear && p.MoalfeVal_Month >= StartMonth && p.MoalfeVal_FKUser == UserID).Select(p => p.MoalfeVal_Value).ToList();
                    foreach (var item in lstvalue)
                    {
                        value += (double)item;
                    }
                }
                if (moalefe.Contains("1_"))
                {
                    // در دست احداث فعلا نمی خواد
                  //  var m = moalefe.Split('_');
                  //  var ID = Convert.ToInt32(m[1]);
                   // var fkcontract = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.ShamsiStartTime_year >= StartYear && p.ShamsiStartTime_month >= StartMonth && p.ShamsiENDTime_year <= Year && p.ShamsiENDTime_month <= Month).Select(p => p.usc_ID).FirstOrDefault();
                }
                else//static moalefe
                {
                    // در دست احداث فعلا نمی خواد
                    ////  var usc = db.tbUserContracts.Where(p => p.FK_UserID == UserID &&p.ShamsiStartTime_year>=StartYear && p.mo).OrderByDescending(p => p.usc_StartTime).FirstOrDefault();

                    //  switch (moalefe)
                    //  {
                    //      case "usc_CountOfChild":
                    //          break;
                    //      case "usc_TypeOfContract":
                    //          break;
                    //      case "usc_StartTime":
                    //          break;
                    //      case "usc_EndTime":
                    //          break;
                    //      case "usc_DurationTime":
                    //          break;
                    //      default:
                    //          break;
                    //  }
                }
                sigmaphrase = "sigma(" + sigmaphrase + ")";
                string pattern = sigmaphrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }

        public string CountFunction(string phrase, int UserID)
        {
            int StartCountIndex = -1;
            int EndCountIndex = -1;
            var Countphrase = "";
            double value = 0;
            while (phrase.Contains("Count") && !phrase.Contains("child"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("Count("))
                    {
                        Countphrase = array[i].Substring(6, array[i].Length - 7);
                        break;
                    }
                }
                var spilit = Countphrase.Split(',');
                var start = spilit[0];
                var end = spilit[1];
                var moalefe = spilit[2];
                var x = start.Split('/');
                var StartDay = Convert.ToInt32(x[0]);
                var StartMonth = Convert.ToInt32(x[1]);
                var StartYear = Convert.ToInt32(x[2]);
                var y = end.Split('/');
                var EndDay = Convert.ToInt32(y[0]);
                var EndMonth = Convert.ToInt32(y[1]);
                var EndYear = Convert.ToInt32(y[2]);
                if  (moalefe.Contains("4_"))
                {
                    var m = moalefe.Split('_');
                    var ID = Convert.ToInt32(m[1]);
                    value = db.tbMoalefeValue.Where(p => p.mlfval_FKMoalafeDastmozdi == ID && p.mlfval_Year <= EndYear && p.mlfval_Month <= EndMonth
                    && p.mlfval_Year >= StartYear && p.mlfval_Month >= StartMonth && p.mlfval_FKUser == UserID).Select(p => p.mlfval_Value).Count();
                }
                if (moalefe.Contains("2_") || moalefe.Contains("7_"))
                {
                    var m = moalefe.Split('_');
                    var ID = Convert.ToInt32(m[1]);
                    value = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKMoalafeDastmozdi == ID && p.MoalfeVal_Year <= EndYear && p.MoalfeVal_Month <= EndMonth
                    && p.MoalfeVal_Year >= StartYear && p.MoalfeVal_Month >= StartMonth && p.MoalfeVal_FKUser == UserID).Select(p => p.MoalfeVal_Value).Count();
                }
                if (moalefe.Contains("1_"))
                {
                    // در دست احداث فعلا نمی خواد
                    //  var m = moalefe.Split('_');
                    //  var ID = Convert.ToInt32(m[1]);
                    // var fkcontract = db.tbUserContracts.Where(p => p.FK_UserID == UserID && p.ShamsiStartTime_year >= StartYear && p.ShamsiStartTime_month >= StartMonth && p.ShamsiENDTime_year <= Year && p.ShamsiENDTime_month <= Month).Select(p => p.usc_ID).FirstOrDefault();
                }
                else//static moalefe
                {
                    // در دست احداث فعلا نمی خواد
                    ////  var usc = db.tbUserContracts.Where(p => p.FK_UserID == UserID &&p.ShamsiStartTime_year>=StartYear && p.mo).OrderByDescending(p => p.usc_StartTime).FirstOrDefault();

                    //  switch (moalefe)
                    //  {
                    //      case "usc_CountOfChild":
                    //          break;
                    //      case "usc_TypeOfContract":
                    //          break;
                    //      case "usc_StartTime":
                    //          break;
                    //      case "usc_EndTime":
                    //          break;
                    //      case "usc_DurationTime":
                    //          break;
                    //      default:
                    //          break;
                    //  }
                }

                Countphrase = "Count(" + Countphrase + ")";
                string pattern = Countphrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }

        public string PowerFunction(string phrase)
        {
            var Power = "";
            var Payeh = "";

            while (phrase.Contains("^"))
            {
                var array = phrase.Split('&');
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("^"))
                    {
                        Power = array[i + 1];
                        Payeh = array[i - 1];
                        break;
                    }
                }
                var PowerReplace = Payeh + "&^&" + Power;
                //power
                if (Power.Contains("Max"))//Max
                {
                    Power = MaxFunction(Power);
                }

                if (Power.Contains("Min"))//Min
                {
                    Power = MinFunction(Power);
                }
                if (Power.Contains("Ave"))//Ave
                {
                    Power = AVGFunction(Power);
                }
                if (Power.Contains("RU"))//RoundUP
                {
                    Power = RoundUPFunction(Power);
                }
                if (Power.Contains("RD"))//RoundDOWN
                {
                    Power = RoundDOWNFunction(Power);
                }
                if (Power.Contains("rn"))//Round
                {
                    Power = RoundFunction(Power);
                }
                if (Power.Contains("IF"))//IF
                {
                    Power = IFFunction(Power);
                }
                if (Power.Contains("zoj"))//zoj
                {
                    Power = zojFunction(Power);
                }
                if (Power.Contains("^"))//Power
                {
                    Power = PowerFunction(Power);
                }
                if (Power.Contains("day"))//day
                {
                    Power = dayFunction(Power);
                }
                if (Power.Contains("month"))//month
                {
                    Power = monthFunction(Power);
                }
                if (Power.Contains("year"))//year
                {
                    Power = yearFunction(Power);
                }
                if (Power.Contains("hour"))//hour
                {
                    Power = hourFunction(Power);
                }
                if (Power.Contains("minute"))//minute
                {
                    Power = minuteFunction(Power);
                }
                //payeh
                if (Payeh.Contains("Max"))//Max
                {
                    Payeh = MaxFunction(Payeh);
                }

                if (Payeh.Contains("Min"))//Min
                {
                    Payeh = MinFunction(Payeh);
                }
                if (Payeh.Contains("Ave"))//Ave
                {
                    Payeh = AVGFunction(Payeh);
                }
                if (Payeh.Contains("RU"))//RoundUP
                {
                    Payeh = RoundUPFunction(Payeh);
                }
                if (Payeh.Contains("RD"))//RoundDOWN
                {
                    Payeh = RoundDOWNFunction(Payeh);
                }
                if (Payeh.Contains("rn"))//Round
                {
                    Payeh = RoundFunction(Payeh);
                }
                if (Payeh.Contains("IF"))//IF
                {
                    Payeh = IFFunction(Payeh);
                }
                if (Payeh.Contains("zoj"))//zoj
                {
                    Payeh = zojFunction(Payeh);
                }
                if (Payeh.Contains("^"))//Power
                {
                    Payeh = PowerFunction(Payeh);
                }
                if (Payeh.Contains("day"))//day
                {
                    Payeh = dayFunction(Payeh);
                }
                if (Payeh.Contains("month"))//month
                {
                    Payeh = monthFunction(Payeh);
                }
                if (Payeh.Contains("year"))//year
                {
                    Payeh = yearFunction(Payeh);
                }
                if (Payeh.Contains("hour"))//hour
                {
                    Payeh = hourFunction(Payeh);
                }
                if (Payeh.Contains("minute"))//minute
                {
                    Payeh = minuteFunction(Payeh);
                }
                var payehresult = Parser.Parse(Payeh);
                var powerresult = Parser.Parse(Power);
                string powerfuncresult = Math.Pow(payehresult, powerresult).ToString();

                string pattern = PowerReplace;
                string replacement = powerfuncresult;

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }

        public string dayFunction(string phrase)
        {
            double value = 0;
            var dayPhrase = "";
            while (phrase.Contains("day"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("day("))
                    {
                        dayPhrase = array[i].Substring(4, array[i].Length - 5);
                        break;
                    }
                }
                DateTime enteredDate = DateTime.Parse(dayPhrase);
                value = enteredDate.Day;



                dayPhrase = "day(" + dayPhrase + ")";
                string pattern = dayPhrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }
        public string yearFunction(string phrase)
        {
            double value = 0;
            var yearPhrase = "";
            while (phrase.Contains("year"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("year("))
                    {
                        yearPhrase = array[i].Substring(5, array[i].Length - 6);
                        break;
                    }
                }
                DateTime enteredDate = DateTime.Parse(yearPhrase);
                value = enteredDate.Year;



                yearPhrase = "year(" + yearPhrase + ")";
                string pattern = yearPhrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }
        public string monthFunction(string phrase)
        {
            double value = 0;
            var monthPhrase = "";
            while (phrase.Contains("month"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("month("))
                    {
                        monthPhrase = array[i].Substring(6, array[i].Length - 7);
                        break;
                    }
                }
                DateTime enteredDate = DateTime.Parse(monthPhrase);
                value = enteredDate.Month;



                monthPhrase = "month(" + monthPhrase + ")";
                string pattern = monthPhrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }
        public string minuteFunction(string phrase)
        {
            double value = 0;
            var minutePhrase = "";
            while (phrase.Contains("minute"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("minute("))
                    {
                        minutePhrase = array[i].Substring(7, array[i].Length - 8);
                        break;
                    }
                }
                DateTime enteredDate = DateTime.Parse(minutePhrase);
                value = enteredDate.Minute;



                minutePhrase = "minute(" + minutePhrase + ")";
                string pattern = minutePhrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }
        public string hourFunction(string phrase)
        {
            double value = 0;
            var hourPhrase = "";
            while (phrase.Contains("hour"))
            {

                var array = phrase.Split('&');

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].Contains("hour("))
                    {
                        hourPhrase = array[i].Substring(5, array[i].Length - 6);
                        break;
                    }
                }
                DateTime enteredDate = DateTime.Parse(hourPhrase);
                value = enteredDate.Hour;



                hourPhrase = "hour(" + hourPhrase + ")";
                string pattern = hourPhrase;
                string replacement = value.ToString();

                phrase = Regex.Replace(phrase, pattern, replacement);
            }
            return phrase;
        }
    }
}