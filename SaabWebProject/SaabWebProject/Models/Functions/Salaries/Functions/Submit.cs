
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using SaabWebProject.Areas.Salaries.Controllers;
using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using Syncfusion.XlsIO;
using Telerik.Web.Spreadsheet;

namespace SaabWebProject.Models.Functions.Salaries.Functions
{
    public class Submit
    {
        SaabEntities db;
        MoalefeController moalefecntrl;
        tbMoalefeDastmozdiValueFromExcelRepository moalefeexcelRepo;
        tbSavedFunctionsRepositories savedfunctionRepo;
        LogFunction logfunc;

        public Submit(SaabEntities context)
        {
            db = context;
            moalefecntrl = new MoalefeController();
            moalefeexcelRepo = new tbMoalefeDastmozdiValueFromExcelRepository(db);
            savedfunctionRepo = new tbSavedFunctionsRepositories(db);
            logfunc = new LogFunction(db);


        }
        public Workbook SetDataExcel_FunctionFirstPerson(int year, int month, List<int?> usrID, List<int> MoalefeID, DateTime FromDate, DateTime ToDate, string Address)
        {
            var Moalefeexcelfile = Workbook.Load(Address);
            Row Row;


            //fill year
            Row = new Row() { Height = 20, Index = 0 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = year,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);

            //fill month
            Row = new Row() { Height = 20, Index = 1 };
            Row.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value = month,
                    FontFamily = "B Nazanin",
                    Bold = false,
                    Enable = true,
                    Wrap = false,
                    FontSize = 12,
                    Italic = false,
                    Underline = false,
                    Index = 1
                },
            });

            Moalefeexcelfile.Sheets[0].AddRow(Row);


            //باید از کامنت خارج شود
            List<string> MoalefeKarkardi = new List<string>();
            foreach (var item in MoalefeID)
            {
                MoalefeKarkardi.Add(db.tbContractMoalefeDastmozdi.FirstOrDefault(p => p.md_ID == item).md_Title);
            }

            var count = 3;
            foreach (var item in MoalefeKarkardi)
            {
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count
                    },
                });
                count++;
                Moalefeexcelfile.Sheets[0].AddRow(Row);
            }

            //name and lastname and groupjob
            int counter = 3;
            List<tbUsers> UserList = new List<tbUsers>();
            foreach (var item in usrID)
            {
                var user = db.tbUsers.Where(p => p.usr_ID == item).FirstOrDefault();
                UserList.Add(user);
            }

            foreach (var item in UserList)
            {
                //نام ، نام خانوادگی و کد پرسنلی
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.usr_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell()
                        {
                            Value = item.usr_Family,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell()
                        {
                            Value = item.usr_Personal_ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        }
                    });
                    //مقادیر مولفه ها
                    int index = 3;
                    foreach (var item2 in MoalefeKarkardi)
                    {
                        var moalefevalue = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.tbContractMoalefeDastmozdi.md_Title == item2 && p.MoalfeVal_FKUser == item.usr_ID
                        && p.tbSavedFunctions.svdfunc_FromDate <= FromDate && p.tbSavedFunctions.svdfunc_ToDate >= FromDate && p.tbSavedFunctions.svdfunc_FromDate <= ToDate && p.tbSavedFunctions.svdfunc_ToDate >= ToDate).Select(p => p.MoalfeVal_Value).FirstOrDefault();

                        Row.AddCells(new List<Cell>()
                        {
                             new Cell()
                        {
                            Value = moalefevalue,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = index
                        }


                         });
                        index++;
                    }
                    counter++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }




            }

            return Moalefeexcelfile;
        }

        public Workbook SetDataExcel_FunctionOtherPerson(string Address)
        {
            var Functionexcelfile = Workbook.Load(Address);
            return Functionexcelfile;
        }
        public string GetDataFromExcel_FunctionLastOne(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate, int Basteh, int PeymanID)
        {
            string name, family, personalid, year, month = "";
            string value;
            List<string> lstMoalefe = new List<string>();
            int userid = UserStuf.GetOlineUser().usr_ID;
            List<tbMoalefeDastmozdiValueFromExcel> lstMoalefeexcel = new List<tbMoalefeDastmozdiValueFromExcel>();

            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                using (ExcelEngine xl = new ExcelEngine())
                {
                    try
                    {


                        var CheckTypeMoalefe = moalefecntrl.CheckFromDate(FromDate, Basteh);
                        if (CheckTypeMoalefe != "ok")
                        {
                            return CheckTypeMoalefe;
                        }



                        IApplication app = xl.Excel;
                        var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                        var Count_Sheets = workbook.Worksheets.Count;
                        if (Count_Sheets > 0)
                        {
                            var sheet = workbook.Worksheets[0];
                            var Rows = sheet.Rows;
                            int Cols = 0;

                            if (Rows.Length >= 1)
                            {
                                //for (int i = 2; i < Rows.Count; i++)
                                //{
                                //    if (sheet.Rows[i].Cells.Count != sheet.Rows[2].Cells.Count)
                                //    {
                                //        return " خطای ارزیابی در سطر " + (sheet.Rows[i].Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                //    }
                                //}

                                var title = workbook.Worksheets[0].Rows[2].Cells;
                                var List = workbook.Worksheets[0].Rows;
                                var count = workbook.Worksheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }

                                for (int i = 3; i < title.Length; i++)
                                {
                                    lstMoalefe.Add(title[i].Value.ToString());
                                }

                                ///// year
                                var Yearrow = workbook.Worksheets[0].Rows[0];
                                var Year = Yearrow.Cells[1];
                                if (Year.Value != null || Year.Value.ToString() != "")
                                {
                                    year = Year.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون سال  سطر " + " ( " + "0" + " ) " + "را پر کنید ";
                                }

                                ////// month
                                var Monthrow = workbook.Worksheets[0].Rows[1];
                                var Month = Monthrow.Cells[1];
                                if (Month.Value != null || Month.Value.ToString() != "")
                                {
                                    month = Month.Value.ToString();
                                }
                                else
                                {
                                    return " لطفا ستون ماه  سطر " + " ( " + "1" + " ) " + "را پر کنید ";
                                }


                                for (int i = 3; i < count; i++)
                                {
                                    var row = workbook.Worksheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null || Name.Value.ToString() != "")
                                    {
                                        name = Name.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var Family = row.Cells[1];
                                    if (Family.Value != null || Family.Value.ToString() != "")
                                    {
                                        family = Family.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام خانوادگی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var PersonalID = row.Cells[2];
                                    if (PersonalID.Value != null || PersonalID.Value.ToString() != "")
                                    {
                                        personalid = PersonalID.Value.ToString();
                                    }
                                    else
                                    {
                                        return " لطفا ستون کد پرسنلی سطر " + " ( " + i.ToString() + " ) " + "را پر کنید ";
                                    }

                                    var perID = Convert.ToInt32(personalid);
                                    var UserID = db.tbUsers.Where(p => p.usr_Name == name && p.usr_Family == family && p.usr_Personal_ID == perID).FirstOrDefault().usr_ID;
                                    int shomarande = 3;
                                    foreach (var item in lstMoalefe)
                                    {
                                        var Value = row.Cells[shomarande];
                                        if (Value.Value != null)
                                        {
                                            if (Value.Value.ToString() != "")
                                            {
                                                var MoalefeID = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID;


                                                value = Value.Value.ToString();
                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = MoalefeID,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = Convert.ToInt32(month),
                                                    MoalfeVal_Year = Convert.ToInt32(year),
                                                    MoalfeVal_Value = Convert.ToDouble(value),


                                                };
                                                lstMoalefeexcel.Add(dastmozdexcel);


                                            }
                                            else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                            {

                                                tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                                {
                                                    MoalfeVal_FKMoalafeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID,
                                                    MoalfeVal_FKUser = UserID,
                                                    MoalfeVal_Month = Convert.ToInt32(month),
                                                    MoalfeVal_Year = Convert.ToInt32(year),
                                                    MoalfeVal_Value = 0
                                                };
                                                if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                                {
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    return "در ذخیره سازی مشکلی به وجود آمده است";
                                                }
                                            }
                                        }

                                        else //اگر مقداری وارد نکرده بود باید صفر در نظر گرفته شود
                                        {

                                            tbMoalefeDastmozdiValueFromExcel dastmozdexcel = new tbMoalefeDastmozdiValueFromExcel
                                            {
                                                MoalfeVal_FKMoalafeDastmozdi = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title == item).FirstOrDefault().md_ID,
                                                MoalfeVal_FKUser = UserID,
                                                MoalfeVal_Month = Convert.ToInt32(month),
                                                MoalfeVal_Year = Convert.ToInt32(year),
                                                MoalfeVal_Value = 0
                                            };
                                            if (moalefeexcelRepo.Create(dastmozdexcel) == "True")
                                            {
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                return "در ذخیره سازی مشکلی به وجود آمده است";
                                            }
                                        }

                                        shomarande++;
                                    }



                                }

                                var exist = db.tbSavedFunctions.Any(p => p.svdfunc_pymnID == PeymanID && p.svdfunc_BastehID == Basteh && p.svdfunc_FromDate == FromDate && p.svdfunc_ToDate == ToDate);
                                if (exist)//update
                                {
                                    tbSavedFunctions savedfunctions = new tbSavedFunctions();
                                    savedfunctions.svdfunc_BastehID = Basteh;
                                    savedfunctions.svdfunc_FromDate = FromDate;
                                    savedfunctions.svdfunc_ToDate = ToDate;
                                    savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                    savedfunctions.svdfunc_pymnID = PeymanID;
                                    savedfunctions.svdfunc_UserSaveID = userid;
                                    savedfunctions.svdfunc_IsSubmmit = true;

                                    var id = savedfunctionRepo.CreateandReturnID(savedfunctions);

                                    if (id != null)
                                    {
                                        foreach (var item in lstMoalefeexcel)
                                        {
                                            var update = db.tbMoalefeDastmozdiValueFromExcel.Where(p => p.MoalfeVal_FKUser == item.MoalfeVal_FKUser && p.MoalfeVal_Month == item.MoalfeVal_Month && p.MoalfeVal_Year == item.MoalfeVal_Year).FirstOrDefault();
                                            if (update.MoalfeVal_Value != item.MoalfeVal_Value)
                                            {
                                                //InsertInLog
                                                tbLogFunctions tblogfunc = new tbLogFunctions
                                                {
                                                    lgfunc_BastehID = Basteh,
                                                    lgfunc_DateTime = DateTime.Now,
                                                    lgfunc_MoalefeID = item.tbContractMoalefeDastmozdi.md_ID,
                                                    lgfunc_NewValue = item.MoalfeVal_Value,
                                                    lgfunc_OldValue = update.MoalfeVal_Value,
                                                    lgfunc_pymnID = PeymanID,
                                                    lgfunc_UserID = userid
                                                };
                                                var checkInsertInLogFunction = logfunc.InsertToLogFunctionTable(tblogfunc);
                                                update.MoalfeVal_Value = item.MoalfeVal_Value;
                                                update.FK_SavedFunctionsID = id;
                                                db.SaveChanges();
                                            }
                                            else
                                            {
                                                //We do not need to update
                                            }

                                        }
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                        return "در ذخیره سازی مشکلی به وجود آمده است";
                                    }






                                }
                                else//add
                                {
                                    tbSavedFunctions savedfunctions = new tbSavedFunctions();
                                    savedfunctions.svdfunc_BastehID = Basteh;
                                    savedfunctions.svdfunc_FromDate = FromDate;
                                    savedfunctions.svdfunc_ToDate = ToDate;
                                    savedfunctions.svdfunc_SavedDateTime = DateTime.Now;
                                    savedfunctions.svdfunc_pymnID = PeymanID;
                                    savedfunctions.svdfunc_UserSaveID = userid;
                                    savedfunctions.svdfunc_IsSubmmit = false;
                                    savedfunctions.tbMoalefeDastmozdiValueFromExcel = lstMoalefeexcel;

                                    if (savedfunctionRepo.Create(savedfunctions) == "True")
                                    {
                                    }
                                    else
                                    {
                                        transaction.Rollback();
                                        return "در ذخیره سازی مشکلی به وجود آمده است";
                                    }
                                }


                            }
                            else
                            {
                                return "این شیت فاقد سطر می باشد";
                            }
                        }
                        else
                        {
                            return "هیچ شیتی در این اکسل وجود ندارد";
                        }

                        transaction.Commit();
                        return "با موفقیت انجام شد";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
        }
        public string GetDataFromExcel_FunctionOtherOne(HttpPostedFileBase MyExcelStream, string Address)
        {
            try
            {
                MyExcelStream.SaveAs(Address);
                return "1";
            }
            catch (Exception)
            {

                return "0";
            }

        }


    }
}