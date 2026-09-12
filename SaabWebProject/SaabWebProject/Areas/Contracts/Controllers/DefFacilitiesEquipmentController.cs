using SaabWebProject.Areas.Contracts.Models.Classes;
using SaabWebProject.Models.Classes;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Contracts;
using SaabWebProject.Models.ViewModels.Contracts.Features;
using SaabWebProject.Models.ViewModels.Equipment;
using SaabWebProject.Utility;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Telerik.Web.Spreadsheet;
using static SaabWebProject.Areas.Users.Controllers.MessageBoxController;
using SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment;
using SaabWebProject.Models.ViewModels.Contracts.Function;
using ExcelLibrary.BinaryFileFormat;
using System.Web.Services.Description;
using Stimulsoft.Report;
using Header = SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment.Header;
using Stimulsoft.Controls.Win.DotNetBar;
using Microsoft.Ajax.Utilities;
using System.Web.WebSockets;
using Stimulsoft.Svg.FilterEffects;
using Stimulsoft.Blockly.Model;

using SaabWebProject.Models.Utilitis;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using static Stimulsoft.Report.Func;
using static Stimulsoft.Report.StiRecentConnections;
using SaabWebProject.Models.ViewModels.Statements;
using System.Diagnostics.Metrics;
using System.Web.UI;
using System.Threading.Tasks;
using System.Security.AccessControl;
using System.Data.SqlTypes;
using System.Data; // برای DataTable

namespace SaabWebProject.Areas.Contracts.Controllers
{
    public class DefFacilitiesEquipmentController : Controller
    {
        SaabEntities db;

        tbEquipmentGroupRepository EquiprpRepo;
        tbEquipmentBucnhRepository EquipbunchRepo;
        tbEquipmentSpecificationsRepository SpecificationsRepository;
        tbEquipmentSpecificationDataRepository tbEquipmentSpecificationDataRepository;
        tbEquipmentRepository tbequipRepo;
        //tbEquipmentFinantialRepository finantialRepo;
        public DefFacilitiesEquipmentController()
        {
            db = new SaabEntities();
            tbEquipmentSpecificationDataRepository = new tbEquipmentSpecificationDataRepository(db);
            EquiprpRepo = new tbEquipmentGroupRepository(db);
            EquipbunchRepo = new tbEquipmentBucnhRepository(db);
            SpecificationsRepository = new tbEquipmentSpecificationsRepository(db);
            tbequipRepo = new tbEquipmentRepository(db);
            //  finantialRepo = new tbEquipmentFinantialRepository(db);
        }
        // GET: Contracts/DefFacilitiesEquipment
        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }
        [AuthorizeAAA]
        public ActionResult _ListAnavin(int count)
        {
            return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListAnavin.cshtml", count);
        }


        public ActionResult ExportMoalefeExcel7()
        {
            var OutPutFile = SetDataExcel_Moalefe6();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "EquipmentGroup" + extension);
        }
        public ActionResult ExportMoalefeExcel8()
        {
            var OutPutFile = SetDataExcel_Moalefe9();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "EquipmentGroup" + extension);
        }
        public ActionResult ExportMoalefeExcelShahrBahrebardari()
        {
            var OutPutFile = SetDataExcel_MoalefeShahrBahrebardari();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "شهر بهره برداری" + extension);

        }

        public Workbook SetDataExcel_MoalefeShahrBahrebardari()
        {
            Workbook Moalefeexcelfile;
            try
            {
                Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Shahrbahrebardari.xlsx"));
            }
            catch (Exception ex)
            {
                throw new Exception("خطا در بارگذاری فایل Excel: " + ex.Message);
            }
            int rowIndex2 = 2;

            var user = db.dbtarifmahdodayt1.ToList();

            int columnIndex1 = 3;
            var row1 = new Row() { Height = 20, Index = 0 };
            foreach (var item in user)
            {
                row1.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value=item.FK_moalfe,
                FontFamily="B Nazanin",
                Bold=true,
                Enable=true,
                Wrap=false,
                FontSize=12,
                Italic=false,
                Underline=false,
                Index=columnIndex1
            }

        });
                columnIndex1++;
                Moalefeexcelfile.Sheets[0].AddRow(row1);
            }
            int columnIndex12 = 3;

            var row121 = new Row() { Height = 20, Index = 1 };
            foreach (var item in user)
            {

                if (Moalefeexcelfile.Sheets[0].Columns.Count > columnIndex12)
                {
                    Moalefeexcelfile.Sheets[0].Columns[columnIndex12].Width = 100;
                }
                else
                {
                    Moalefeexcelfile.Sheets[0].Columns.Add(new Column
                    {
                        Index = columnIndex12,
                        Width = 100
                    });
                }

                string title = item.tbContractMoalefeDastmozdi.md_Title;
                row121.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value=title,
                FontFamily="B Nazanin",
                Bold=true,
                Enable=true,
                Wrap=false,
                FontSize=12,
                Italic=false,
                Underline=false,
                TextAlign = "center",
                VerticalAlign = "middle",
                Index=columnIndex12
            }

        });
                //row121.Height = CalculateRowHeight(title, 12, 15);
                columnIndex12++;
                Moalefeexcelfile.Sheets[0].AddRow(row121);
            }
            int rowIndex = 0;

            foreach (var item in db.tbPeymanContracts.ToList())
            {
                foreach (var item2 in db.tbpeymancities.Where(p => p.FK_PYMN == item.pec_ID).ToList())
                {

                    Row row22 = new Row() { Height = 20, Index = rowIndex };
                    row22.AddCells(new List<Cell>
{
    new Cell
    {
        Value = item2.tbCities.Name,
        FontFamily = "B Nazanin",
        Bold= false ,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },
    new Cell
    {
        Value = item.pec_Title,
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 1
    }
    ,



});
                    Moalefeexcelfile.Sheets[1].AddRow(row22);
                    rowIndex++;
                }

            }
            var bahre = db.dbtarifmahdodayt3.Where(p => p.name != null).ToList();
            int rowIndex5 = 0;


            foreach (var item in bahre)
            {
                var row25 = new Row() { Height = 20, Index = rowIndex5 };
                row25.AddCells(new List<Cell>()
        {
            new Cell()
            {
                Value=item.name,
                FontFamily="B Nazanin",
                Bold=true,
                Enable=true,
                Wrap=false,
                FontSize=12,
                Italic=false,
                Underline=false,
                Index=2
            }

        });
                rowIndex5++;
                Moalefeexcelfile.Sheets[1].AddRow(row25);
            }

            return Moalefeexcelfile;
        }


        public async Task<ActionResult>  ImportExcel_ShahrBahrebardari(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message =await GetDataFromExcel_ShahrBahrebardari(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }













        public async Task<string> GetDataFromExcel_ShahrBahrebardari(HttpPostedFileBase MyExcelStream)
        {
            //int idbahre = 0;
            var shahr2 = "";
            var peyman2 = "";
            var marahel2 = "";

            //List<dbtarifmahdodayt2> mahdodayt = new List<dbtarifmahdodayt2>();
            List<dbtarifmahdodayt5> mahdodayt2 = new List<dbtarifmahdodayt5>();


            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title2 = workbook.Sheets[0].Rows[0].Cells;

                                //var title = workbook.Sheets[0].Rows[0].Cells;
                                //var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                List<int> IDtitr = new List<int>();
                                for (int i = 0; i < title2.Count; i++)
                                {
                                    IDtitr.Add(int.Parse(title2[i].Value.ToString()));
                                }

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var tbCities = await db.tbCities.ToListAsync();
                                var tbPeymanContracts = await db.tbPeymanContracts.ToListAsync();
                                var dbtarifmahdodayt1 = await db.dbtarifmahdodayt1.ToListAsync();
                                var dbtarifmahdodayt5 = await db.dbtarifmahdodayt5.ToListAsync();

                                var dbtarifmahdodayt3 = await db.dbtarifmahdodayt3.ToListAsync();

                                for (int i = 2; i < count; i++)
                                {

                                    var row = workbook.Sheets[0].Rows[i];
                                    var shar = row.Cells[0];
                                    var peman = row.Cells[1];
                                    var marahel = row.Cells[2];

                                    if (shar.Value != null && peman.Value != null && marahel.Value != null)
                                    {
                                        shahr2 = shar.Value.ToString();
                                        peyman2 = peman.Value.ToString();
                                        marahel2 = marahel.Value.ToString();

                                        //name = Name.Value; // Convert Name.Value to string
                                        //if (!int.TryParse(Name.Value?.ToString(), out name))
                                        //{
                                        //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        //}

                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 3;

                                    if (shomarande < row.Cells.Count)
                                    {
                                        var findShahr = tbCities.Where(c => c.Name == shahr2).FirstOrDefault();
                                        var findPeyman = tbPeymanContracts.Where(p => p.pec_Title == peyman2).FirstOrDefault();
                                        var findMarahel = dbtarifmahdodayt3.Where(s => s.name == marahel2).FirstOrDefault();

                                        foreach (var it in IDtitr)

                                        {
                                            var find3 = dbtarifmahdodayt1.Where(r => r.FK_moalfe == it).FirstOrDefault();
                                            var find44 = dbtarifmahdodayt5.Where(p => p.FK_city == findShahr.ID && p.FK_marhal4 == findMarahel.ID && p.FK_moalf == find3.ID && p.FK_Pymn == findPeyman.pec_ID).FirstOrDefault();
                                            if (find44 == null)
                                            {
                                                var Value = row.Cells[shomarande];
                                                if (Value.Value != null)
                                                {
                                                    dbtarifmahdodayt5 dbtarifmahdodayt = new dbtarifmahdodayt5();
                                                    dbtarifmahdodayt.FK_city = findShahr.ID;
                                                    dbtarifmahdodayt.FK_Pymn = findPeyman.pec_ID;
                                                    dbtarifmahdodayt.FK_marhal4 = findMarahel.ID;
                                                    dbtarifmahdodayt.FK_moalf = find3.ID;
                                                    dbtarifmahdodayt.value = long.Parse(Value.Value.ToString());
                                                    mahdodayt2.Add(dbtarifmahdodayt);
                                                }
                                                else
                                                {
                                                    dbtarifmahdodayt5 dbtarifmahdodayt = new dbtarifmahdodayt5();
                                                    dbtarifmahdodayt.FK_city = findShahr.ID;
                                                    dbtarifmahdodayt.FK_Pymn = findPeyman.pec_ID;
                                                    dbtarifmahdodayt.FK_marhal4 = findMarahel.ID;
                                                    dbtarifmahdodayt.FK_moalf = find3.ID;
                                                    dbtarifmahdodayt.value = 0;
                                                    mahdodayt2.Add(dbtarifmahdodayt);
                                                }
                                            }
                                            else
                                            {
                                                var Value = row.Cells[shomarande];
                                                if (Value.Value != null)
                                                {
                                                 
                                                    find44.value = long.Parse(Value.Value.ToString());
                                                    await db.SaveChangesAsync();
                                                }
                                                else
                                                {
                                                    find44.value = 0;
                                                    await db.SaveChangesAsync();

                                                }
                                            }

                                            shomarande++;
                                        }
                                    }
                                    else
                                    {
                                        return "ایندکس خارج از محدوده در سطر " + (i + 1);
                                    }


                                }
                                db.dbtarifmahdodayt5.AddRange(mahdodayt2);
                                await db.SaveChangesAsync();
                                break;
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

        public async Task<ActionResult>  ExportMoalefeExcelAdamcode()
        {
            var OutPutFile =await SetDataExcel_MoalefeAdamCode();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "آدم کد" + extension);
        }




        public async Task<ActionResult> ImportExcel_taied_sadt_kontorol(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message = await GetDataFromExcel_taied_sadt_kontorol(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }

  public async Task<string> GetDataFromExcel_taied_sadt_kontorol(HttpPostedFileBase MyExcelStream)
        {
            //int idbahre = 0;
            var CodeKargozary2 = 0;
            var CodeShahr2 = 0;
            var CodeBaste2 = 0;

            //List<dbtarifmahdodayt2> mahdodayt = new List<dbtarifmahdodayt2>();
            //List<dbtarifmahdodayt5> mahdodayt2 = new List<dbtarifmahdodayt5>();
            List<tbmarahesabt4> tbmarahesabt = new List<tbmarahesabt4>();


            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title2 = workbook.Sheets[0].Rows[0].Cells;

                                //var title = workbook.Sheets[0].Rows[0].Cells;
                                //var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                List<int> IDtitr = new List<int>();
                                //for (int i = 0; i < title2.Count; i++)
                                //{
                                //    IDtitr.Add(int.Parse(title2[i].Value.ToString()));
                                //}

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                var tbmarahlsabtbastedit2 = await db.tbmarahlsabtbastedit2.ToListAsync();
                                var tbUsers = await db.tbUsers.ToListAsync();
                                var tbmarahesabt41 = await db.tbmarahesabt4.ToListAsync();

                                for (int i = 1; i < count; i++)
                                {

                                    var row = workbook.Sheets[0].Rows[i];
                                    var CodeKargozary = row.Cells[0];
                                    var CodeShahr = row.Cells[1];
                                    var CodeBaste = row.Cells[2];

                                    if (CodeKargozary.Value != null && CodeShahr.Value != null && CodeBaste.Value != null)
                                    {
                                        CodeKargozary2 = int.Parse(CodeKargozary.Value.ToString());
                                        CodeShahr2 = int.Parse(CodeShahr.Value.ToString());
                                        CodeBaste2 = int.Parse(CodeBaste.Value.ToString());
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    int idFk_baste = 0;
                                    var findFkbaste = tbmarahlsabtbastedit2.Where(p => p.FK_City == CodeShahr2 && p.FK_pymn == CodeKargozary2 && p.FK_namebAST == CodeBaste2).FirstOrDefault();
                                    if (findFkbaste != null)
                                    {
                                        idFk_baste = findFkbaste.ID;
                                    }

                                    int shomarande = 6;
                                    var findFkuser = tbUsers.Where(p => p.usr_Personal_ID == int.Parse(row.Cells[9].Value.ToString())).FirstOrDefault();
                                    if (idFk_baste != 0)
                                    {
                                        if (shomarande < row.Cells.Count)
                                        {
                                            if (row.Cells[shomarande].Value != null && row.Cells[shomarande].Value.ToString() == "تایید")
                                            {

                                                var findtaeed = tbmarahesabt41.Where(p => p.FK_basteh == idFk_baste && p.taeed == true).ToList();

                                                foreach (var ut in findtaeed)
                                                {
                                                    ut.del = true;
                                                    await db.SaveChangesAsync();
                                                }
                                                tbmarahesabt4 tbmarahesabt4 = new tbmarahesabt4();
                                                tbmarahesabt4.taeed = true;
                                                tbmarahesabt4.FK_basteh = idFk_baste;
                                                tbmarahesabt4.number = int.Parse(row.Cells[7].Value.ToString());
                                                tbmarahesabt4.FK_usr = findFkuser.usr_ID;
                                                tbmarahesabt4.mohlat = int.Parse(row.Cells[10].Value.ToString());
                                                tbmarahesabt4.Year = int.Parse(row.Cells[11].Value.ToString());
                                                tbmarahesabt4.Month = int.Parse(row.Cells[12].Value.ToString());
                                                var b = int.Parse(row.Cells[13].Value.ToString());
                                                if (b == 1)
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = true;

                                                }
                                                else
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = false;
                                                }
                                                tbmarahesabt.Add(tbmarahesabt4);

                                            }
                                            else if (row.Cells[shomarande].Value != null && row.Cells[shomarande].Value.ToString() == "کنترل")
                                            {
                                                var findtaeed = tbmarahesabt41.Where(p => p.FK_basteh == idFk_baste && p.control == true).ToList();
                                                foreach (var ut in findtaeed)
                                                {
                                                    ut.del = true;
                                                    await db.SaveChangesAsync();
                                                }
                                                tbmarahesabt4 tbmarahesabt4 = new tbmarahesabt4();
                                                tbmarahesabt4.control = true;
                                                tbmarahesabt4.FK_basteh = idFk_baste;
                                                tbmarahesabt4.number = int.Parse(row.Cells[7].Value.ToString());
                                                tbmarahesabt4.FK_usr = findFkuser.usr_ID;
                                                tbmarahesabt4.mohlat = int.Parse(row.Cells[10].Value.ToString());
                                                tbmarahesabt4.Year = int.Parse(row.Cells[11].Value.ToString());
                                                tbmarahesabt4.Month = int.Parse(row.Cells[12].Value.ToString());
                                                var b = int.Parse(row.Cells[13].Value.ToString());
                                                if (b == 1)
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = true;

                                                }
                                                else
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = false;
                                                }
                                                tbmarahesabt.Add(tbmarahesabt4);
                                            }
                                            else if (row.Cells[shomarande].Value != null && row.Cells[shomarande].Value.ToString() == "ثبت")
                                            {
                                                var findtaeed = tbmarahesabt41.Where(p => p.FK_basteh == idFk_baste && p.sabt == true).ToList();
                                                foreach (var ut in findtaeed)
                                                {
                                                    ut.del = true;
                                                    await db.SaveChangesAsync();
                                                }
                                                tbmarahesabt4 tbmarahesabt4 = new tbmarahesabt4();
                                                tbmarahesabt4.sabt = true;
                                                tbmarahesabt4.FK_basteh = idFk_baste;
                                                tbmarahesabt4.number = int.Parse(row.Cells[7].Value.ToString());
                                                tbmarahesabt4.FK_usr = findFkuser.usr_ID;
                                                tbmarahesabt4.day1 = int.Parse(row.Cells[10].Value.ToString());
                                                int dayy = int.Parse(row.Cells[10].Value.ToString()) + int.Parse(row.Cells[11].Value.ToString());
                                                tbmarahesabt4.day2 = dayy;
                                                tbmarahesabt4.Year = int.Parse(row.Cells[12].Value.ToString());
                                                tbmarahesabt4.Month = int.Parse(row.Cells[13].Value.ToString());
                                                var b = int.Parse(row.Cells[14].Value.ToString());
                                                if (b == 1)
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = true;

                                                }
                                                else
                                                {
                                                    tbmarahesabt4.isDarHamanChecked = false;
                                                }
                                                tbmarahesabt.Add(tbmarahesabt4);
                                            }



                                            //if ( row.Cells[shomarande]=="تایید")

                                            //var findShahr = db.tbCities.Where(c => c.Name == shahr2).FirstOrDefault();
                                            //var findPeyman = db.tbPeymanContracts.Where(p => p.pec_Title == peyman2).FirstOrDefault();
                                            //var findMarahel = db.dbtarifmahdodayt3.Where(s => s.name == marahel2).FirstOrDefault();

                                            //foreach (var it in IDtitr)

                                            //{
                                            //    var find3 = db.dbtarifmahdodayt1.Where(r => r.FK_moalfe == it).FirstOrDefault();
                                            //    var Value = row.Cells[shomarande];
                                            //    if (Value.Value != null)
                                            //    {
                                            //        dbtarifmahdodayt5 dbtarifmahdodayt = new dbtarifmahdodayt5();
                                            //        dbtarifmahdodayt.FK_city = findShahr.ID;
                                            //        dbtarifmahdodayt.FK_Pymn = findPeyman.pec_ID;
                                            //        dbtarifmahdodayt.FK_marhal4 = findMarahel.ID;
                                            //        dbtarifmahdodayt.FK_moalf = find3.FK_moalfe;
                                            //        dbtarifmahdodayt.value = long.Parse(Value.Value.ToString());
                                            //        mahdodayt2.Add(dbtarifmahdodayt);
                                            //    }
                                            //    else
                                            //    {
                                            //        dbtarifmahdodayt5 dbtarifmahdodayt = new dbtarifmahdodayt5();
                                            //        dbtarifmahdodayt.FK_city = findShahr.ID;
                                            //        dbtarifmahdodayt.FK_Pymn = findPeyman.pec_ID;
                                            //        dbtarifmahdodayt.FK_marhal4 = findMarahel.ID;
                                            //        dbtarifmahdodayt.FK_moalf = find3.FK_moalfe;
                                            //        dbtarifmahdodayt.value = 0;
                                            //        mahdodayt2.Add(dbtarifmahdodayt);
                                            //    }
                                            //    shomarande++;
                                            //}




                                        }

                                        else
                                        {
                                            return "ایندکس خارج از محدوده در سطر " + (i + 1);
                                        }
                                    }



                                }
                                //db.dbtarifmahdodayt5.AddRange(mahdodayt2);
                                db.tbmarahesabt4.AddRange(tbmarahesabt);
                                await db.SaveChangesAsync();
                                break;
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



        public async Task< Workbook> SetDataExcel_MoalefeAdamCode()
        {


            Workbook Moalefeexcelfile;
            try
            {
                Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/Personsaghf.xlsx"));
            }
            catch (Exception ex)
            {
                throw new Exception("خطا در بارگذاری فایل Excel: " + ex.Message);
            }

            var users = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();

            int rowIndex = 2;

            foreach (var user in users)
            {
                Row row = new Row() { Height = 20, Index = rowIndex };


                row.AddCells(new List<Cell>
    {
        new Cell
        {
            Value = user.FullName,
            FontFamily = "B Nazanin",
            Bold= false ,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },
        new Cell
        {
            Value = user.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
    });


                Moalefeexcelfile.Sheets[0].AddRow(row);
                rowIndex++;

            }
            var user2 = db.dbtarifmahdodayt1.ToList();
            int columnIndex1 = 2;
            var row1 = new Row() { Height = 20, Index = 0 };
            foreach (var item in user2)
            {
                row1.AddCells(new List<Cell>()
            {
                new Cell()
                {
                    Value=item.FK_moalfe,
                    FontFamily="B Nazanin",
                    Bold=true,
                    Enable=true,
                    Wrap=false,
                    FontSize=12,
                    Italic=false,
                    Underline=false,
                    TextAlign = "center",
                    Index=columnIndex1
                }

            });
                columnIndex1++;
                Moalefeexcelfile.Sheets[0].AddRow(row1);
            }

            int columnIndex12 = 2;

            var row121 = new Row() { Height = 20, Index = 1 };

            foreach (var item in user2)
            {
                if (Moalefeexcelfile.Sheets[0].Columns.Count > columnIndex12)
                {
                    Moalefeexcelfile.Sheets[0].Columns[columnIndex12].Width = 100;
                }
                else
                {
                    Moalefeexcelfile.Sheets[0].Columns.Add(new Column
                    {
                        Index = columnIndex12,
                        Width = 100
                    });
                }
                string title = item.tbContractMoalefeDastmozdi.md_Title;
                row121.AddCells(new List<Cell>
{
    new Cell
    {
        Value = title,
        FontFamily = "B Nazanin",
        Bold = true,
        Enable = true,
        Wrap = true,
        FontSize = 12,
        Italic = false,
        Underline = false,
        TextAlign = "center",
        VerticalAlign = "middle",
        Index = columnIndex12
    }
});


                //row121.Height = CalculateRowHeight(title, 12, 15);

                columnIndex12++;
            }


            Moalefeexcelfile.Sheets[0].AddRow(row121);

            return Moalefeexcelfile;
        }



        public async Task< ActionResult> ImportExcel_AdamCode(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message =await GetDataFromExcel_AdamCode(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }




        public async Task<ActionResult> ImportExcel_bime(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message = await GetDataFromExcel_bime(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }
        public async Task<string> GetDataFromExcel_bime(HttpPostedFileBase MyExcelStream)
        {
            db.Configuration.AutoDetectChangesEnabled = false;

            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));

            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var sheet = workbook.Sheets.FirstOrDefault();
                    if (sheet == null)
                        return "هیچ شیتی در این اکسل وجود ندارد";

                    var rows = sheet.Rows;
                    if (rows.Count < 2)
                        return "فایل اکسل فاقد اطلاعات می باشد";
                    ;

                    // Preload ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ
                    var allUsers = await db.tbUsers.ToListAsync();
                    var allMoalefe = await db.tbContractMoalefeDastmozdi.ToListAsync();
                    var allInstallments = await db.tbinstallments.ToListAsync();
                    var allSubsets = await db.tb_Subset_of_installments.ToListAsync();
                    List<tb_Subset_of_installments> tb_Subset_of_installments1 = new List<tb_Subset_of_installments>();
                    // تبدیل لیست‌ها به دیکشنری برای جستجوی بسیار سریع
                    //var usersDict = allUsers.ToDictionary(x => x.usr_Personal_ID, x => x);
                    var usersDict = allUsers
    .GroupBy(x => x.usr_Personal_ID)
    .ToDictionary(g => g.Key, g => g.First());
                    var installmentsDict = allInstallments.ToDictionary(x => x.installment_ID, x => x);

                    for (int i = 1; i < rows.Count; i++)
                    {
                        var row = rows[i];

                        int idkarbar = System.Convert.ToInt32(row.Cells[0].Value);
                        int year = System.Convert.ToInt32(row.Cells[3].Value);
                        int month = System.Convert.ToInt32(row.Cells[4].Value);
                        int meghdareTamin = System.Convert.ToInt32(row.Cells[5].Value);
                        int sum = System.Convert.ToInt32(row.Cells[6].Value);
                        int idmoalefe = System.Convert.ToInt32(row.Cells[8].Value);
                        string sharh = row.Cells[9].Value?.ToString();

                        if (!usersDict.ContainsKey(idkarbar))
                            return $"کاربر با کد پرسنلی {idkarbar} یافت نشد.";

                        var user = usersDict[idkarbar];

                        // پیدا کردن installment
                        var installment = allInstallments
                            .FirstOrDefault(p => p.Fk_molfe == idmoalefe && p.User_ID == user.usr_ID && p.Year == year);

                        if (installment == null)
                        {
                            installment = new tbinstallments
                            {
                                User_ID = user.usr_ID,
                                Fk_molfe = idmoalefe,
                                Total_Amount = sum,
                                Year = year,
                                Month = month,
                                Calculation_type = true,
                                Title = sharh
                            };

                            db.tbinstallments.Add(installment);
                            await db.SaveChangesAsync();

                            allInstallments.Add(installment);
                        }

                        var subset = allSubsets.FirstOrDefault(p =>
                            p.fk_installment_ID == installment.installment_ID &&
                            p.Month == month &&
                            p.Year == year);

                        if (subset == null)
                        {
                            int maxNumber = allSubsets
                                .Where(p => p.fk_installment_ID == installment.installment_ID)
                                .Select(p => p.number)
                                .DefaultIfEmpty(0)
                                .Max();

                            subset = new tb_Subset_of_installments
                            {
                                fk_installment_ID = installment.installment_ID,
                                Month = month,
                                Year = year,
                                value = meghdareTamin,
                                number = maxNumber + 1
                            };

                            tb_Subset_of_installments1.Add(subset);
                            allSubsets.Add(subset);
                        }
                        else
                        {
                            subset.value = meghdareTamin;
                        }
                    }
                    db.tb_Subset_of_installments.AddRange(tb_Subset_of_installments1);
                    await db.SaveChangesAsync();
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

        //public async Task<string> GetDataFromExcel_bime(HttpPostedFileBase MyExcelStream)
        //       {
        //           //int idbahre = 0;
        //           var idkarbar2 = 0;
        //           var year2 = 0;
        //           var month2 = 0;
        //           var sum2 = 0;
        //           var idmoalefe2 = 0;
        //           var meghdartamin2 = 0;
        //           var sharh2 = "";
        //           //List<dbtarifmahdodayt2> mahdodayt = new List<dbtarifmahdodayt2>();
        //           //List<dbtarifmahdodayt5> mahdodayt2 = new List<dbtarifmahdodayt5>();
        //           //List<tbmarahesabt4> tbmarahesabt = new List<tbmarahesabt4>();


        //           var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
        //           using (DbContextTransaction transaction = db.Database.BeginTransaction())
        //           {
        //               try
        //               {
        //                   var Count_Sheets = workbook.Sheets.Count;
        //                   if (Count_Sheets >= 1)
        //                   {
        //                       var sheet = workbook.Sheets[0];
        //                       var Rows = sheet.Rows;
        //                       if (Rows.Count >= 1)
        //                       {
        //                           foreach (var item in Rows)
        //                           {
        //                               if (item.Cells.Count != sheet.Rows[0].Cells.Count)
        //                               {
        //                                   return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
        //                                   ;
        //                               }
        //                               //var title2 = workbook.Sheets[0].Rows[0].Cells;
        //                               var count = workbook.Sheets[0].Rows.Count();
        //                               List<int> IDtitr = new List<int>();
        //                               if (count == 1)
        //                               {
        //                                   return "فایل اکسل فاقد اطلاعات می باشد";
        //                               }

        //                               //var tbmarahlsabtbastedit2 = await db.tbmarahlsabtbastedit2.ToListAsync();

        //                               //var tbmarahesabt41 = await db.tbmarahesabt4.ToListAsync();
        //                               var tbUsers = await db.tbUsers.ToListAsync();
        //                               var tbmoalefe = await db.tbContractMoalefeDastmozdi.ToListAsync();
        //                               for (int i = 1; i < count; i++)
        //                               {

        //                                   var row = workbook.Sheets[0].Rows[i];
        //                                   var idkarbar = row.Cells[0];
        //                                   var year = row.Cells[3];
        //                                   var month = row.Cells[4];
        //                                   var meghdartamin = row.Cells[5];
        //                                   var sum = row.Cells[6];
        //                                   var idmoalefe = row.Cells[8];
        //                                   var sharh = row.Cells[9];
        //                                   //if (CodeKargozary.Value != null && CodeShahr.Value != null && CodeBaste.Value != null)
        //                                   //{
        //                                   //    CodeKargozary2 = int.Parse(CodeKargozary.Value.ToString());
        //                                   //    CodeShahr2 = int.Parse(CodeShahr.Value.ToString());
        //                                   //    CodeBaste2 = int.Parse(CodeBaste.Value.ToString());
        //                                   //}
        //                                   //else
        //                                   //{
        //                                   //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
        //                                   //}
        //                                   idkarbar2 = int.Parse(idkarbar.Value.ToString());
        //                                   year2 = int.Parse(year.Value.ToString());
        //                                   month2 = int.Parse(month.Value.ToString());
        //                                   sum2 = int.Parse(sum.Value.ToString());
        //                                   idmoalefe2 = int.Parse(idmoalefe.Value.ToString());
        //                                   meghdartamin2 = int.Parse(meghdartamin.Value.ToString());
        //                                   sharh2 = sharh.Value.ToString();

        //                                   var findkarbar = tbUsers.Where(p => p.usr_Personal_ID == idkarbar2).SingleOrDefault();
        //                                   tbinstallments tbinstallments2 = new tbinstallments();

        //                                   var tbinstallments = db.tbinstallments.Where(p => p.Fk_molfe == idmoalefe2 && p.User_ID == findkarbar.usr_ID && p.Year == year2).FirstOrDefault();
        //                                   if (tbinstallments == null)
        //                                   {
        //                                       tbinstallments2.User_ID = findkarbar.usr_ID;
        //                                       tbinstallments2.Fk_molfe = idmoalefe2;
        //                                       tbinstallments2.Total_Amount = sum2;
        //                                       tbinstallments2.Year = year2;
        //                                       tbinstallments2.Month = month2;
        //                                       tbinstallments2.Calculation_type = true;
        //                                       tbinstallments2.Title = sharh2;

        //                                       db.tbinstallments.Add(tbinstallments2);
        //                                       await db.SaveChangesAsync();
        //                                       var tbinstallments3 = db.tbinstallments.Where(p => p.Fk_molfe == idmoalefe2 && p.User_ID == findkarbar.usr_ID && p.Year == year2 && p.Total_Amount == sum2).FirstOrDefault();
        //                                       var idinstal = tbinstallments3.installment_ID;
        //                                       var sdf = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == tbinstallments2.installment_ID && p.Month == month2 && p.Year == year2).FirstOrDefault();
        //                                       var sdf22 = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == tbinstallments2.installment_ID).OrderByDescending(s => s.Subsetins_ID).FirstOrDefault();
        //                                       int number = 1;
        //                                       if (sdf22 != null)
        //                                       {
        //                                           number = sdf22.number + 1;
        //                                       }
        //                                       if (sdf != null)
        //                                       {
        //                                           sdf.value = meghdartamin2;
        //                                           await db.SaveChangesAsync();
        //                                       }
        //                                       else
        //                                       {
        //                                           tb_Subset_of_installments tb_Subset_of_installments1 = new tb_Subset_of_installments();
        //                                           //tb_Subset_of_installments1.fk_installment_ID=idinstal;

        //                                           tb_Subset_of_installments1.value = meghdartamin2;
        //                                           tb_Subset_of_installments1.Year = year2;
        //                                           tb_Subset_of_installments1.Month = month2;
        //                                           tb_Subset_of_installments1.fk_installment_ID = tbinstallments2.installment_ID;
        //                                           tb_Subset_of_installments1.number = number;

        //                                           db.tb_Subset_of_installments.Add(tb_Subset_of_installments1);
        //                                           await db.SaveChangesAsync();
        //                                       }

        //                                   }
        //                                   else
        //                                   {
        //                                       tbinstallments2 = tbinstallments;

        //                                       var sdf = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == tbinstallments2.installment_ID && p.Month == month2 && p.Year == year2).FirstOrDefault();
        //                                       var sdf22 = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == tbinstallments2.installment_ID ).OrderByDescending(s=>s.Subsetins_ID).FirstOrDefault();
        //                                       int number = 1;
        //                                            if (sdf22 != null)
        //                                       {
        //                                           number = sdf22.number + 1;
        //                                       }
        //                                       if (sdf != null)
        //                                       {
        //                                           sdf.value = meghdartamin2;
        //                                           await db.SaveChangesAsync();
        //                                       }

        //                                       else
        //                                       {
        //                                           tb_Subset_of_installments tb_Subset_of_installments1 = new tb_Subset_of_installments();
        //                                           //tb_Subset_of_installments1.fk_installment_ID=idinstal;

        //                                           tb_Subset_of_installments1.value = meghdartamin2;
        //                                           tb_Subset_of_installments1.Year = year2;
        //                                           tb_Subset_of_installments1.Month = month2;
        //                                           tb_Subset_of_installments1.number = number;

        //                                           tb_Subset_of_installments1.fk_installment_ID = tbinstallments2.installment_ID;

        //                                           db.tb_Subset_of_installments.Add(tb_Subset_of_installments1);
        //                                           await db.SaveChangesAsync();
        //                                       }
        //                                   }







        //                               }
        //                               //db.dbtarifmahdodayt5.AddRange(mahdodayt2);
        //                               //db.tbmarahesabt4.AddRange(tbmarahesabt);
        //                               //await db.SaveChangesAsync();
        //                               break;
        //                           }
        //                       }
        //                       else
        //                       {
        //                           return "این شیت فاقد سطر می باشد";
        //                       }
        //                   }
        //                   else
        //                   {
        //                       return "هیچ شیتی در این اکسل وجود ندارد";
        //                   }
        //                   transaction.Commit();
        //                   return "با موفقیت انجام شد";
        //               }
        //               catch (Exception ex)
        //               {
        //                   transaction.Rollback();
        //                   return ex.Message;

        //               }
        //           }
        //       }
        public async Task<ActionResult> ImportExcel_moadelkar(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message = await GetDataFromExcel_moadelkar(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }




        public async Task<string> GetDataFromExcel_moadelkar(HttpPostedFileBase MyExcelStream)
        {

            var idkarbar2 = 0;
            double moadelkarkard2 = 0;
            long ayabzahab2 = 0;
            var year2 = 0;
            var month2 = 0;
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    List<tbMoadelPadashJarimeAyab> tbMoadelPadashJarimeAyab11 = new List<tbMoadelPadashJarimeAyab>();
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }

                                var count = workbook.Sheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                var tbUsers = await db.tbUsers.ToListAsync();
                                for (int i = 2; i < count; i++)
                                {

                                    var row = workbook.Sheets[0].Rows[i];
                                    var idkarbar = row.Cells[0];
                                    var moadelkarkard = row.Cells[6];
                                    var ayabzahab = row.Cells[7];
                                    var month = row.Cells[8];
                                    var year = row.Cells[9];
                                    idkarbar2 = int.Parse(idkarbar.Value.ToString());
                                    year2 = int.Parse(year.Value.ToString());
                                    month2 = int.Parse(month.Value.ToString());
                                    moadelkarkard2 = double.Parse(moadelkarkard.Value.ToString());
                                    ayabzahab2 = long.Parse(ayabzahab.Value.ToString());


                                    var findkarbar = tbUsers.Where(p => p.usr_Personal_ID == idkarbar2).SingleOrDefault();
                                    tbMoadelPadashJarimeAyab tbMoadelPadashJarimeAyab1 = new tbMoadelPadashJarimeAyab();


                                    if (findkarbar != null)
                                    {
                                        tbMoadelPadashJarimeAyab1.UserID = findkarbar.usr_ID;
                                        tbMoadelPadashJarimeAyab1.Year = year2;
                                        tbMoadelPadashJarimeAyab1.Month = month2;
                                        tbMoadelPadashJarimeAyab1.AyabOZahab = ayabzahab2;
                                        tbMoadelPadashJarimeAyab1.Moadel = moadelkarkard2;
                                        tbMoadelPadashJarimeAyab1.Padash = 0;
                                        tbMoadelPadashJarimeAyab1.Jarime = 0;
                                        tbMoadelPadashJarimeAyab11.Add(tbMoadelPadashJarimeAyab1);




                                    }
                                    else
                                    {
                                        return "کد پرسنلی نامعتبر است";
                                    }

                                }
                                db.tbMoadelPadashJarimeAyab.AddRange(tbMoadelPadashJarimeAyab11);
                                await db.SaveChangesAsync();

                                break;
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
        public async Task<ActionResult> ImportExcel_Gharardady(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {
                string Message = await GetDataFromExcel_Gharardady(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);
            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }
        }





        public async Task<string> GetDataFromExcel_Gharardady(HttpPostedFileBase MyExcelStream)
        {

            int idkarbar2 = 0;
            int value2 = 0;
            var harchy2 = 0;

            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    List<tbUserContractsAndMoalefeGhararDadi> tbUserContractsAndMoalefeGhararDadi2 = new List<tbUserContractsAndMoalefeGhararDadi>();
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }

                                var count = workbook.Sheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }


                                var tbUsers = await db.tbUsers.ToListAsync();
                                var tbUserContracts = await db.tbUserContracts.ToListAsync();

                                for (int i = 1; i < count; i++)
                                {

                                    var row = workbook.Sheets[0].Rows[i];
                                    var idkarbar = row.Cells[0];
                                    var value = row.Cells[1];
                                    var harchy = row.Cells[2];

                                    idkarbar2 = int.Parse(idkarbar.Value.ToString());
                                    value2 = int.Parse(value.Value.ToString());
                                    harchy2 = int.Parse(harchy.Value.ToString());


                                    var findkarbar = tbUsers.Where(p => p.usr_Personal_ID == idkarbar2).SingleOrDefault();
                                    var findkarbar2 = tbUserContracts.Where(n => n.FK_UserID == findkarbar.usr_ID).SingleOrDefault();
                                    tbUserContractsAndMoalefeGhararDadi tbUserContractsAndMoalefeGhararDadi1 = new tbUserContractsAndMoalefeGhararDadi();


                                    if (findkarbar2 != null)
                                    {
                                        tbUserContractsAndMoalefeGhararDadi1.FKContractID = findkarbar2.usc_ID;
                                        tbUserContractsAndMoalefeGhararDadi1.Value = value2;
                                        tbUserContractsAndMoalefeGhararDadi1.FKMoalefeGhararDadi = harchy2;
                                        tbUserContractsAndMoalefeGhararDadi2.Add(tbUserContractsAndMoalefeGhararDadi1);
                                    }
                                    else
                                    {
                                        return "کد پرسنلی نامعتبر است";
                                    }

                                }

                                db.tbUserContractsAndMoalefeGhararDadi.AddRange(tbUserContractsAndMoalefeGhararDadi2);
                                await db.SaveChangesAsync();

                                break;
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










































        public class CheckboxData
        {
            public List<string> Checkboxes { get; set; }
            public string Message { get; set; }
        }

        [HttpPost]
        public ActionResult creatnew1(Dictionary<string, object> data)
        {
            // استخراج داده‌ها
            bool monthChecked = data.ContainsKey("monthChecked") ? System.Convert.ToBoolean(data["monthChecked"]) : false;
            bool yearChecked = data.ContainsKey("yearChecked") ? System.Convert.ToBoolean(data["yearChecked"]) : false;
            bool legalChecked = data.ContainsKey("legalChecked") ? System.Convert.ToBoolean(data["legalChecked"]) : false;
            bool realChecked = data.ContainsKey("realChecked") ? System.Convert.ToBoolean(data["realChecked"]) : false;
            bool descriptionChecked = data.ContainsKey("descriptionChecked") ? System.Convert.ToBoolean(data["descriptionChecked"]) : false;
            bool unitChecked = data.ContainsKey("unitChecked") ? System.Convert.ToBoolean(data["unitChecked"]) : false;
            bool projectChecked = data.ContainsKey("projectChecked") ? System.Convert.ToBoolean(data["projectChecked"]) : false;
            string alertMessage = data.ContainsKey("alertMessage") ? data["alertMessage"].ToString() : "";

            // انجام پردازش‌های لازم
            return Json("true"); // یا هر چیزی که بخواهید به فرانت ارسال کنید
        }

        [HttpPost]
        public ActionResult creatnew(bool monthChecked, bool yearChecked, bool legalChecked, bool realChecked, bool DAyChecked, bool shomarehnameh, bool timestartChecked, bool timeendChecked, bool timerepetChecked, bool timetamdidChecked, bool mablaghChecked, bool mozohChecked,
                              bool descriptionChecked, bool unitChecked, bool projectChecked, string alertMessage)
        {
            dbAsnaduplaodfil2 dbAsnaduplaodfil2 = new dbAsnaduplaodfil2();

            // استفاده از داده‌های دریافت شده از فرانت
            // به طور مثال:
            bool monthStatus = monthChecked;
            bool yearStatus = yearChecked;
            bool legalStatus = legalChecked;
            bool realStatus = realChecked;
            bool descriptionStatus = descriptionChecked;
            bool unitStatus = unitChecked;
            bool projectStatus = projectChecked;

            // مقدار alertMessage
            string alertMessageReceived = alertMessage;
            
            var findd = db.dbAsnaduplaodfil1.Where(s => s.codmahsol == alertMessageReceived).FirstOrDefault();
            if (findd != null)
            {
                var find2 = db.dbAsnaduplaodfil2.Where(s => s.FK_NAME == findd.ID).FirstOrDefault();
                if (find2 != null)
                {
                    find2.defaultRadio1 = monthStatus;
                    find2.defaultRadio2 = yearChecked;
                    find2.defaultRadio3 = legalChecked;
                    find2.defaultRadio4 = realChecked;
                    find2.defaultRadio5 = descriptionChecked;
                    find2.defaultRadio6 = unitChecked;
                    find2.defaultRadio7 = projectChecked;
                    find2.DAyChecked = DAyChecked;
                    find2.shomarehnameh = shomarehnameh;
                    find2.timestartChecked = timestartChecked;
                    find2.timeendChecked = timeendChecked;
                    find2.mozohChecked = mozohChecked;
                    find2.mablaghChecked = mablaghChecked;
                    find2.timerepetChecked = timerepetChecked;
                    find2.timetamdidChecked = timetamdidChecked;

                    db.SaveChanges();
                }
                else
                {
                    dbAsnaduplaodfil2.FK_NAME = findd.ID;

                    dbAsnaduplaodfil2.defaultRadio1 = monthStatus;
                    dbAsnaduplaodfil2.defaultRadio2 = yearChecked;
                    dbAsnaduplaodfil2.defaultRadio3 = legalChecked;
                    dbAsnaduplaodfil2.defaultRadio4 = realChecked;
                    dbAsnaduplaodfil2.defaultRadio5 = descriptionChecked;
                    dbAsnaduplaodfil2.defaultRadio6 = unitChecked;
                    dbAsnaduplaodfil2.defaultRadio7 = projectChecked;
                    dbAsnaduplaodfil2.DAyChecked = DAyChecked;
                    dbAsnaduplaodfil2.shomarehnameh = shomarehnameh;
                    dbAsnaduplaodfil2.timestartChecked = timestartChecked;
                    dbAsnaduplaodfil2.timeendChecked = timeendChecked;
                    dbAsnaduplaodfil2.mozohChecked = mozohChecked;
                    dbAsnaduplaodfil2.mablaghChecked = mablaghChecked;
                    dbAsnaduplaodfil2.timerepetChecked = timerepetChecked;
                    dbAsnaduplaodfil2.timetamdidChecked = timetamdidChecked;
                    db.dbAsnaduplaodfil2.Add(dbAsnaduplaodfil2);
                    db.SaveChanges();
                }


            }
      
            // پردازش‌های مورد نظر را انجام دهید

            // در اینجا می‌توانید پاسخ مناسب را ارسال کنید
            return Json("true");
        }









        public async Task<string>  GetDataFromExcel_AdamCode(HttpPostedFileBase MyExcelStream)
        {
            int karbar = 0;
            List<dbtarifmahdodayt2> mahdodayt = new List<dbtarifmahdodayt2>();


            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            var tbUsers = await db.tbUsers.ToListAsync();
                            var dbtarifmahdodayt1 = await db.dbtarifmahdodayt1.ToListAsync();
                            var dbtarifmahdodayt2 = await db.dbtarifmahdodayt2.ToListAsync();

                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title2 = workbook.Sheets[0].Rows[0].Cells;

                                //var title = workbook.Sheets[0].Rows[0].Cells;
                                //var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                List<int> ID = new List<int>();
                                for (int i = 0; i < title2.Count; i++)
                                {
                                    if (title2[i]?.Value == null)
                                        continue;

                                    ID.Add(int.Parse(title2[i].Value.ToString()));
                                }

                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                for (int i = 2; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var cells = row.Cells.ToList();
                                    for (int j = 0; j < sheet.Rows[0].Cells.Count; j++)
                                    {
                                        if (j >= cells.Count || cells[j] == null)
                                        {
                                            cells.Add(new Telerik.Web.Spreadsheet.Cell { Value = null });
                                        }
                                    }

                                    var Name = row.Cells[1];
                                    if (Name.Value != null)
                                    {
                                        karbar = int.Parse(Name.Value.ToString());
                                        //name = Name.Value; // Convert Name.Value to string
                                        //if (!int.TryParse(Name.Value?.ToString(), out name))
                                        //{
                                        //    return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                        //}

                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    int shomarande = 2;

                                    //if (shomarande < row.Cells.Count)
                                    //{

                                    var find = tbUsers.Where(s => s.usr_Personal_ID == karbar).FirstOrDefault();

                                    foreach (var it in ID)

                                    {
                                        if (shomarande < cells.Count)
                                        {
                                            

                                        
                                        if (it != 0 && find!=null)
                                        {
                                            var find3 = dbtarifmahdodayt1.Where(r => r.FK_moalfe == it).FirstOrDefault();
                                            var finddbtarifmahdodayt2 = dbtarifmahdodayt2.Where(p => p.FK_moalfe == find3.ID && p.FK_usr == find.usr_ID).FirstOrDefault();
                                            if (finddbtarifmahdodayt2 == null)
                                            {
                                                var Value = row.Cells[shomarande];
                                                if (Value.Value != null)
                                                {
                                                    dbtarifmahdodayt2 dbtarifmahdodayt = new dbtarifmahdodayt2();
                                                    dbtarifmahdodayt.FK_moalfe = find3.ID;
                                                    dbtarifmahdodayt.FK_usr = find.usr_ID;
                                                    dbtarifmahdodayt.value = int.Parse(Value.Value.ToString());
                                                    mahdodayt.Add(dbtarifmahdodayt);
                                                }
                                                else
                                                {
                                                    dbtarifmahdodayt2 dbtarifmahdodayt = new dbtarifmahdodayt2();
                                                    dbtarifmahdodayt.FK_moalfe = find3.ID;
                                                    dbtarifmahdodayt.FK_usr = find.usr_ID;
                                                    dbtarifmahdodayt.value = 0;
                                                    mahdodayt.Add(dbtarifmahdodayt);
                                                }
                                            }
                                            else
                                            {
                                                var Value = row.Cells[shomarande];
                                                if (Value.Value != null)
                                                {

                                                    finddbtarifmahdodayt2.value = int.Parse(Value.Value.ToString());
                                                    //await db.SaveChangesAsync();

                                                }
                                                else
                                                {

                                                    finddbtarifmahdodayt2.value = 0;

                                                    //await db.SaveChangesAsync();
                                                }
                                            }
                                            shomarande++;

                                        }
                                        }
                                    }
                                    //}
                                    //else
                                    //{
                                    //    return "ایندکس خارج از محدوده در سطر " + (i + 1);
                                    //}


                                }
                                db.dbtarifmahdodayt2.AddRange(mahdodayt);
                                await db.SaveChangesAsync();
                                break;
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


        public ActionResult ExportMoalefeExcel888()
        {
            var OutPutFile = SetDataExcel_Moalefe9();
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "EquipmentGroup" + extension);
        }

        public ActionResult ImportExcel_AddSabtmarahel(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_FishtAEAD(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }




        public ActionResult ImportExcel_Addsabtdasteh(HttpPostedFileBase MyExcelStream)
        {
            if (MyExcelStream != null)
            {

                string Message = GetDataFromExcel_FishtAEAD2(MyExcelStream);
                //TempData["Message"] = Message;
                //return RedirectToAction("ManageMoalefeDastMozdi", "MoalefeDastMozdi");
                return Content(Message);

            }
            else
            {
                return Content("فایل بطور صحیح بارگذاری نشده است");

            }

        }

        public string GetDataFromExcel_FishtAEAD2(HttpPostedFileBase MyExcelStream)
        {
            int mlfvlfsh_Year, mlfvlfsh_Month, malefe;
            float mlfvlfsh_Value;
            string name = "";
            string name2 = "";
            List<tbMoalefeValueFish> IsFish = new List<tbMoalefeValueFish>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title = workbook.Sheets[0].Rows[0].Cells;
                                var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null)
                                    {
                                        name = (string)Name.Value; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }
                                    var Name2 = row.Cells[1];
                                    if (Name.Value != null)
                                    {
                                        name2 = (string)Name2.Value; // Assuming Name.Value is convertible to an integer.
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var c2 = db.tbEquipmentGroup.Where(p => p.Eqpgrp_Name == name).Select(s => s.Eqpgrp_ID).FirstOrDefault();


                                    tbEquipmentBunch tbm = new tbEquipmentBunch();
                                    tbm.Eqpbnch_Name = name2;
                                    tbm.FK_EqpgrpID = c2;
                                    db.tbEquipmentBunch.Add(tbm);
                                    db.SaveChanges();


                                }

                                break;
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


        public string GetDataFromExcel_FishtAEAD(HttpPostedFileBase MyExcelStream)
        {
            int mlfvlfsh_Year, mlfvlfsh_Month, malefe;
            float mlfvlfsh_Value;
            string name = "";
            List<tbMoalefeValueFish> IsFish = new List<tbMoalefeValueFish>();
            var workbook = Telerik.Web.Spreadsheet.Workbook.Load(MyExcelStream.InputStream, Path.GetExtension(MyExcelStream.FileName));
            using (DbContextTransaction transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var Count_Sheets = workbook.Sheets.Count;
                    if (Count_Sheets >= 1)
                    {
                        var sheet = workbook.Sheets[0];
                        var Rows = sheet.Rows;
                        if (Rows.Count >= 1)
                        {
                            foreach (var item in Rows)
                            {
                                if (item.Cells.Count != sheet.Rows[0].Cells.Count)
                                {
                                    return " خطای ارزیابی در سطر " + (item.Index + 1).ToString() + " رخ داده است. یکی از سلول های این ستون فاقد اطلاعات می باشد ";
                                    ;
                                }
                                var title = workbook.Sheets[0].Rows[0].Cells;
                                var List = workbook.Sheets[0].Rows;
                                var count = workbook.Sheets[0].Rows.Count();
                                if (count == 1)
                                {
                                    return "فایل اکسل فاقد اطلاعات می باشد";
                                }
                                for (int i = 1; i < count; i++)
                                {
                                    var row = workbook.Sheets[0].Rows[i];
                                    var Name = row.Cells[0];
                                    if (Name.Value != null)
                                    {
                                        name = Name.Value.ToString(); // Convert Name.Value to string
                                        name = name.TrimEnd('\n');
                                    }
                                    else
                                    {
                                        return " لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ";
                                    }

                                    var c2 = db.tbContractMoalefeDastmozdi.Where(p => p.md_Title.Contains(name)).Select(s => s.md_ID).FirstOrDefault();


                                    tbEquipmentGroup tbm = new tbEquipmentGroup();
                                    tbm.Eqpgrp_Name = name;
                                    db.tbEquipmentGroup.Add(tbm);
                                    db.SaveChanges();


                                }

                                break;
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






        public Workbook SetDataExcel_Moalefe6()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EquipmentGroup.xlsx"));
            Row Row;
            var x = db.tbEquipmentGroup.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.Eqpgrp_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },


                    });

                    counter++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }
        public Workbook SetDataExcel_Moalefe9()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EquipmentBunch.xlsx"));
            Row Row;
            var x = db.tbEquipmentGroup.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.Eqpgrp_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },


                    });

                    counter++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }




        public Workbook esd()
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/EquipmentBunch.xlsx"));
            Row Row;
            var x = db.tbEquipmentGroup.ToList();

            int counter = 1;
            foreach (var item in x)
            {
                Row = new Row() { Height = 20, Index = counter };
                {
                    Row.AddCells(new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = item.Eqpgrp_Name,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },


                    });

                    counter++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }




            }
            return Moalefeexcelfile;
        }



        [AuthorizeAAA]
        public ActionResult CreateGroup(string GroupTitle)
        {
            if (GroupTitle.Trim() == "")
            {
                TempData["response"] = "NULL";
                return RedirectToAction("Index", "DefFacilitiesEquipment");

            }
            else
            {
                tbEquipmentGroup entity = new tbEquipmentGroup
                {
                    Eqpgrp_Name = GroupTitle
                };
                if (EquiprpRepo.Create(entity) == "false")
                {
                    TempData["response"] = "False";
                    return RedirectToAction("Index", "DefFacilitiesEquipment");

                }
                else
                {
                    TempData["response"] = "True";
                    return RedirectToAction("Index", "DefFacilitiesEquipment");

                }
            }


        }
        [AuthorizeAAA]
        public ActionResult CreateBunch(int GrpId, string BunchTitle)
        {
            if (BunchTitle.Trim() == "")
            {
                TempData["response"] = "NULL";
                return RedirectToAction("Index", "DefFacilitiesEquipment");

            }
            else
            {
                tbEquipmentBunch entity = new tbEquipmentBunch
                {
                    Eqpbnch_Name = BunchTitle,
                    FK_EqpgrpID = GrpId,

                };
                if (EquipbunchRepo.Create(entity) == "false")
                {
                    TempData["response"] = "False";
                    return RedirectToAction("Index", "DefFacilitiesEquipment");

                }
                else
                {
                    TempData["response"] = "True";
                    return RedirectToAction("Index", "DefFacilitiesEquipment");

                }
            }


        }
        [AuthorizeAAA]
        public ActionResult _ListGroup()
        {
            var f = db.tbmarahelsabt.Where(p => p.FK_Group != null).ToList();
            List<tbEquipmentGroup> Model = new List<tbEquipmentGroup>();
            foreach (var it in f)
            {
                var t = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == it.FK_Group).FirstOrDefault();
                if (t != null)
                {
                    Model.Add(t);
                }
            }
            if (Model != null)
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListGroup.cshtml", Model);

            }
            else
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListGroup.cshtml", new tbEquipmentGroup());

            }

        }
        [AuthorizeAAA]
        public ActionResult _ListGroupForBunch()
        {
            var Model = EquiprpRepo.Listt();
            if (Model != null)
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListGroupForBunch.cshtml", Model);

            }
            else
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListGroupForBunch.cshtml", new tbEquipmentGroup());

            }

        }
        [AuthorizeAAA]
        public ActionResult _ListBunch(int id)
        {
            var Model = EquipbunchRepo.FindByGRP(id);
            if (Model != null)
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListBunch.cshtml", Model);

            }
            else
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListBunch.cshtml", new tbEquipmentBunch());

            }

        }

        public ActionResult FilterSabt_Usr_Tools(int fk_basteh = 0, int Month = 0, int Year = 0)
        {
            List<tbEquipmentMoalefeValueReffrenceSave> id = new List<tbEquipmentMoalefeValueReffrenceSave>();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            int idddd = 0;
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        idddd = User.usr_ID;
                    }

                }
            }

            var x = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == fk_basteh && p.FK_User == idddd && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).ToList();
            if (x.Count != 0)
            {

                foreach (var item in x)
                {

                    id.Add(item);


                }
                return View(id);


            }

            return View(new List<tbEquipmentMoalefeValue>());

        }



        public ActionResult viewsoartsabt()
        {
            return PartialView();
        }

        public ActionResult Fishsabtview()
        {
            return View();
        }
        public ActionResult Acceptviewsorat()
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository rep_accept = new SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository();
                        var result = rep_accept.ListtForproje(User.usr_ID);
                        List<int?> baste_id = new List<int?>();
                        foreach (var item in result)
                        {
                            baste_id.Add(item.FK_PeymanID);
                        }
                        return View(NotSubmitedMoalefe(baste_id));
                    }

                }
            }
            return View(new List<tbSavedFunctions>());

        }
        public List<tbCommodityProject> NotSubmitedMoalefe(List<int?> Bastehlst)
        {

            return db.tbCommodityProject.Where(p => Bastehlst.Contains((int)p.FK_PeymanContracts) && p.tbPeymanContracts.Inactive != true).ToList();

        }


        [AuthorizeAAA]
        public ActionResult UpdateBunch(int cnt, int GrpIdBunch, int bunch)
        {
            var entity = EquipbunchRepo.Find(bunch);
            entity.FK_EqpgrpID = GrpIdBunch;
            entity.Eqpbnch_CountFeatures = cnt;

            if (System.Convert.ToBoolean(db.SaveChanges()))
            {
                TempData["response"] = "False";
                return RedirectToAction("Index", "DefFacilitiesEquipment");

            }
            else
            {
                TempData["response"] = "True";
                return RedirectToAction("Index", "DefFacilitiesEquipment");

            }
        }
        [AuthorizeAAA]
        public ActionResult CreateFeatures(FeatureDetail res)
        {
            return RedirectToAction("Index", "DefFacilitiesEquipment");
        }


        /// <summary>
        /// تابع اضافه کردن فیلد مشخصات  برای گروه و دسته بندی مشخص شده
        /// </summary>
        /// <param name="ListSpec">لیست عناوین فیلد ها</param>
        /// <param name="FK_EqpGroup_ID">آیدی گروه فیلد</param>
        /// <param name="FK_EqpBunch_ID">آیدی دسته بندی فیلد</param>
        /// <returns></returns>
        [AuthorizeAAA]
        [HttpPost]
        public ActionResult AddSpecification(List<string> ListSpec, int FK_EqpGroup_ID, int FK_EqpBunch_ID)
        {
            foreach (var item in ListSpec)
            {
                tbEquipmentSpecifications tbEquipmentSpecifications = new tbEquipmentSpecifications()
                {
                    sp_SpecTitle = item,
                    FK_EqpBunch_ID = FK_EqpBunch_ID,

                };
                var State = SpecificationsRepository.Create(tbEquipmentSpecifications);
                if (State == "Error")
                {
                    return Content(State);
                }
            }

            return Content(SpecificationsRepository.SaveChanges().ToString());

        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult GetSpecifications(int FK_EqpBunch_ID, int counter = 1)
        {
            var Spes = SpecificationsRepository.GetSpecGroupAndBunch(FK_EqpBunch_ID);
            EquipmentVM Model = new EquipmentVM
            {
                Counter = counter,
                listEquip = Spes,
            };

            return PartialView(Model);
        }

        [AuthorizeAAA]
        public ActionResult _GetInformationEquipmentGroup()
        {
            var Model = EquiprpRepo.Listt();
            return PartialView(Model);
        }

        [AuthorizeAAA]
        [HttpPost]
        public ActionResult SaveSpecData(List<tbEquipments> equipments)
        {


            if (tbequipRepo.AddRange(equipments))
            {
                return Content("True");
            }
            else
            {
                return Content("False");

            }

        }
        [AuthorizeAAA]
        public ActionResult _ShowSpecificationTable(int Basteh, int peyman)
        {

            var Model = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == peyman && p.tbPeymanContracts.Inactive != true).ToList();
            List<tbEquipments> db3 = new ArrayList<tbEquipments>();
            foreach (var item2 in Model)
            {
                tbEquipments db2 = new tbEquipments();

                db2.TotalPrice = item2.TotalPrice;
                db2.TypeTazmin = item2.TypeTazmin;
                db2.Unit = item2.Unit;
                db2.UnitCount = item2.UnitCount;
                db2.Count_InQardad = item2.Count_InQardad;
                db2.EachValue = item2.EachValue;
                db2.ExpiredDate = item2.ExpiredDate;
                db2.FK_Bunch = item2.FK_Bunch;
                db2.FK_Peyman = item2.FK_Peyman;
                db2.Name = item2.Name;
                db2.tbEquipmentBunch = item2.tbEquipmentBunch;
                db2.tbEquipmentSpecificationData = item2.tbEquipmentSpecificationData;
                db2.tbPeymanContracts = item2.tbPeymanContracts;
                db2.tbEquipmentMoalefeValue = item2.tbEquipmentMoalefeValue;
                db2.personal = item2.personal;
                db2.ID = item2.ID;
                db2.Fk_usr = item2.Fk_usr;
                db2.Soght = item2.Soght;
                db2.egareh = item2.egareh;
                db2.type = item2.type;
                db2.uniform = item2.uniform;
                db2.model = item2.model;
                var c = "";
                if (item2.Users != null)
                {
                    var model2 = item2.Users.Split(',');
                    foreach (var item3 in model2)
                    {
                        var id = System.Convert.ToInt32(item3);
                        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                        c += temp.FullName + " , ";

                    }
                }
                db2.Users = c;


                db3.Add(db2);


            }
            //foreach (var item in finlist)
            //{
            //    string typetazmin = "";
            //    switch (item.TypeTazmin)
            //    {
            //        case 1:
            //            typetazmin = "چک";
            //            break;
            //        case 2:
            //            typetazmin = "سفته";
            //            break;
            //        case 3:
            //            typetazmin = "ضمانت نامه بانکی";
            //            break;
            //        case 4:
            //            typetazmin = "وجه نقد";
            //            break;
            //        case 5:
            //            typetazmin = "نامه ضمانت";
            //            break;
            //        default:
            //            break;
            //    }
            //    string timeunit = "";

            //    switch (item.Unit)
            //    {
            //        case 1:
            //            timeunit = "ساعتی";
            //            break;
            //        case 2:
            //            timeunit = "روزانه";
            //            break;
            //        case 3:
            //            timeunit = "ماهیانه";
            //            break;
            //        case 4:
            //            timeunit = "سالانه";
            //            break;
            //        case 5:
            //            timeunit = "نامه ضمانت";
            //            break;
            //        default:
            //            break;
            //    }

            //    var usr = item.Users.Split(',');
            //    StringBuilder usrnamelist = new StringBuilder();
            //    foreach (var item2 in usr)
            //    {
            //        if (item2 != "")
            //        {
            //            var y = System.Convert.ToInt32(item2);
            //            var x = db.tbUsers.FirstOrDefault(p => p.usr_ID == y).FullName;
            //            usrnamelist.Append(x + " , ");
            //        }

            //    }
            //    List<EquipmentInfo> Equipment = new List<EquipmentInfo>();
            //    var specificdata = db.tbEquipmentSpecifications.Where(p => p.FK_EqpBunch_ID == Basteh).ToList();

            //    foreach (var item3 in specificdata)
            //    {
            //        EquipmentInfo quipinfo = new EquipmentInfo
            //        {
            //            Name = item3.sp_SpecTitle,
            //            Value = db.tbEquipmentSpecificationData.FirstOrDefault(p => p.FK_Spec_ID == item3.sp_ID && p.FK_Equipment == Basteh).SpcData
            //        };

            //        Equipment.Add(quipinfo);
            //    }

            //    FinantialEquipmentVM finequip = new FinantialEquipmentVM
            //    {
            //        EachValue = (long)item.EachValue,
            //        ExpireDate = item.ExpiredDate,
            //        TazminPrice = (long)item.TazminPrice,
            //        TazminType = typetazmin,
            //        Unit = timeunit,
            //        TotalPrice = (long)item.TotalPrice,
            //        UnitCount = (int)item.UnitCount,
            //        UsersName = usrnamelist.ToString(),
            //        Tajhizat = Equipment

            //    };

            //    Model.Add(finequip);


            //}

            return PartialView("_ShowSpecificationTable", db3);
        }
        public ActionResult editeqview(int id)
        {
            var y = db.tbEquipments.Where(p => p.ID == id).FirstOrDefault();
            return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/editeqview.cshtml", y);
        }

        public ActionResult editeq(long EachValue, long TotalPrice, int UnitCount, int id, float Soght, float egareh, int type, float uniform)
        {
            var y = db.tbEquipments.Where(p => p.ID == id).FirstOrDefault();
            y.EachValue = EachValue;
            y.TotalPrice = TotalPrice;
            y.UnitCount = UnitCount;
            y.uniform = uniform;
            y.type = type;
            y.Soght = Soght;
            y.egareh = egareh;
            db.SaveChanges();
            return Content("True");

        }
        [AuthorizeAAA]
        public ActionResult DeleteEquipment(int ID)
        {
            try
            {
                if (tbequipRepo.Delete(ID))
                {
                    return Content("True");
                }
                else
                {
                    return Content("False");
                }
            }
            catch (Exception)
            {

                return Content("False");
            }

        }

        //public ActionResult FinantialEquipment(string TotalPrice, int Unit, int Count, string single, int pymnID, int BunchID, int groupID)
        //{
        //    try
        //    {

        //        tbEquipmentFinantial equfin = new tbEquipmentFinantial
        //        {
        //            equpfin_Single = System.Convert.ToInt32(single),
        //            equpfin_TotalPrice = System.Convert.ToInt32(TotalPrice),
        //            equpfin_Count = Count,
        //            equpfin_Unit = Unit,
        //            FK_BunchID = BunchID,
        //            FK_GroupID = groupID,
        //            FK_PeymanID = pymnID,

        //        };
        //        if (finantialRepo.Create(equfin) == "True")
        //        {
        //            return Content("1");

        //        }
        //        else
        //        {
        //            return Content("0");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        return Content("0");

        //    }
        //}


        #region ثبت کارکرد تجهیزات tbequipmentMoalefeValue
        [AuthorizeAAA]
        public ActionResult _CreateEquipmentMoalefeValue()
        {
            return PartialView();
        }

        public ActionResult ExportExcel_MachinsOrToolsForcheack(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0, int Basteh = 0)


        {
            //var c = db.tbReffrenceSaveLevel.Where(p => p.ID==Basteh).Select(p=>p.FK_RRSave).FirstOrDefault();
            //var exi=db.tbReffrenceSave.Where(p=>p.ID==c).Select(p=>p.FK_PeymanID).FirstOrDefault();
            //if (exi== Peyman_ID) {
            var find = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == Basteh && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefault();

            if (find != null)
            {
                return Content("False");
            }
            else
            {
                return Content("true");
            }
        }

        //else
        //{
        //    return Content("Falses");
        //}
        //}





        [AuthorizeAAA]
        public ActionResult ExportExcel_MachinsOrTools(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)
        {


            // جداکننده مورد نظر








            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel4(Peyman_ID, type, Month, Year);

            string extension = ".xlsx";
            var py = db.tbPeymanContracts.Where(p => p.pec_ID == Peyman_ID).FirstOrDefault();
            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            string peymanName = py.pec_Title;

            char separator = '-';
            string ty = "";
            // یافتن موقعیت جداکننده در رشته
            int startIndex = peymanName.IndexOf(separator);
            string extractedFileName = peymanName.Substring(startIndex + 1).Trim();
            if (type == 3)
            {
                ty = "خودرو";
            }
            else
            {
                ty = "ابزار";

            }
            return File(stream.ToArray(), mimeType, ty + '_' + extractedFileName + extension);

        }
        public ActionResult ExportExcel_MachinsOrTools222tst(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)
        {


            // جداکننده مورد نظر








            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel4080323(Peyman_ID, type, Month, Year);

            string extension = ".xlsx";
            var py = db.tbPeymanContracts.Where(p => p.pec_ID == Peyman_ID).FirstOrDefault();
            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            string peymanName = py.pec_Title;

            char separator = '-';
            string ty = "";
            // یافتن موقعیت جداکننده در رشته
            int startIndex = peymanName.IndexOf(separator);
            string extractedFileName = peymanName.Substring(startIndex + 1).Trim();
            if (type == 3)
            {
                ty = "خودرو";
            }
            else
            {
                ty = "ابزار";

            }
            return File(stream.ToArray(), mimeType, ty + '_' + extractedFileName + extension);

        }

        //http://localhost:6061/contracts/defFacilitiesEquipment/ExportExcel_MachinsOrTools2224tst?Peyman_ID=49&type=4&Month=6&Year=1403

        public ActionResult ExportExcel_MachinsOrTools2224tst(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)
        {


            // جداکننده مورد نظر








            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel408032trst3(Peyman_ID, type, Month, Year);

            string extension = ".xlsx";
            var py = db.tbPeymanContracts.Where(p => p.pec_ID == Peyman_ID).FirstOrDefault();
            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            string peymanName = py.pec_Title;

            char separator = '-';
            string ty = "";
            // یافتن موقعیت جداکننده در رشته
            int startIndex = peymanName.IndexOf(separator);
            string extractedFileName = peymanName.Substring(startIndex + 1).Trim();
            var fib = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == type).Select(s => s.Eqpgrp_Name).FirstOrDefault();
            if (type == 3)
            {
                ty = "خودرو";
            }
            else if (type == 4)
            {
                ty = "ابزار";

            }
            else
            {
                ty = fib;
            }
            return File(stream.ToArray(), mimeType, ty + '_' + extractedFileName + extension);

        }
        //=================================================================================================================

        private DataTable GetDataForMachinsOrTools(int peyman_ID, int type, int Month, int Year = 0)
        {
            // ایجاد یک DataTable برای ذخیره داده‌ها
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("نام وسیله", typeof(string));
            dataTable.Columns.Add("نوع ماشین", typeof(string));
            dataTable.Columns.Add("مشخصات", typeof(string));
            dataTable.Columns.Add("تعداد روزها", typeof(double));
            dataTable.Columns.Add("مبلغ کل", typeof(double));

            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();

            foreach (var item in equipments)
            {
                // پردازش اطلاعات
                string allinfo = "";
                foreach (var item2 in item.tbEquipmentSpecificationData)
                {
                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + " / ";
                }

                string machineType = "";
                if (item.type != null)
                {
                    if (item.type == 1) machineType = "تیرسان";
                    else if (item.type == 2) machineType = "یونیفرم ساده";
                    else if (item.type == 3) machineType = "بدون یونیفرم";
                }

                // مقداردهی و اضافه کردن ردیف به DataTable
                double totalValue = 0;  // مثال برای محاسبه، بر اساس داده‌های مورد نیاز شما
                DataRow row = dataTable.NewRow();
                row["ID"] = item.ID;
                row["نام وسیله"] = item.tbEquipmentBunch?.Eqpbnch_Name;
                row["نوع ماشین"] = machineType;
                row["مشخصات"] = allinfo;
                row["تعداد روزها"] = 0; // این مقدار را باید محاسبه کنید
                row["مبلغ کل"] = totalValue;

                dataTable.Rows.Add(row);
            }

            return dataTable; // برگرداندن DataTable به جای فایل اکسل
        }
        public ActionResult MachinsOrTools_Peyvast(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)

        {
            var OutPutFile = GetDataForMachinsOrTools(Peyman_ID, type, Month, Year);
            ViewBag.OutPutFile = OutPutFile;
            return View("~/Areas/Salaries/Views/Formula/ViewPeyvast_Abzaar.cshtml");
        }
        
            
            //=================================================================================================================

            public ActionResult ExportExcel_MachinsOrTools222(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)
        {


            // جداکننده مورد نظر








            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel4080(Peyman_ID, type, Month, Year);

            string extension = ".xlsx";
            var py = db.tbPeymanContracts.Where(p => p.pec_ID == Peyman_ID).FirstOrDefault();
            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];
            string peymanName = py.pec_Title;

            char separator = '-';
            string ty = "";
            // یافتن موقعیت جداکننده در رشته
            int startIndex = peymanName.IndexOf(separator);
            string extractedFileName = peymanName.Substring(startIndex + 1).Trim();
            if (type == 3)
            {
                ty = "خودرو";
            }
            else
            {
                ty = "ابزار";

            }
            return File(stream.ToArray(), mimeType, ty + '_' + extractedFileName + extension);

        }
        public ActionResult ExportExcel_MachinsOrToolsforadmin(int Peyman_ID = 0, int type = 0, int Month = 0, int Year = 0)
        {





            //var OutPutFile = SetDataToMachinsOrToolsExcel(Peyman_ID, type, Month, Year);
            var OutPutFile = SetDataToMachinsOrToolsExcel40(Peyman_ID, type, Month, Year);

            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Excel_MachinsOrTools" + extension);

        }
        public ActionResult ExportExcel_MachinsOrTools2(string Peyman_ID2 = "", int type2 = 0, DateTime Fromdata = default, DateTime Todate = default, int id = 0)
        {
            var Peyman_ID = db.tbPeymanContracts.Where(p => p.pec_Title == Peyman_ID2 && p.Inactive != true).FirstOrDefault();

            var OutPutFile = SetDataToMachinsOrToolsExcel2(Peyman_ID.pec_ID, type2, Fromdata, Todate, id);
            string extension = ".xlsx";

            var stream = new MemoryStream();
            OutPutFile.Save(stream, extension);
            var mimeType = MimeTypes.ByExtension[extension];

            return File(stream.ToArray(), mimeType, "Excel_MachinsOrTools" + extension);
        }


        private Workbook SetDataToMachinsOrToolsExcel4(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
            Row Row;
            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
            var pynm22 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
            if (equipments.Count() > 0)
            {
                int counter = 1;
                var count = 2;
                var counttt = 2;

                var count1 = 4;
                var count2 = 1;

                var count3 = 2;


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }




                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
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


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }

                    if (type == 3)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----"+"  نوع ماشین:  "+  t1,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }
                    else
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----",
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }



                    count3++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }
                var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true).ToList();
                var pymn2 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
                var company = db.tbCompanies.Where(p => p.ID == pymn2.FK_KarfarmaID).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = company.CompanyName,
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
                        Value = company.Company_National_ID,
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

                Row = new Row() { Height = 20, Index = 3 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = pymn2.pec_Title,
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
                        Value = pymn2.pec_ProjectCode,
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

                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = item.tbUsers.FullName,
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
            Value = item.tbUsers.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
    });



                    count1++; // افزایش ایندکس ردیف‌ها

                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }



                var user = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
                var pymn = db.tbPeymanContracts.Where(p => p.Inactive == null).ToList();
                foreach (var item in user)
                {
                    Row = new Row() { Height = 20, Index = count1 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.FullName,
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
                        Value = item.usr_Personal_ID,
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
                    count1++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }

                foreach (var item in pymn)
                {
                    Row = new Row() { Height = 20, Index = count2 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 2
                    },
                         new Cell()
                    {
                        Value = item.pec_ProjectCode,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 3
                    },
                });
                    count2++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }


            }
            return Moalefeexcelfile;

        }



        private Workbook SetDataToMachinsOrToolsExcel40(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
            Row Row;
            var pymn = db.tbPeymanContracts.Where(p => p.Inactive == null).ToList();
            int counter = 1;
            var count = 2;
            var count1 = 1;
            var count2 = 1;
            var count3 = 1;

            foreach (var itemm in pymn)
            {
                var equipments = db.tbEquipments.Where(p => p.FK_Peyman == itemm.pec_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
                var pynm22 = db.tbPeymanContracts.Where(p => p.pec_ID == itemm.pec_ID && p.Inactive != true).FirstOrDefault();
                if (equipments.Count() > 0)
                {




                    foreach (var item in equipments)
                    {
                        string allinfo = "";
                        foreach (var item2 in item.tbEquipmentSpecificationData)
                        {
                            allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                        }
                        var c = "";

                        if (item.Users != null)
                        {
                            var model2 = item.Users.Split(',');
                            foreach (var item3 in model2)
                            {
                                var id = System.Convert.ToInt32(item3);
                                var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                                c += temp.FullName + " , ";

                            }
                        }
                        var t = "";
                        if (item.personal != null)
                        {
                            if (item.personal == true)
                            {
                                t = "پرسنلی";
                            }
                            else
                            {
                                t = "کارگزاری";

                            }
                        }
                        var t1 = "";
                        if (item.type != null)
                        {
                            if (item.type == 1)
                            {
                                t1 = "تیرسان ";
                            }
                            else if (item.type == 3)
                            {
                                t1 = "بدون یونیفرم";

                            }
                            else if (item.type == 2)
                            {
                                t1 = "یونیفرم ساده";

                            }
                        }




                        Row = new Row() { Height = 20, Index = 0 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
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


                    foreach (var item in equipments)
                    {
                        string allinfo = "";
                        foreach (var item2 in item.tbEquipmentSpecificationData)
                        {
                            allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                        }
                        var c = "";

                        if (item.Users != null)
                        {
                            var model2 = item.Users.Split(',');
                            foreach (var item3 in model2)
                            {
                                var id = System.Convert.ToInt32(item3);
                                var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                                c += temp.FullName + " , ";

                            }
                        }
                        var t = "";
                        if (item.personal != null)
                        {
                            if (item.personal == true)
                            {
                                t = "پرسنلی";
                            }
                            else
                            {
                                t = "کارگزاری";

                            }
                        }
                        var t1 = "";
                        if (item.type != null)
                        {
                            if (item.type == 1)
                            {
                                t1 = "تیرسان ";
                            }
                            else if (item.type == 2)
                            {
                                t1 = "غیر تیرسان";

                            }
                            else if (item.type == 3)
                            {
                                t1 = "یونیفرم ساده";

                            }
                        }




                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+" نام:"+item.tbEquipmentBunch.Eqpbnch_Name+ "---"+"نام تحویل گیرندگان  :  "+c+ "----"+"نوع مالیکت:"+ t+"----"+"  نوع ماشین:  "+  t1+"----"+"نام پیمان  "+pynm22.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                        count3++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }



                }
            }
            var user = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
            foreach (var item in user)
            {
                Row = new Row() { Height = 20, Index = count1 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.FullName,
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
                        Value = item.usr_Personal_ID,
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
                count1++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }

            foreach (var item in pymn)
            {
                Row = new Row() { Height = 20, Index = count2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 2
                    },
                         new Cell()
                    {
                        Value = item.pec_ProjectCode,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 3
                    },
                });
                count2++;
                Moalefeexcelfile.Sheets[1].AddRow(Row);
            }

            return Moalefeexcelfile;

        }



        private Workbook SetDataToMachinsOrToolsExcel408032773(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
            var equipments = db.tbEquipments
                .Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type)
                .ToList();
            var pymn2 = db.tbPeymanContracts
                .Where(p => p.pec_ID == peyman_ID && p.Inactive != true)
                .FirstOrDefault();
            var company = db.tbCompanies
                .Where(p => p.ID == pymn2.FK_KarfarmaID)
                .FirstOrDefault();

            if (!equipments.Any())
                return Moalefeexcelfile;

            int count1 = 4, count3 = 2;
            AddHeaderRows(Moalefeexcelfile, company, pymn2);

            var user2 = db.Link_User_And_Peyman
                .Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true)
                .ToList();

            foreach (var equipment in equipments)
            {
                string allinfo = GetEquipmentSpecifications(equipment);
                string users = GetUsersNames(equipment);
                string typeDesc = GetTypeDescription(equipment.type);

                Row row = CreateRow(equipment.ID, allinfo, users, typeDesc, count3);
                Moalefeexcelfile.Sheets[0].AddRow(row);
                count3++;
            }

            double count334 = 0;

            foreach (var user in user2)
            {
                Row row = CreateUserRow(user, equipments, Month, Year, ref count334, count1);
                Moalefeexcelfile.Sheets[0].AddRow(row);
                count1++;
            }

            return Moalefeexcelfile;
        }

        private void AddHeaderRows(Workbook workbook, tbCompanies company, tbPeymanContracts pymn2)
        {
            // افزودن ردیف‌های هدر
            Row row1 = new Row() { Height = 20, Index = 2 };
            row1.AddCells(new List<Cell>
    {
        new Cell { Value = company.CompanyName, FontFamily = "B Nazanin", FontSize = 12, Index = 0 },
        new Cell { Value = company.Company_National_ID, FontFamily = "B Nazanin", FontSize = 12, Index = 1 }
    });
            workbook.Sheets[0].AddRow(row1);

            Row row2 = new Row() { Height = 20, Index = 3 };
            row2.AddCells(new List<Cell>
    {
        new Cell { Value = pymn2.pec_Title, FontFamily = "B Nazanin", FontSize = 12, Index = 0 },
        new Cell { Value = pymn2.pec_ProjectCode, FontFamily = "B Nazanin", FontSize = 12, Index = 1 }
    });
            workbook.Sheets[0].AddRow(row2);
        }

        private string GetEquipmentSpecifications(tbEquipments equipment)
        {
            var allinfo = string.Empty;
            foreach (var spec in equipment.tbEquipmentSpecificationData)
            {
                allinfo += spec.tbEquipmentSpecifications.sp_SpecTitle + ":" + spec.SpcData + "/";
            }
            return allinfo;
        }

        private string GetUsersNames(tbEquipments equipment)
        {
            if (string.IsNullOrEmpty(equipment.Users))
                return string.Empty;

            var userIds = equipment.Users.Split(',').Select(int.Parse).ToList();
            var users = string.Empty;
            foreach (var userId in userIds)
            {
                var user = db.tbUsers.FirstOrDefault(u => u.usr_ID == userId);
                if (user != null)
                {
                    users += user.FullName + ", ";
                }
            }
            return users.TrimEnd(' ', ',');
        }

        private string GetTypeDescription(int? type)
        {
            if (type == 1)
            {
                return "تیرسان";
            }
            else if (type == 2)
            {
                return "یونیفرم ساده";
            }
            else if (type == 3)
            {
                return "بدون یونیفرم";
            }
            return string.Empty;
        }

        private Row CreateRow(int id, string allinfo, string users, string typeDesc, int index)
        {
            var row = new Row { Height = 20, Index = 0 };
            row.AddCells(new List<Cell>
    {
        new Cell { Value = id, FontFamily = "B Nazanin", FontSize = 12, Index = index },
        new Cell { Value = allinfo, FontFamily = "B Nazanin", FontSize = 12, Index = index + 1 },
        new Cell { Value = users, FontFamily = "B Nazanin", FontSize = 12, Index = index + 2 },
        new Cell { Value = typeDesc, FontFamily = "B Nazanin", FontSize = 12, Index = index + 3 }
    });
            return row;
        }

        private Row CreateUserRow(Link_User_And_Peyman user, List<tbEquipments> equipments, int month, int year, ref double count334, int rowIndex)
        {
            var row = new Row { Height = 20, Index = rowIndex };
            row.AddCell(new Cell { Value = user.tbUsers.FullName, FontFamily = "B Nazanin", FontSize = 12, Index = 0 });

            int cellIndex = 2;
            foreach (var equipment in equipments)
            {
                double val = db.tbEquipmentMoalefeValue
                    .Where(ev => ev.FK_Equipment == equipment.ID && ev.Fk_user == user.FK_User_ID && ev.Month == month && ev.Year == year)
                    .Sum(ev => ev.CountDays ?? 0);

                count334 += val;
                row.AddCell(new Cell { Value = val, FontFamily = "B Nazanin", FontSize = 12, Index = cellIndex });
                cellIndex++;
            }

            return row;
        }

        private Workbook SetDataToMachinsOrToolsExcel4080(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
            Row Row;
            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
            var pynm22 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
            if (equipments.Count() > 0)
            {
                int counter = 1;
                var count = 2;
                var counttt = 2;

                var count1 = 4;
                var count2 = 1;

                var count3 = 2;


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }




                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
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


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }

                    if (type == 3)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----"+"  نوع ماشین:  "+  t1,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }
                    else
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----",
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }



                    count3++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }
                var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true).ToList();
                var pymn2 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
                var company = db.tbCompanies.Where(p => p.ID == pymn2.FK_KarfarmaID).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = company.CompanyName,
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
                        Value = company.Company_National_ID,
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

                Row = new Row() { Height = 20, Index = 3 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = pymn2.pec_Title,
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
                        Value = pymn2.pec_ProjectCode,
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

                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = item.tbUsers.FullName,
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
            Value = item.tbUsers.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month
                                     && p.Year == Year)
                            .ToList();

                        // جمع‌آوری مقدار CountDays
                        foreach (var tt in find)
                        {
                            val += tt.CountDays ?? 0;
                        }

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }

                    count1++; // افزایش ایندکس ردیف‌ها

                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }
                    double val = 0;
                    var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                    foreach (var tt in find)
                    {
                        val += tt.CountDays ?? 0;
                    }


                    Row = new Row() { Height = 20, Index = 3 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = val,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt
                    },
                });
                    counttt++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }










                var user = db.tbUsers.Where(p => p.usr_Personal_ID != null).ToList();
                var pymn = db.tbPeymanContracts.Where(p => p.Inactive == null).ToList();
                foreach (var item in user)
                {
                    Row = new Row() { Height = 20, Index = count1 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.FullName,
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
                        Value = item.usr_Personal_ID,
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
                    count1++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }

                foreach (var item in pymn)
                {
                    Row = new Row() { Height = 20, Index = count2 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.pec_Title,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 2
                    },
                         new Cell()
                    {
                        Value = item.pec_ProjectCode,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = 3
                    },
                });
                    count2++;
                    Moalefeexcelfile.Sheets[1].AddRow(Row);
                }


            }
            return Moalefeexcelfile;

        }
        private Workbook SetDataToMachinsOrToolsExcel4080323(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzar.xlsx"));
            Row Row;
            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
            var pynm22 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
            if (equipments.Count() > 0)
            {
                int counter = 1;
                var count = 2;
                var counttt = 2;
                var counttt22 = 2;

                var count1 = 4;
                var count2 = 1;

                var count3 = 2;
                List<Tuple<int, double>> tedada = new List<Tuple<int, double>>();
                Dictionary<int, double> tedada2 = new Dictionary<int, double>();

                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }




                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
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


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }

                    if (type == 3)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----"+"  نوع ماشین:  "+  t1,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }
                    else
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----",
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }



                    count3++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }
                var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true).ToList();
                var pymn2 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
                var company = db.tbCompanies.Where(p => p.ID == pymn2.FK_KarfarmaID).FirstOrDefault();

                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = company.CompanyName,
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
                        Value = company.Company_National_ID,
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

                Row = new Row() { Height = 20, Index = 3 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = pymn2.pec_Title,
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
                        Value = pymn2.pec_ProjectCode,
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
                double count334 = 0;

                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = item.tbUsers.FullName,
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
            Value = item.tbUsers.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year)
                            .ToList();

                        // جمع‌آوری مقدار CountDays
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;

                            val += tt.CountDays ?? 0;
                            if (tedada2.ContainsKey(itemmm.ID))
                            {
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                                tedada2[itemmm.ID] += tt.CountDays ?? 0;
                            }
                            else
                            {
                                // اضافه کردن یک مقدار جدید
                                tedada2.Add(itemmm.ID, tt.CountDays ?? 0);
                            }
                        }

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }

                    count1++; // افزایش ایندکس ردیف‌ها

                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }
                    double val = 0;
                    var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                    foreach (var tt in find)
                    {
                        count334 += tt.CountDays ?? 0;

                        val += tt.CountDays ?? 0;
                        if (tedada2.ContainsKey(item.ID))
                        {
                            // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            tedada2[item.ID] += tt.CountDays ?? 0;
                        }
                        else
                        {
                            // اضافه کردن یک مقدار جدید
                            tedada2.Add(item.ID, tt.CountDays ?? 0);
                        }
                        //tedada.Add(new Tuple<int, double>(item.ID, tt.CountDays ?? 0));

                    }


                    Row = new Row() { Height = 20, Index = 3 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = val,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt
                    },
                });
                    counttt++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }





                foreach (var item in equipments)
                {
                    //string allinfo = "";
                    //foreach (var item2 in item.tbEquipmentSpecificationData)
                    //{
                    //    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    //}
                    //var c = "";

                    //if (item.Users != null)
                    //{
                    //    var model2 = item.Users.Split(',');
                    //    foreach (var item3 in model2)
                    //    {
                    //        var id = System.Convert.ToInt32(item3);
                    //        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                    //        c += temp.FullName + " , ";

                    //    }
                    //}
                    //var t = "";
                    //if (item.personal != null)
                    //{
                    //    if (item.personal == true)
                    //    {
                    //        t = "پرسنلی";
                    //    }
                    //    else
                    //    {
                    //        t = "کارگزاری";

                    //    }
                    //}
                    //var t1 = "";
                    //if (item.type != null)
                    //{
                    //    if (item.type == 1)
                    //    {
                    //        t1 = "تیرسان ";
                    //    }
                    //    else if (item.type == 3)
                    //    {
                    //        t1 = "بدون یونیفرم";

                    //    }
                    //    else if (item.type == 2)
                    //    {
                    //        t1 = "یونیفرم ساده";

                    //    }
                    //}
                    //double val = 0;
                    //var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                    //foreach (var tt in find)
                    //{
                    //    count334 += tt.CountDays ?? 0;

                    //    val += tt.CountDays ?? 0;
                    //}


                    Row = new Row() { Height = 20, Index = count1 };
                    Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "رقم واحد",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

});
                    //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();
                    double value = (double)item.EachValue;


                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                    counttt22++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }

                count1++;
                foreach (var item in equipments)
                {
                    //string allinfo = "";
                    //foreach (var item2 in item.tbEquipmentSpecificationData)
                    //{
                    //    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    //}
                    //var c = "";

                    //if (item.Users != null)
                    //{
                    //    var model2 = item.Users.Split(',');
                    //    foreach (var item3 in model2)
                    //    {
                    //        var id = System.Convert.ToInt32(item3);
                    //        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                    //        c += temp.FullName + " , ";

                    //    }
                    //}
                    //var t = "";
                    //if (item.personal != null)
                    //{
                    //    if (item.personal == true)
                    //    {
                    //        t = "پرسنلی";
                    //    }
                    //    else
                    //    {
                    //        t = "کارگزاری";

                    //    }
                    //}
                    //var t1 = "";
                    //if (item.type != null)
                    //{
                    //    if (item.type == 1)
                    //    {
                    //        t1 = "تیرسان ";
                    //    }
                    //    else if (item.type == 3)
                    //    {
                    //        t1 = "بدون یونیفرم";

                    //    }
                    //    else if (item.type == 2)
                    //    {
                    //        t1 = "یونیفرم ساده";

                    //    }
                    //}
                    //double val = 0;
                    //var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                    //foreach (var tt in find)
                    //{
                    //    count334 += tt.CountDays ?? 0;

                    //    val += tt.CountDays ?? 0;
                    //}


                    Row = new Row() { Height = 20, Index = count1 };
                    Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "رقم واحد",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

});
                    //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();
                    double value = (double)item.EachValue;


                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                    counttt22++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }

                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
   Value = "جمع کل",
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },

    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {

                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year)
                            .ToList();
                        var find2 = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == itemmm.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        double val1 = 0; // مقدار برای find
                        double val2 = 0; // مقدار برای find2

                        // حلقه مربوط به find
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;
                            val1 += tt.CountDays ?? 0; // ذخیره مقدار find
                        }

                        // حلقه مربوط به find2
                        foreach (var tt in find2)
                        {
                            count334 += tt.CountDays ?? 0;
                            val2 += tt.CountDays ?? 0; // ذخیره مقدار find2
                        }

                        // جمع مقادیر val1 و val2
                        val = val1 + val2;
                        // جمع‌آوری مقدار CountDays
                        //foreach (var tt in find)
                        //{
                        //    count334 += tt.CountDays ?? 0;

                        //    val += tt.CountDays ?? 0;
                        //}
                        //foreach (var tt in find2)
                        //{
                        //    count334 += tt.CountDays ?? 0;

                        //    val += tt.CountDays ?? 0;
                        //}

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }


                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                count1++;
                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
   Value = "جمع کل",
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },

    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year)
                            .ToList();
                        var find2 = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == itemmm.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        double val1 = 0; // مقدار برای find
                        double val2 = 0; // مقدار برای find2

                        // حلقه مربوط به find
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;
                            val1 += tt.CountDays ?? 0; // ذخیره مقدار find
                        }

                        // حلقه مربوط به find2
                        foreach (var tt in find2)
                        {
                            count334 += tt.CountDays ?? 0;
                            val2 += tt.CountDays ?? 0; // ذخیره مقدار find2
                        }

                        // جمع مقادیر val1 و val2
                        val = val1 + val2;
                        val *= (double)itemmm.EachValue;

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }


                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }
                count1++;
                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
   Value = "جمع کل",
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },

    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year)
                            .ToList();
                        var find2 = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == itemmm.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        double val1 = 0; // مقدار برای find
                        double val2 = 0; // مقدار برای find2

                        // حلقه مربوط به find
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;
                            val1 += tt.CountDays ?? 0; // ذخیره مقدار find
                        }

                        // حلقه مربوط به find2
                        foreach (var tt in find2)
                        {
                            count334 += tt.CountDays ?? 0;
                            val2 += tt.CountDays ?? 0; // ذخیره مقدار find2
                        }

                        // جمع مقادیر val1 و val2
                        val = val1 + val2;
                        val *= (double)itemmm.EachValue;

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val1,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }


                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }

                count1++;
                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
   Value = "جمع کل",
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 0
        },

    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = db.tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year)
                            .ToList();
                        var find2 = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == itemmm.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        double val1 = 0; // مقدار برای find
                        double val2 = 0; // مقدار برای find2

                        // حلقه مربوط به find
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;
                            val1 += tt.CountDays ?? 0; // ذخیره مقدار find
                        }

                        // حلقه مربوط به find2
                        foreach (var tt in find2)
                        {
                            count334 += tt.CountDays ?? 0;
                            val2 += tt.CountDays ?? 0; // ذخیره مقدار find2
                        }

                        // جمع مقادیر val1 و val2
                        val = val1 + val2;
                        val *= (double)itemmm.EachValue;

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val2,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }


                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }


            }
            return Moalefeexcelfile;

        }
        private Workbook SetDataToMachinsOrToolsExcel408032trst3(int peyman_ID, int type, int Month, int Year = 0)
        {
            var Moalefeexcelfile = Workbook.Load(Server.MapPath("~/Content/ExcelFiles/abzargzaresh.xlsx"));
            Row Row;
            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
            var pynm22 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
            if (equipments.Count() > 0)
            {
                int counter = 1;
                var count = 2;
                var counttt = 2;
                var counttt22 = 2;

                var count1 = 4;
                var count2 = 1;

                var count3 = 2;
                List<Tuple<int, double>> tedada = new List<Tuple<int, double>>();
                Dictionary<int, double> tedada2 = new Dictionary<int, double>();

                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }




                    Row = new Row() { Height = 20, Index = 0 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = item.ID,
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


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }

                    if (type == 3)
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----"+"  نوع ماشین:  "+  t1,
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }
                    else
                    {
                        Row = new Row() { Height = 20, Index = 1 };
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value ="نام وسیله: "+item.tbEquipmentBunch.Eqpbnch_Name +"-----",
                        FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = count3
                    },
                });
                    }



                    count3++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }
                var user2 = db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == peyman_ID && p.tbUsers.usr_Personal_ID != null && p.Status == true && p.tbPeymanContracts.Inactive != true).ToList();
                var pymn2 = db.tbPeymanContracts.Where(p => p.pec_ID == peyman_ID && p.Inactive != true).FirstOrDefault();
                var company = db.tbCompanies.Where(p => p.ID == pymn2.FK_KarfarmaID).FirstOrDefault();
                var tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.ToList();
                Row = new Row() { Height = 20, Index = 2 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = company.CompanyName,
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
                        Value = company.Company_National_ID,
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

                Row = new Row() { Height = 20, Index = 3 };
                Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = pymn2.pec_Title,
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
                        Value = pymn2.pec_ProjectCode,
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
                double count334 = 0;

                foreach (var item in user2)
                {
                    // ایجاد ردیف جدید
                    Row = new Row() { Height = 20, Index = count1 };

                    // اضافه کردن سلول‌های مربوط به کاربر
                    Row.AddCells(new List<Cell>()
    {
        new Cell()
        {
            Value = item.tbUsers.FullName,
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
            Value = item.tbUsers.usr_Personal_ID,
            FontFamily = "B Nazanin",
            Bold = false,
            Enable = true,
            Wrap = false,
            FontSize = 12,
            Italic = false,
            Underline = false,
            Index = 1
        }
    });

                    int coutyyu = 2;

                    // حلقه مربوط به تجهیزات
                    foreach (var itemmm in equipments)
                    {
                        double val = 0;
                        var find = tbEquipmentMoalefeValue
                            .Where(p => p.FK_Equipment == itemmm.ID
                                     && p.Fk_user == item.FK_User_ID
                                     && p.Month == Month - 1
                                     && p.Year == Year
                                     && p.tbEquipmentMoalefeValueReffrenceSave.FK_User == 1826
                                     && p.tbEquipmentMoalefeValueReffrenceSave.Final_Accept == true

                                     )
                            .ToList();

                        // جمع‌آوری مقدار CountDays
                        foreach (var tt in find)
                        {
                            count334 += tt.CountDays ?? 0;

                            val += tt.CountDays ?? 0;
                            if (tedada2.ContainsKey(itemmm.ID))
                            {
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                                tedada2[itemmm.ID] += tt.CountDays ?? 0;
                            }
                            else
                            {
                                // اضافه کردن یک مقدار جدید
                                tedada2.Add(itemmm.ID, tt.CountDays ?? 0);
                            }
                        }

                        // اضافه کردن سلول به ردیف مربوط به تجهیزات
                        Row.AddCell(new Cell()
                        {
                            Value = val,
                            //FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = coutyyu
                        });

                        coutyyu++; // افزایش ایندکس ستون‌ها
                    }

                    count1++; // افزایش ایندکس ردیف‌ها

                    // اضافه کردن ردیف کامل به فایل اکسل بعد از تکمیل حلقه تجهیزات
                    Moalefeexcelfile.Sheets[0].AddRow(Row);
                }


                foreach (var item in equipments)
                {
                    string allinfo = "";
                    foreach (var item2 in item.tbEquipmentSpecificationData)
                    {
                        allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                    }
                    var c = "";

                    if (item.Users != null)
                    {
                        var model2 = item.Users.Split(',');
                        foreach (var item3 in model2)
                        {
                            var id = System.Convert.ToInt32(item3);
                            var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                            c += temp.FullName + " , ";

                        }
                    }
                    var t = "";
                    if (item.personal != null)
                    {
                        if (item.personal == true)
                        {
                            t = "پرسنلی";
                        }
                        else
                        {
                            t = "کارگزاری";

                        }
                    }
                    var t1 = "";
                    if (item.type != null)
                    {
                        if (item.type == 1)
                        {
                            t1 = "تیرسان ";
                        }
                        else if (item.type == 3)
                        {
                            t1 = "بدون یونیفرم";

                        }
                        else if (item.type == 2)
                        {
                            t1 = "یونیفرم ساده";

                        }
                    }
                    double val = 0;
                    var find = tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null && p.tbEquipmentMoalefeValueReffrenceSave.Final_Accept == true).ToList();
                    foreach (var tt in find)
                    {
                        count334 += tt.CountDays ?? 0;

                        val += tt.CountDays ?? 0;
                        if (tedada2.ContainsKey(item.ID))
                        {
                            // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            tedada2[item.ID] += tt.CountDays ?? 0;
                        }
                        else
                        {
                            // اضافه کردن یک مقدار جدید
                            tedada2.Add(item.ID, tt.CountDays ?? 0);
                        }
                        //tedada.Add(new Tuple<int, double>(item.ID, tt.CountDays ?? 0));

                    }


                    Row = new Row() { Height = 20, Index = 3 };
                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value = val,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt
                    },
                });
                    counttt++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }


                int counttt2233 = 2;

                foreach (var item in equipments)
                {


                    Row = new Row() { Height = 20, Index = count1 };
                    Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "جمع کل",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0;
                    if (tedada2.ContainsKey(item.ID))
                    {
                        // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                        value = tedada2[item.ID];
                    }                    //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();


                    Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt2233
                    },
                });
                    counttt2233++;
                    Moalefeexcelfile.Sheets[0].AddRow(Row);









                }


                count1++;

                if (type == 3)
                {
                    foreach (var item in equipments)
                    {



                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = " رقم واحد اجاره",
        //FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

});
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();
                        double value = 0.0;
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();
                        if (item.Unit == 3)
                        {
                            if (item.egareh != null)
                            {
                                value = (double)(item.egareh * 12 / 366);

                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (item.egareh != null)
                            {
                                value = (double)(item.egareh);

                            }
                        }
                        else
                        {
                            if (item.egareh != null)
                            {
                                value = (double)(item.egareh * 12 / 366);

                            }
                        }


                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                        counttt22++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }
                    count1++;
                    counttt22 = 2;
                    foreach (var item in equipments)
                    {
                        //string allinfo = "";
                        //foreach (var item2 in item.tbEquipmentSpecificationData)
                        //{
                        //    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                        //}
                        //var c = "";

                        //if (item.Users != null)
                        //{
                        //    var model2 = item.Users.Split(',');
                        //    foreach (var item3 in model2)
                        //    {
                        //        var id = System.Convert.ToInt32(item3);
                        //        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                        //        c += temp.FullName + " , ";

                        //    }
                        //}
                        //var t = "";
                        //if (item.personal != null)
                        //{
                        //    if (item.personal == true)
                        //    {
                        //        t = "پرسنلی";
                        //    }
                        //    else
                        //    {
                        //        t = "کارگزاری";

                        //    }
                        //}
                        //var t1 = "";
                        //if (item.type != null)
                        //{
                        //    if (item.type == 1)
                        //    {
                        //        t1 = "تیرسان ";
                        //    }
                        //    else if (item.type == 3)
                        //    {
                        //        t1 = "بدون یونیفرم";

                        //    }
                        //    else if (item.type == 2)
                        //    {
                        //        t1 = "یونیفرم ساده";

                        //    }
                        //}
                        //double val = 0;
                        //var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        //foreach (var tt in find)
                        //{
                        //    count334 += tt.CountDays ?? 0;

                        //    val += tt.CountDays ?? 0;
                        //}


                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = " رقم واحد سوخت",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0.0;
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();
                        //if (item.Soght != null)
                        //{
                        //    value = (double)(item.Soght * 12 / 366);

                        //}
                        if (item.Unit == 3)
                        {
                            if (item.Soght != null)
                            {
                                value = (double)(item.Soght * 12 / 366);

                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (item.Soght != null)
                            {
                                value = (double)(item.Soght);

                            }
                        }
                        else
                        {
                            if (item.Soght != null)
                            {
                                value = (double)(item.Soght * 12 / 366);

                            }
                        }
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                        counttt22++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }
                    count1++;
                    counttt22 = 2;

                    foreach (var item in equipments)
                    {
                        //string allinfo = "";
                        //foreach (var item2 in item.tbEquipmentSpecificationData)
                        //{
                        //    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                        //}
                        //var c = "";

                        //if (item.Users != null)
                        //{
                        //    var model2 = item.Users.Split(',');
                        //    foreach (var item3 in model2)
                        //    {
                        //        var id = System.Convert.ToInt32(item3);
                        //        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                        //        c += temp.FullName + " , ";

                        //    }
                        //}
                        //var t = "";
                        //if (item.personal != null)
                        //{
                        //    if (item.personal == true)
                        //    {
                        //        t = "پرسنلی";
                        //    }
                        //    else
                        //    {
                        //        t = "کارگزاری";

                        //    }
                        //}
                        //var t1 = "";
                        //if (item.type != null)
                        //{
                        //    if (item.type == 1)
                        //    {
                        //        t1 = "تیرسان ";
                        //    }
                        //    else if (item.type == 3)
                        //    {
                        //        t1 = "بدون یونیفرم";

                        //    }
                        //    else if (item.type == 2)
                        //    {
                        //        t1 = "یونیفرم ساده";

                        //    }
                        //}
                        //double val = 0;
                        //var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        //foreach (var tt in find)
                        //{
                        //    count334 += tt.CountDays ?? 0;

                        //    val += tt.CountDays ?? 0;
                        //}


                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = " رقم واحد یونیفرم",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0.0;
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();

                        if (item.Unit == 3)
                        {
                            if (item.uniform != null)
                            {
                                value = (double)(item.uniform * 12 / 366);

                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (item.uniform != null)
                            {
                                value = (double)(item.uniform);

                            }
                        }
                        else
                        {
                            if (item.uniform != null)
                            {
                                value = (double)(item.uniform * 12 / 366);

                            }
                        }

                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                        counttt22++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }

                }
                else
                {
                    foreach (var item in equipments)
                    {
                        //string allinfo = "";
                        //foreach (var item2 in item.tbEquipmentSpecificationData)
                        //{
                        //    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                        //}
                        //var c = "";

                        //if (item.Users != null)
                        //{
                        //    var model2 = item.Users.Split(',');
                        //    foreach (var item3 in model2)
                        //    {
                        //        var id = System.Convert.ToInt32(item3);
                        //        var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                        //        c += temp.FullName + " , ";

                        //    }
                        //}
                        //var t = "";
                        //if (item.personal != null)
                        //{
                        //    if (item.personal == true)
                        //    {
                        //        t = "پرسنلی";
                        //    }
                        //    else
                        //    {
                        //        t = "کارگزاری";

                        //    }
                        //}
                        //var t1 = "";
                        //if (item.type != null)
                        //{
                        //    if (item.type == 1)
                        //    {
                        //        t1 = "تیرسان ";
                        //    }
                        //    else if (item.type == 3)
                        //    {
                        //        t1 = "بدون یونیفرم";

                        //    }
                        //    else if (item.type == 2)
                        //    {
                        //        t1 = "یونیفرم ساده";

                        //    }
                        //}
                        //double val = 0;
                        //var find = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == item.ID && p.Fk_pymn == pymn2.pec_ID && p.Month == Month && p.Year == Year && p.CountDays != 0 && p.CountDays != null).ToList();
                        //foreach (var tt in find)
                        //{
                        //    count334 += tt.CountDays ?? 0;

                        //    val += tt.CountDays ?? 0;
                        //}


                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "رقم واحد",
        //FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

});
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();

                        double value = 0.0;

                        if (item.Unit == 3)
                        {
                            value = (double)(item.EachValue * 12 / 366);

                        }
                        else if (item.Unit == 2)
                        {
                            if (item.EachValue != null)
                            {
                                value = (double)(item.EachValue);

                            }
                        }
                        else
                        {
                            value = (double)(item.EachValue * 12 / 366);

                        }
                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt22
                    },
                });
                        counttt22++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }

                }
                count1++;
                counttt2233 = 2;

                if (type == 3)
                {
                    foreach (var item in equipments)
                    {
                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "مبلغ کل اجاره ",
        //FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0;
                        if (item.Unit == 3)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.egareh != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.egareh * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.egareh != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.egareh));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.egareh != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.egareh * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();


                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt2233
                    },
                });
                        counttt2233++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }
                    count1++;

                    counttt2233 = 2;

                    foreach (var item in equipments)
                    {



                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "مبلغ کل سوخت ",

        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0;
                        if (item.Unit == 3)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.Soght != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.Soght * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.Soght != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.Soght));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.Soght != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.Soght * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();


                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        ////FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt2233
                    },
                });
                        counttt2233++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }
                    count1++;

                    counttt2233 = 2;

                    foreach (var item in equipments)
                    {



                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
                Value = "مبلغ کل یونیفرم ",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0;
                        if (item.Unit == 3)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.uniform != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.uniform * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.uniform != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.uniform));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.uniform != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.uniform * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();


                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt2233
                    },
                });
                        counttt2233++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }

                }
                else
                {
                    foreach (var item in equipments)
                    {



                        Row = new Row() { Height = 20, Index = count1 };
                        Row.AddCells(new List<Cell>()
{
    new Cell()
    {
        Value = "مبلغ کل",
        FontFamily = "B Nazanin",
        Bold = false,
        Enable = true,
        Wrap = false,
        FontSize = 12,
        Italic = false,
        Underline = false,
        Index = 0
    },

}); double value = 0;
                        if (item.Unit == 3)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.EachValue != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.EachValue * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else if (item.Unit == 2)
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.EachValue != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.EachValue));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        else
                        {
                            if (tedada2.ContainsKey(item.ID))
                            {
                                if (item.EachValue != null)
                                {
                                    value = (double)(tedada2[item.ID] * (item.EachValue * 12 / 366));

                                }
                                // اضافه کردن مقدار tt.CountDays به مقدار قبلی
                            }
                        }
                        //var find23 = db.tbEquipments.Where(p => p.ID == item.FK_Bunch && p.FK_Peyman == pymn2.pec_ID).FirstOrDefault();


                        Row.AddCells(new List<Cell>()
                {
                    new Cell()
                    {
                        Value =  value,
                        //FontFamily = "B Nazanin",
                        Bold = false,
                        Enable = true,
                        Wrap = false,
                        FontSize = 12,
                        Italic = false,
                        Underline = false,
                        Index = counttt2233
                    },
                });
                        counttt2233++;
                        Moalefeexcelfile.Sheets[0].AddRow(Row);









                    }

                }




            }
            return Moalefeexcelfile;

        }


        [AuthorizeAAA]
        private Workbook SetDataToMachinsOrToolsExcel(int peyman_ID, int type, int Month, int Year = 0)
        {
            var NamingTagssampleFile = new Workbook();

            if (type == 3) // وسیله نقلیه
            {
                NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/EquipmentMoalefeValueMachine2.xlsx"));

                Row Row;
                using (SaabEntities db = new SaabEntities())
                {
                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
                    if (equipments.Count() > 0)
                    {
                        int counter = 1;

                        foreach (var item in equipments)
                        {
                            string allinfo = "";
                            foreach (var item2 in item.tbEquipmentSpecificationData)
                            {
                                allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                            }
                            var c = "";

                            if (item.Users != null)
                            {
                                var model2 = item.Users.Split(',');
                                foreach (var item3 in model2)
                                {
                                    var id = System.Convert.ToInt32(item3);
                                    var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                                    c += temp.FullName + " , ";

                                }
                            }
                            var t = "";
                            if (item.personal != null)
                            {
                                if (item.personal == true)
                                {
                                    t = "پرسنلی";
                                }
                                else
                                {
                                    t = "کارگزاری";

                                }
                            }
                            var t1 = "";
                            if (item.type != null)
                            {
                                if (item.type == 1)
                                {
                                    t1 = "تیرسان ";
                                }
                                else if (item.type == 2)
                                {
                                    t1 = "غیر تیرسان";

                                }
                                else if (item.type == 3)
                                {
                                    t1 = "یونیفرم ساده";

                                }
                            }

                            Row = new Row() { Height = 20, Index = counter };

                            Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.ID,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 0
                            },new Cell()
                            {
                                Value = counter,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 1
                            },
                            new Cell
                            {
                                Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+" نام:"+item.tbEquipmentBunch.Eqpbnch_Name+ "---"+"نام تحویل گیرندگان: "+c+ "-"+"نوع مالیکت:"+t+"-"+"نوع ماشین"+t1,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 2
                            }
                            ,
                            new Cell
                            {
                                Value = allinfo,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 3
                            }
                        });
                            counter++;
                            NamingTagssampleFile.Sheets[0].AddRow(Row);
                        }
                    }
                }
                //                using (SaabEntities db = new SaabEntities())
                //                {
                //                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();



                //                    if (equipments.Count() > 0)
                //                    {
                //                        var firstEquipmentId = equipments.First().ID;
                //                        var exist = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == firstEquipmentId && p.Month == Month && p.Year == Year).Select(p => p.FK_tbEquipmentMoalefeValueReffrenceSave).ToList();            // or equipments[0]
                //                        List<tbEquipmentMoalefeValueReffrenceSave> exist_use = new List<tbEquipmentMoalefeValueReffrenceSave>();  // Replace YourType with the actual type of tbEquipmentMoalefeValueReffrenceSave

                //                        foreach (var itme3 in exist)
                //                        {
                //                            var itemsFromDatabase = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == itme3).ToList();
                //                            exist_use.AddRange(itemsFromDatabase);
                //                        }


                //                        int counter = 1;
                //                        int c = 2;

                //                        int startRowIndex = 2;

                //                        // Adjust this based on your actual start row
                //                        int endRowIndex = 2 + exist_use.Count() - 1;
                //                        foreach (var item in equipments)
                //                        {
                //                            string allinfo = "";
                //                            foreach (var item2 in item.tbEquipmentSpecificationData)
                //                            {
                //                                allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                //                            }
                //                            foreach (var item2 in exist_use)
                //                            {





                //                                Row = new Row() { Height = 20, Index = counter };

                //                                Row.AddCells(new List<Cell>()
                //              {
                //                  new Cell()
                //                  {
                //                      Value = item.ID,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 0
                //                  },new Cell()
                //                  {
                //                      Value = counter,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 1
                //                  },
                //                  new Cell
                //                  {
                //                      Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+"نام:"+item.tbEquipmentBunch.Eqpbnch_Name,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 2
                //                  }
                //                  ,
                //                  new Cell
                //                  {
                //                      Value = allinfo,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 3
                //                  },
                //                   new Cell
                //                  {
                //                      Value = item2.tbUsers.FullName,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 4
                //                  },
                //                 new Cell
                //{
                //    Value = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.CountDays)
                //                .FirstOrDefault(), // or .SingleOrDefault() depending on your data model
                //    FontFamily = "B Nazanin",
                //    Bold = false,
                //    Enable = true,
                //    Wrap = false,
                //    FontSize = 12,
                //    Italic = false,
                //    Underline = false,
                //    Index = 5
                //},

                //                     new Cell
                //                  {
                //                      Value  = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.Status)
                //                .FirstOrDefault(),
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 6
                //                  },
                //                      new Cell
                //                  {
                //                        Value  = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.Description)
                //                .FirstOrDefault(),
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 7
                //                  }
                //              });
                //                                counter++;
                //                                NamingTagssampleFile.Sheets[2].AddRow(Row);
                //                            }
                //                            var x = exist_use.Count();
                //                            if (x > 1)
                //                            {
                //                                // Adjust this based on your actual start row
                //                                // Adjust this based on your actual end row

                //                                NamingTagssampleFile.Sheets[2].AddMergedCells($"D{startRowIndex}:D{endRowIndex}");
                //                                NamingTagssampleFile.Sheets[2].AddMergedCells($"C{startRowIndex}:C{endRowIndex}");

                //                                startRowIndex += x;
                //                                endRowIndex += x;
                //                                // Adjust this based on your actual start row



                //                            }
                //                        }
                //                    }
                //                }
            }
            else if (type == 4)// ابزار کار و لوازم ایمنی
            {
                NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/EquipmentMoalefeValueTools.xlsx"));
                Row Row;
                using (SaabEntities db = new SaabEntities())
                {
                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
                    if (equipments.Count() > 0)
                    {
                        int counter = 1;

                        foreach (var item in equipments)
                        {
                            string allinfo = "";
                            foreach (var item2 in item.tbEquipmentSpecificationData)
                            {
                                allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                            }
                            var c = "";

                            if (item.Users != null)
                            {
                                var model2 = item.Users.Split(',');
                                foreach (var item3 in model2)
                                {
                                    var id = System.Convert.ToInt32(item3);
                                    var temp = db.tbUsers.Where(p => p.usr_ID == id).FirstOrDefault();
                                    c += temp.FullName + " , ";

                                }
                            }
                            var t = "";

                            if (item.personal != null)
                            {
                                if (item.personal == true)
                                {
                                    t = "پرسنلی";
                                }
                                else
                                {
                                    t = "کارگزاری";

                                }
                            }

                            Row = new Row() { Height = 20, Index = counter };

                            Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.ID,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 0
                            },new Cell()
                            {
                                Value = counter,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 1
                            },
                            new Cell
                            {
                                Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+"نام:"+item.tbEquipmentBunch.Eqpbnch_Name+ "-"+"نام تحویل گیرندگان: "+c+ "-"+"نوع مالیکت:"+t,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 2
                            }
                            ,
                            new Cell
                            {
                                Value = allinfo,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 3
                            }
                        });
                            counter++;
                            NamingTagssampleFile.Sheets[0].AddRow(Row);
                        }
                    }
                }
                //                using (SaabEntities db = new SaabEntities())
                //                {
                //                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();
                //                    if (equipments.Count() > 0)
                //                    {
                //                        var firstEquipmentId = equipments.First().ID;
                //                        var exist = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == firstEquipmentId && p.Month == Month && p.Year == Year).Select(p => p.FK_tbEquipmentMoalefeValueReffrenceSave).ToList();            // or equipments[0]
                //                        List<tbEquipmentMoalefeValueReffrenceSave> exist_use = new List<tbEquipmentMoalefeValueReffrenceSave>();  // Replace YourType with the actual type of tbEquipmentMoalefeValueReffrenceSave
                //                        int startRowIndex = 2;
                //                        foreach (var itme3 in exist)
                //                        {
                //                            var itemsFromDatabase = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == itme3).ToList();
                //                            exist_use.AddRange(itemsFromDatabase);
                //                        }
                //                        int endRowIndex = startRowIndex + exist_use.Count() - 1;

                //                        int counter = 1;

                //                        foreach (var item in equipments)
                //                        {
                //                            string allinfo = "";
                //                            foreach (var item2 in item.tbEquipmentSpecificationData)
                //                            {
                //                                allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                //                            }
                //                            foreach (var item2 in exist_use)
                //                            {

                //                                Row = new Row() { Height = 20, Index = counter };

                //                                Row.AddCells(new List<Cell>()
                //                        {
                //                            new Cell()
                //                            {
                //                                Value = item.ID,
                //                                FontFamily = "B Nazanin",
                //                                Bold = false,
                //                                Enable = true,
                //                                Wrap = false,
                //                                FontSize = 12,
                //                                Italic = false,
                //                                Underline = false,
                //                                Index = 0
                //                            },new Cell()
                //                            {
                //                                Value = counter,
                //                                FontFamily = "B Nazanin",
                //                                Bold = false,
                //                                Enable = true,
                //                                Wrap = false,
                //                                FontSize = 12,
                //                                Italic = false,
                //                                Underline = false,
                //                                Index = 1
                //                            },
                //                            new Cell
                //                            {
                //                                Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+"نام:"+item.tbEquipmentBunch.Eqpbnch_Name,
                //                                FontFamily = "B Nazanin",
                //                                Bold = false,
                //                                Enable = true,
                //                                Wrap = false,
                //                                FontSize = 12,
                //                                Italic = false,
                //                                Underline = false,
                //                                Index = 2
                //                            }
                //                            ,
                //                            new Cell
                //                            {
                //                                Value = allinfo,
                //                                FontFamily = "B Nazanin",
                //                                Bold = false,
                //                                Enable = true,
                //                                Wrap = false,
                //                                FontSize = 12,
                //                                Italic = false,
                //                                Underline = false,
                //                                Index = 3
                //                            },
                //                     new Cell
                //                  {
                //                      Value = item2.tbUsers.FullName,
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 4
                //                  },
                //                                    new Cell
                //{
                //    Value = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.NumberSavedInContracts)
                //                .FirstOrDefault(), // or .SingleOrDefault() depending on your data model
                //    FontFamily = "B Nazanin",
                //    Bold = false,
                //    Enable = true,
                //    Wrap = false,
                //    FontSize = 12,
                //    Italic = false,
                //    Underline = false,
                //    Index = 5
                //},
                //                                                   new Cell
                //{
                //    Value = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.NumberSavedInContractsThatAccepted)
                //                .FirstOrDefault(), // or .SingleOrDefault() depending on your data model
                //    FontFamily = "B Nazanin",
                //    Bold = false,
                //    Enable = true,
                //    Wrap = false,
                //    FontSize = 12,
                //    Italic = false,
                //    Underline = false,
                //    Index = 6
                //},





                //                 new Cell
                //{
                //    Value = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.CountDays)
                //                .FirstOrDefault(), // or .SingleOrDefault() depending on your data model
                //    FontFamily = "B Nazanin",
                //    Bold = false,
                //    Enable = true,
                //    Wrap = false,
                //    FontSize = 12,
                //    Italic = false,
                //    Underline = false,
                //    Index = 7
                //},

                //                     new Cell
                //                  {
                //                      Value  = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.Status)
                //                .FirstOrDefault(),
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 8
                //                  },
                //                      new Cell
                //                  {
                //                        Value  = db.tbEquipmentMoalefeValue
                //                .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item2.ID && p.FK_Equipment == item.ID)
                //                .Select(p => p.Description)
                //                .FirstOrDefault(),
                //                      FontFamily = "B Nazanin",
                //                      Bold = false,
                //                      Enable = true,
                //                      Wrap = false,
                //                      FontSize = 12,
                //                      Italic = false,
                //                      Underline = false,
                //                      Index = 9
                //                  }
                //              });
                //                                counter++;
                //                                NamingTagssampleFile.Sheets[2].AddRow(Row);
                //                            }
                //                            var x = exist_use.Count();
                //                            if (x > 1)
                //                            {
                //                                NamingTagssampleFile.Sheets[2].AddMergedCells($"D{startRowIndex}:D{endRowIndex}");
                //                                NamingTagssampleFile.Sheets[2].AddMergedCells($"C{startRowIndex}:C{endRowIndex}");

                //                                startRowIndex += x;
                //                                endRowIndex += x;
                //                            }
                //                        }
                //                    }
                //                }
            }


            return NamingTagssampleFile;
        }



        private Workbook SetDataToMachinsOrToolsExcel2(int peyman_ID, int type, DateTime Fromdata, DateTime Todate, int id)
        {
            var NamingTagssampleFile = new Workbook();
            if (type == 3) // وسیله نقلیه
            {
                NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/EquipmentMoalefeValueMachine.xlsx"));
                Row Row;

                using (SaabEntities db = new SaabEntities())
                {
                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();

                    if (equipments.Count() > 0)
                    {
                        foreach (var item in equipments)
                        {
                            int counter = 1; // Initialize the counter outside the inner loop
                            var now = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == id).ToList();

                            foreach (var it in now)
                            {
                                Row = new Row() { Height = 20, Index = counter }; // Create a new instance of Row for each iteration

                                string allinfo = "";
                                foreach (var item2 in item.tbEquipmentSpecificationData)
                                {
                                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                                }
                                var c = "";

                                if (item.Users != null)
                                {
                                    var model2 = item.Users.Split(',');
                                    foreach (var item3 in model2)
                                    {
                                        var id3 = System.Convert.ToInt32(item3);
                                        var temp = db.tbUsers.Where(p => p.usr_ID == id3).FirstOrDefault();
                                        c += temp.FullName + " , ";

                                    }
                                }
                                var t = "";
                                if (item.personal != null)
                                {
                                    if (item.personal == true)
                                    {
                                        t = "پرسنلی";
                                    }
                                    else
                                    {
                                        t = "کارگزاری";

                                    }
                                }
                                var t1 = "";
                                if (item.type != null)
                                {
                                    if (item.type == 1)
                                    {
                                        t1 = "تیرسان ";
                                    }
                                    else if (item.type == 2)
                                    {
                                        t1 = "غیر تیرسان";

                                    }
                                    else if (item.type == 3)
                                    {
                                        t1 = "یونیفرم ساده";

                                    }
                                }
                                Row.AddCells(new List<Cell>
                    {
                        new Cell
                        {
                            Value = item.ID,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 0
                        },
                        new Cell
                        {
                            Value = counter, // Use the counter here
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,

                            Italic = false,
                            Underline = false,
                            Index = 1
                        },
                        new Cell
                        {
                                Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+" نام:"+item.tbEquipmentBunch.Eqpbnch_Name+ "---"+"نام تحویل گیرندگان: "+c+ "-"+"نوع مالیکت:"+t+"-"+"نوع ماشین"+t1,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 2
                        },
                        new Cell
                        {
                            Value = allinfo,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 3
                        },
                        new Cell
                        {
                            Value = it.CountDays,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 4
                        },
                        new Cell
                        {
                            Value = it.Status,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 5
                        },
                        new Cell
                        {
                            Value = it.Description,
                            FontFamily = "B Nazanin",
                            Bold = false,
                            Enable = true,
                            Wrap = false,
                            FontSize = 12,
                            Italic = false,
                            Underline = false,
                            Index = 6
                        }
                    });

                                NamingTagssampleFile.Sheets[0].AddRow(Row);
                                counter++; // Increment the counter for each iteration of the inner loop
                            }
                        }
                    }
                }
            }

            else if (type == 4)// ابزار کار و لوازم ایمنی
            {
                NamingTagssampleFile = Workbook.Load(Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/EquipmentMoalefeValueTools.xlsx"));
                Row Row;
                using (SaabEntities db = new SaabEntities())
                {
                    var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == type).ToList();

                    if (equipments.Count() > 0)
                    {
                        foreach (var item in equipments)
                        {
                            int counter = 1;
                            var now = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == id).ToList();

                            foreach (var it in now)
                            {
                                Row = new Row() { Height = 20, Index = counter };
                                // Create a new instance of Row for each iteration



                                string allinfo = "";
                                foreach (var item2 in item.tbEquipmentSpecificationData)
                                {
                                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
                                }

                                Row.AddCells(new List<Cell>()
                        {
                            new Cell()
                            {
                                Value = item.ID,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 0
                            },new Cell()
                            {
                                Value = counter,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 1
                            },
                            new Cell
                            {
                                Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+"نام:"+item.tbEquipmentBunch.Eqpbnch_Name,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 2
                            }
                            ,
                            new Cell
                            {
                                Value = allinfo,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 3
                            }
                              ,
                            new Cell
                            {
                                Value = it.NumberSavedInContracts,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 4
                            }
                              ,
                            new Cell
                            {
                                Value = it.NumberSavedInContractsThatAccepted,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 5
                            }
                              ,
                            new Cell
                            {
                                Value = it.CountDays,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 6
                            }
                              ,
                            new Cell
                            {
                                Value = it.Status,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 7
                            }
                              ,
                            new Cell
                            {
                                Value = it.Description,
                                FontFamily = "B Nazanin",
                                Bold = false,
                                Enable = true,
                                Wrap = false,
                                FontSize = 12,
                                Italic = false,
                                Underline = false,
                                Index = 8
                            }

                        });

                                counter++;
                                NamingTagssampleFile.Sheets[0].AddRow(Row);
                            }

                        }
                    }
                }

            }


            return NamingTagssampleFile;
        }


        public string Cheak(int peymanID, int Type, int Month, int Year)
        {

            var c = db.tbSavedFunctions.Where(p => p.svdfunc_pymnID == peymanID && p.tbPeymanContracts.Inactive != true && p.svdfunc_BastehID == Type && p.tbMoalefeDastmozdiValueFromExcel.Any(s => s.MoalfeVal_Month == Month && s.MoalfeVal_Year == Year)).ToList();

            List<tbSavedFunctions> b = new List<tbSavedFunctions>();
            if (c.Count != 0)
            {
                foreach (var item in c)
                {

                    if (item.Final_Accept == true)
                    {
                        b.Add(item);

                    }



                }
                if (b.Count == 0)
                {
                    return "True";

                }
                else
                {
                    return "False";

                }
            }


            else
            {
                return "False2";
            }

        }


        [AuthorizeAAA]
        public async Task<ActionResult> UploadAttachmentFile(int peymanID, int Year, int Month, int Type, int FK_Baste = 0,
           IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {
            if (files == null)
            {
                return MessageBox.Show("فایل بطور صحیح بارگذاری نشده است", MessageType.Error, false, MessageAlignment.TopRight);
            }

            var Message = await GetDataFromExcelMachinsOrTools41(files.First(), Year, Month, Type, peymanID, FK_Baste);
            return Message;
        }
        public async Task<ActionResult> UploadAttachmentFile5(int peymanID, DateTime FromDate, DateTime Todate, int Year, int Month, int Type,
         IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null, int FK_BASTE = 0)
        {
            if (files == null)
            {
                return MessageBox.Show("فایل بطور صحیح بارگذاری نشده است", MessageType.Error, false, MessageAlignment.TopRight);
            }
            var segment = files.First().FileName.Split('.');
            string file_type = segment[segment.Length - 1].ToLower();
            if (file_type != "xlsx")
            {
                return MessageBox.Show("فایل با فایل اصلی مطابقت ندارد", MessageType.Error, false, MessageAlignment.TopRight);
            }
            var a = await GetDataFromExcelMachinsOrTools4(files.First(), FromDate, Todate, Year, Month, Type, peymanID, FK_BASTE);
            var b = "true";
            return a;

        }
        public async Task<ActionResult> UploadAttachmentFile5edit(int peymanID, int Year, int Month,
     IEnumerable<HttpPostedFileBase> files = null)
        {
            if (files == null)
            {
                return MessageBox.Show("فایل بطور صحیح بارگذاری نشده است", MessageType.Error, false, MessageAlignment.TopRight);
            }

            var a = await GetDataFromExcelMachinsOrTools4edit(files.First(), Year, Month, peymanID);
            var b = "true";
            return a;

        }



        [AuthorizeAAA]
        public string GetDataFromExcelMachinsOrTools(HttpPostedFileBase MyExcelStream,
            int Year, int Month, int Type, int peymanID, int FK_Baste)
        {



            var us = db.tbReffrenceSave
    .Where(p => p.FK_PeymanID == peymanID && p.IsFor == Type && p.tbPeymanContracts.Inactive != true)


     .OrderByDescending(p => p.ID)
     .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
     .Where(p => p.FK_RRSave == us.ID)
     .OrderByDescending(p => p.ID)
     .FirstOrDefault();

            var rus3 = db.tbReffrenceSaveLevelUser
          .Where(p => p.FK_LevelID == us2.ID)
          .OrderByDescending(p => p.ID)
          .Select(p => p.FK_UserID)
          .FirstOrDefault();
            var User = new tbUsers();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User != null)
                {

                    userid = User.usr_ID;

                }
            }

            using (SaabEntities db = new SaabEntities())
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    using (ExcelEngine xl = new ExcelEngine())
                    {
                        if (Type == 3)
                        {
                            try
                            {

                                IApplication app = xl.Excel;
                                var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                                var Count_Sheets = workbook.Worksheets.Count;
                                if (Count_Sheets > 1)
                                {
                                    var sheet = workbook.Worksheets[0];
                                    var Rows = sheet.Rows;
                                    int Cols = 0;
                                    if (Rows.Length >= 1)
                                    {
                                        Cols = Rows[0].Cells.Length;

                                        //تعداد ستون ها
                                        if (Count_Sheets != 3 || Cols != 7)
                                        {
                                            return "اکسل وارد شده با اکسل نمونه  متفاوت است";
                                        }

                                        var title = workbook.Worksheets[0].Rows[0].Cells;
                                        var count = workbook.Worksheets[0].Rows.Count();
                                        if (count == 1)
                                        {
                                            return "فایل اکسل فاقد اطلاعات می باشد";
                                        }

                                        Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                        int counter = 0;
                                        foreach (var item in title)
                                        {
                                            dicTitles.Add(item.Value.ToString(), counter);
                                            counter++;
                                        }
                                        tbEquipmentMoalefeValueReffrenceSave obj_EquipmentMoalefeValueReffrenceSav = new tbEquipmentMoalefeValueReffrenceSave();
                                        obj_EquipmentMoalefeValueReffrenceSav.DateTime = DateTime.Now;
                                        obj_EquipmentMoalefeValueReffrenceSav.FK_User = User.usr_ID;
                                        obj_EquipmentMoalefeValueReffrenceSav.For_Acceot = true;
                                        obj_EquipmentMoalefeValueReffrenceSav.FK_Baste = FK_Baste;











                                        var exist = db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == FK_Baste && p.FK_PeymanID == peymanID && p.tbPeymanContracts.Inactive != true).FirstOrDefault();
                                        if (exist != null)
                                        {
                                            var exist_us = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist.ID).OrderByDescending(p => p.ID).FirstOrDefault();
                                            if (exist_us != null)
                                            {
                                                var exist_final = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).OrderByDescending(p => p.ID).Select(p => p.FK_User).FirstOrDefault();
                                                if (exist_final == User.usr_ID)
                                                {
                                                    // در اینجا شرط برقرار است

                                                    // دریافت شئی متناظر با exist_final از جدول tbSavedFunctions


                                                    obj_EquipmentMoalefeValueReffrenceSav.Final_Accept = true;

                                                    var exxistpy = db.tbReffrenceSaveLevel.Where(p => p.ID == FK_Baste && p.Deleted != true).Select(p => p.FK_RRSave).FirstOrDefault();
                                                    if (exxistpy != null)
                                                    {
                                                        var existpy = db.tbReffrenceSave.Where(s => s.ID == exxistpy).Select(s => s.IsFor).FirstOrDefault();
                                                        if (existpy == 3)
                                                        {
                                                            //                  var matchedRows = db.tbMoalefeValuePishkhan
                                                            //.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان - پیمان - مستندات وضعیت - تایید ماشین آلات پیمان"&&p.mlfval_FKPeyman== peymanID&&p.mlfval_Month==Month&&p.mlfval_Year==Year)
                                                            //.FirstOrDefault();


                                                            //                  matchedRows.mlfval_Value = "1";


                                                            //                  db.SaveChanges();
                                                            //پیشخوان - پیمان - مستندات وضعیت - تایید ماشین آلات پیمان
                                                            //    پیشخوان - پیمان - مستندات وضعیت - تایید ماشین آلات پیمان
                                                        }
                                                    }






                                                    // ذخیره تغییرات در دیتابیس
                                                }
                                            }

                                        }




















                                        List<tbEquipmentMoalefeValue> myList = new List<tbEquipmentMoalefeValue>();
                                        for (int i = 1; i < count; i++)
                                        {
                                            tbEquipmentMoalefeValue equipment_obj = new tbEquipmentMoalefeValue();
                                            var row = workbook.Worksheets[0].Rows[i];
                                            var cell_code = row.Cells[0];

                                            if (cell_code != null)
                                            {
                                                if (cell_code.Value != null || cell_code.Value != "")
                                                {
                                                    equipment_obj.FK_Equipment = System.Convert.ToInt32(cell_code.Value);
                                                }
                                                else
                                                {
                                                    return "اکسل وارد شده با اکسل نمونه تفاوت دارد";
                                                }
                                            }
                                            else
                                            {
                                                return "اکسل وارد شده با اکسل نمونه تفاوت دارد";
                                            }

                                            var cell_countDays = row.Cells[4];
                                            if (cell_countDays != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_countDays.Value))
                                                {
                                                    equipment_obj.CountDays = System.Convert.ToInt32(cell_countDays.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون تعداد روز سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }

                                            var cell_status = row.Cells[1];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_status.Value))
                                                {
                                                    equipment_obj.Status = System.Convert.ToInt32(cell_status.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }
                                            else
                                            {
                                                return " لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }

                                            var cell_description = row.Cells[2];
                                            if (cell_description != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_description.Value))
                                                {

                                                    equipment_obj.Description = cell_description.Value;

                                                }
                                            }

                                            //equipment_obj.FromDate = FromDate;
                                            //equipment_obj.ToDate = ToDate;
                                            equipment_obj.Year = Year;
                                            equipment_obj.Month = Month;


                                            obj_EquipmentMoalefeValueReffrenceSav.tbEquipmentMoalefeValue.Add(equipment_obj);

                                        }

                                        if (MyExcelStream != null)
                                        {

                                            if (MyExcelStream.ContentLength > 0)
                                            {
                                                var segment = MyExcelStream.FileName.Split('.');
                                                string file_type = segment[segment.Length - 1];
                                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                                MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                                                obj_EquipmentMoalefeValueReffrenceSav.File_SystemNameForAccept = filename;
                                                obj_EquipmentMoalefeValueReffrenceSav.FileNameForAccept = MyExcelStream.FileName;
                                            }
                                        }

                                        db.tbEquipmentMoalefeValueReffrenceSave.Add(obj_EquipmentMoalefeValueReffrenceSav);
                                        db.SaveChanges();
                                        if (db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.Final_Accept == true && p.FK_Baste == FK_Baste && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).FirstOrDefault() != null)
                                        {
                                            foreach (var item in db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_Baste && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).ToList())
                                            {
                                                item.is_submit = 1;
                                                db.SaveChanges();
                                            }
                                        }
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        return "این شیت فاقد سطر می باشد";
                                    }

                                    return "با موفقیت ثبت شد";
                                }
                                else
                                {
                                    return "هیچ شیتی در این اکسل وجود ندارد";
                                }
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                return ex.Message;
                            }

                        }
                        else if (Type == 4)
                        {
                            try
                            {
                                IApplication app = xl.Excel;
                                var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                                var Count_Sheets = workbook.Worksheets.Count;
                                if (Count_Sheets > 1)
                                {
                                    var sheet = workbook.Worksheets[0];
                                    var Rows = sheet.Rows;
                                    int Cols = 0;
                                    if (Rows.Length >= 1)
                                    {
                                        Cols = Rows[0].Cells.Length;

                                        //تعداد ستون ها
                                        if (Count_Sheets != 3 || Cols != 9)
                                        {
                                            return "اکسل وارد شده با اکسل نمونه  متفاوت است";
                                        }

                                        var title = workbook.Worksheets[0].Rows[0].Cells;
                                        var count = workbook.Worksheets[0].Rows.Count();
                                        if (count == 1)
                                        {
                                            return "فایل اکسل فاقد اطلاعات می باشد";
                                        }

                                        Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                        int counter = 0;
                                        foreach (var item in title)
                                        {
                                            dicTitles.Add(item.Value.ToString(), counter);
                                            counter++;
                                        }
                                        tbEquipmentMoalefeValueReffrenceSave obj_EquipmentMoalefeValueReffrenceSav = new tbEquipmentMoalefeValueReffrenceSave();
                                        obj_EquipmentMoalefeValueReffrenceSav.DateTime = DateTime.Now;
                                        obj_EquipmentMoalefeValueReffrenceSav.FK_User = User.usr_ID;
                                        obj_EquipmentMoalefeValueReffrenceSav.For_Acceot = true;
                                        obj_EquipmentMoalefeValueReffrenceSav.FK_Baste = FK_Baste;


                                        var exist = db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == FK_Baste && p.FK_PeymanID == peymanID && p.tbPeymanContracts.Inactive != true).FirstOrDefault();
                                        if (exist != null)
                                        {
                                            var exist_us = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist.ID).OrderByDescending(p => p.ID).FirstOrDefault();
                                            if (exist_us != null)
                                            {
                                                var exist_final = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).OrderByDescending(p => p.ID).Select(p => p.FK_User).FirstOrDefault();
                                                if (exist_final == User.usr_ID)
                                                {
                                                    // در اینجا شرط برقرار است

                                                    // دریافت شئی متناظر با exist_final از جدول tbSavedFunctions


                                                    obj_EquipmentMoalefeValueReffrenceSav.Final_Accept = true;
                                                    var exxistpy = db.tbReffrenceSaveLevel.Where(p => p.ID == FK_Baste && p.Deleted != true).Select(p => p.FK_RRSave).FirstOrDefault();
                                                    if (exxistpy != null)
                                                    {
                                                        var existpy = db.tbReffrenceSave.Where(s => s.ID == exxistpy).Select(s => s.IsFor).FirstOrDefault();
                                                        if (existpy == 4)
                                                        {
                                                            //                                                            var matchedRows = db.tbMoalefeValuePishkhan
                                                            //.Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-تایید ابزار پیمان" && p.mlfval_FKPeyman == peymanID && p.mlfval_Month == Month && p.mlfval_Year == Year)
                                                            //.FirstOrDefault();


                                                            //                                                            matchedRows.mlfval_Value = "1";


                                                            //                                                            db.SaveChanges();
                                                            //پیشخوان - پیمان - مستندات وضعیت - تایید ماشین آلات پیمان
                                                            //    پیشخوان - پیمان - مستندات وضعیت - تایید ماشین آلات پیمان
                                                        }
                                                    }






                                                    // ذخیره تغییرات در دیتابیس
                                                }
                                            }

                                        }


                                        List<tbEquipmentMoalefeValue> myList = new List<tbEquipmentMoalefeValue>();
                                        for (int i = 1; i < count; i++)
                                        {
                                            tbEquipmentMoalefeValue equipment_obj = new tbEquipmentMoalefeValue();
                                            var row = workbook.Worksheets[0].Rows[i];
                                            var cell_code = row.Cells[0];

                                            if (cell_code != null)
                                            {
                                                if (cell_code.Value != null || cell_code.Value != "")
                                                {
                                                    equipment_obj.FK_Equipment = System.Convert.ToInt32(cell_code.Value);
                                                }
                                                else
                                                {
                                                    return "اکسل وارد شده با اکسل نمونه تفاوت دارد";
                                                }
                                            }
                                            else
                                            {
                                                return "اکسل وارد شده با اکسل نمونه تفاوت دارد";
                                            }

                                            var cell_countDays = row.Cells[6];
                                            if (cell_countDays != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_countDays.Value))
                                                {
                                                    equipment_obj.CountDays = System.Convert.ToInt32(cell_countDays.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون تعداد روز سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }

                                            var cell_status = row.Cells[7];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_status.Value))
                                                {
                                                    equipment_obj.Status = System.Convert.ToInt32(cell_status.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }
                                            else
                                            {
                                                return " لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }

                                            var cell_Accept = row.Cells[5];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_Accept.Value))
                                                {
                                                    equipment_obj.NumberSavedInContractsThatAccepted = System.Convert.ToInt32(cell_Accept.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون تعداد مورد تایید سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }
                                            else
                                            {
                                                return " لطفا ستون تعداد موردتایید سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }

                                            var contractNumber = row.Cells[4];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(contractNumber.Value))
                                                {
                                                    equipment_obj.NumberSavedInContracts = System.Convert.ToInt32(contractNumber.Value);
                                                }
                                                else
                                                {
                                                    return " لطفا ستون تعداد ثبت شده درقرارداد سطر " + " ( " + i + " ) " + "را پر کنید ";
                                                }
                                            }
                                            else
                                            {
                                                return " لطفا ستون تعداد ثبت شده در قراداد سطر " + " ( " + i + " ) " + "را پر کنید ";
                                            }

                                            var cell_description = row.Cells[8];
                                            if (cell_description != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_description.Value))
                                                {

                                                    equipment_obj.Description = cell_description.Value;

                                                }
                                            }

                                            equipment_obj.Year = Year;
                                            equipment_obj.Month = Month;

                                            obj_EquipmentMoalefeValueReffrenceSav.tbEquipmentMoalefeValue.Add(equipment_obj);
                                        }
                                        if (MyExcelStream != null)
                                        {

                                            if (MyExcelStream.ContentLength > 0)
                                            {
                                                var segment = MyExcelStream.FileName.Split('.');
                                                string file_type = segment[segment.Length - 1];
                                                var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                                                MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                                                obj_EquipmentMoalefeValueReffrenceSav.File_SystemNameForAccept = filename;
                                                obj_EquipmentMoalefeValueReffrenceSav.FileNameForAccept = MyExcelStream.FileName;
                                            }
                                        }
                                        var exist3 = db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == FK_Baste && p.FK_PeymanID == peymanID && p.tbPeymanContracts.Inactive != true).FirstOrDefault();
                                        if (exist3 != null)
                                        {
                                            var exist_us = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist.ID).OrderByDescending(p => p.ID).FirstOrDefault();
                                            if (exist_us != null)
                                            {
                                                var exist_final = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).OrderByDescending(p => p.ID).Select(p => p.FK_User).FirstOrDefault();
                                                if (exist_final == User.usr_ID)
                                                {
                                                    obj_EquipmentMoalefeValueReffrenceSav.Final_Accept = true;
                                                }
                                            }
                                        }
                                        db.tbEquipmentMoalefeValueReffrenceSave.Add(obj_EquipmentMoalefeValueReffrenceSav);
                                        //db.tbEquipmentMoalefeValue.AddRange(myList);
                                        db.SaveChanges();

                                        if (db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.Final_Accept == true && p.FK_Baste == FK_Baste && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).FirstOrDefault() != null)
                                        {
                                            foreach (var item in db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_Baste && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).ToList())
                                            {
                                                item.is_submit = 1;
                                                db.SaveChanges();
                                            }
                                        }
                                        transaction.Commit();
                                    }
                                    else
                                    {
                                        return "این شیت فاقد سطر می باشد";
                                    }

                                    return "با موفقیت ثبت شد";
                                }
                                else
                                {
                                    return "هیچ شیتی در این اکسل وجود ندارد";
                                }
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                return ex.Message;
                            }
                        }
                        else
                        {
                            return "ثبت برای این دسته از کارکرد ها تعریف نشده است";
                        }
                    }
                }
            }
        }











        //public ActionResult Getpyy(int Year, int Month, int Type, int peymanID, int FK_BASTE)
        //{
        //    using (SaabEntities db = new SaabEntities())
        //    {
        //        var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peymanID && p.tbEquipmentBunch.FK_EqpgrpID == Type).ToList();
        //        if (equipments.Count() > 0)
        //        {
        //            int counter = 1;

        //            foreach (var item in equipments)
        //            {
        //                string allinfo = "";
        //                foreach (var item2 in item.tbEquipmentSpecificationData)
        //                {
        //                    allinfo += item2.tbEquipmentSpecifications.sp_SpecTitle + ":" + item2.SpcData + "/";
        //                }
        //                Row = new Row() { Height = 20, Index = counter };

        //                Row.AddCells(new List<Cell>()
        //  {
        //      new Cell()
        //      {
        //          Value = item.ID,
        //          FontFamily = "B Nazanin",
        //          Bold = false,
        //          Enable = true,
        //          Wrap = false,
        //          FontSize = 12,
        //          Italic = false,
        //          Underline = false,
        //          Index = 0
        //      },new Cell()
        //      {
        //          Value = counter,
        //          FontFamily = "B Nazanin",
        //          Bold = false,
        //          Enable = true,
        //          Wrap = false,
        //          FontSize = 12,
        //          Italic = false,
        //          Underline = false,
        //          Index = 1
        //      },
        //      new Cell
        //      {
        //          Value = item.Name + "-"+ item.tbEquipmentSpecificationData.FirstOrDefault().tbEquipmentSpecifications.sp_SpecTitle+":"+item.tbEquipmentSpecificationData.FirstOrDefault().SpcData+"نام:"+item.tbEquipmentBunch.Eqpbnch_Name,
        //          FontFamily = "B Nazanin",
        //          Bold = false,
        //          Enable = true,
        //          Wrap = false,
        //          FontSize = 12,
        //          Italic = false,
        //          Underline = false,
        //          Index = 2
        //      }
        //      ,
        //      new Cell
        //      {
        //          Value = allinfo,
        //          FontFamily = "B Nazanin",
        //          Bold = false,
        //          Enable = true,
        //          Wrap = false,
        //          FontSize = 12,
        //          Italic = false,
        //          Underline = false,
        //          Index = 3
        //      } ; } } }
        //    var pyy = db.tbPeymanContracts.Where(p => p.pec_ID == peymanID).FirstOrDefault();
        //    var tool = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == Type).FirstOrDefault();
        //    var pymanname = pyy.pec_Title;
        //    var tools = tool.Eqpgrp_Name; if (Type == 3)

        //        {

        //            Virwforexcel equipment_obj = new Virwforexcel();


        //            equipment_obj.SabtHeader = new SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment.Header();  // اضافه کردن این خط برای مقداردهی اولیه
        //            equipment_obj.informationToolsss = new List<informationTools>();  // اضافه کردن این خط برای مقداردهی اولیه
        //            equipment_obj.FK_BASTE = FK_BASTE;

        //            //equipment_obj.SabtHeader.FromDate = FromDate;
        //            //equipment_obj.SabtHeader.ToDate = System.Convert.ToDateTime(ToDate);
        //            equipment_obj.SabtHeader.MoalfeVal_Year = Year;
        //            equipment_obj.SabtHeader.MoalfeVal_Month = Month;
        //            equipment_obj.SabtHeader.PeymanName = pymanname;
        //            equipment_obj.SabtHeader.Type = tools;
        //            var find = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefault();
        //        }

        //}








        public async Task<ActionResult> GetDataFromExcelMachinsOrTools41(HttpPostedFileBase MyExcelStream,
    int Year, int Month, int Type, int peymanID, int FK_BASTE)
        {

            var pyy = await db.tbPeymanContracts.Where(p => p.pec_ID == peymanID && p.Inactive != true).FirstOrDefaultAsync();
            var tool = await db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == Type).FirstOrDefaultAsync();
            var pymanname = pyy.pec_Title;
            var tools = tool.Eqpgrp_Name;
            //       var us = await db.tbReffrenceSave
            //           .Where(p => p.FK_PeymanID == peymanID && p.IsFor == Type)


            //            .OrderByDescending(p => p.ID)
            //            .FirstOrDefaultAsync();
            //       var us2 = await db.tbReffrenceSaveLevel
            //.Where(p => p.FK_RRSave == us.ID)
            //.OrderByDescending(p => p.ID)
            //.FirstOrDefaultAsync();

            //       var rus3 =await db.tbReffrenceSaveLevelUser
            //     .Where(p => p.FK_LevelID == us2.ID)
            //     .OrderByDescending(p => p.ID)
            //     .Select(p => p.FK_UserID)
            //     .FirstOrDefaultAsync();
            var User = new tbUsers();
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = await db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefaultAsync();
                }
                if (User != null)
                {

                    userid = User.usr_ID;

                }
            }
            var sabtt = await db.tbEquipmentMoalefeValueReffrenceSave
.Where(p => p.FK_Baste == FK_BASTE &&

          p.FK_User == userid &&

          p.tbEquipmentMoalefeValue
              .Any(m => m.Year == Year && m.Month == Month))
.FirstOrDefaultAsync();
            var find = await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefaultAsync();
            if (find != null)
            {
                return Content("این بسته در این ماه و سال  تایید نهایی شده است ");
            }
            if (sabtt != null)
            {
                return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
            }

            List<string> lstMoalefe = new List<string>();
            List<int> intt = new List<int>();

            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;
            tbEquipmentSpecificationData obj2 = new tbEquipmentSpecificationData();
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();
            List<tbEquipmentMoalefeValue> list = new List<tbEquipmentMoalefeValue>();

            tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
            objsave.FK_User = userid;
            objsave.DateTime = DateTime.Now;
            objsave.FK_Baste = FK_BASTE;


            //var firstFilter = Filters.First();
            //var fkBASTE = firstFilter.FK_BASTE;

            //            var rus33 = db.tbReffrenceSaveLevelUser
            //.Where(p => p.FK_LevelID == FK_BASTE)
            //.OrderByDescending(p => p.ID)
            //.Select(p => p.FK_UserID)
            //.FirstOrDefault();
            //            if (userid == rus3)
            //            {
            //                objsave.Final_Sabt = true;
            //            }
            var exist = await db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == FK_BASTE && p.FK_PeymanID == peymanID).FirstOrDefaultAsync();
            if (exist != null)
            {
                var exist3 = await db.tbReffrenceAccept.Where(p => p.FK_ReffrenceSaveLevel == FK_BASTE && p.FK_PeymanID == peymanID).FirstOrDefaultAsync();
                if (exist3 != null)
                {
                    var exist_us = await db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == exist.ID).OrderByDescending(p => p.ID).FirstOrDefaultAsync();
                    if (exist_us != null)
                    {
                        var exist_final = await db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == exist_us.ID).OrderByDescending(p => p.ID).Select(p => p.FK_User).FirstOrDefaultAsync();
                        if (exist_final == User.usr_ID)
                        {
                            objsave.Final_Accept = true;
                        }
                    }
                }
                //db.tbEquipmentMoalefeValue.AddRange(myList);
            }
            if (MyExcelStream != null)
            {

                if (MyExcelStream.ContentLength > 0)
                {
                    var segment = MyExcelStream.FileName.Split('.');
                    string file_type = segment[segment.Length - 1];
                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                    MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                    objsave.File_SystemNameForAccept = filename;
                    objsave.FileNameForAccept = MyExcelStream.FileName;
                }
            }



            db.tbEquipmentMoalefeValueReffrenceSave.Add(objsave);
            await db.SaveChangesAsync();





            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();


                            if (count == 1)
                            {
                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                            }

                            for (int i = 2; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }



                            var user = await db.tbUsers.Where(p => p.usr_Personal_ID != null).ToListAsync();
                            var pymn = await db.tbPeymanContracts.Where(p => p.Inactive == null).ToListAsync();

                            for (int i = 3; i < count; i++)
                            {
                                tbEquipmentMoalefeValue list2 = new tbEquipmentMoalefeValue();

                                var row = workbook.Worksheets[0].Rows[i];
                                var Name = row.Cells[1];


                                //int savedItemId = db.tbEquipments
                                //                    .OrderByDescending(p => p.ID)
                                //                    .Select(s => s.ID)
                                //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                                int shomarande = 2;
                                List<tbEquipmentMoalefeValue> objectsToAdd = new List<tbEquipmentMoalefeValue>();

                                //var Model = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();
                                foreach (var it in lstMoalefe)
                                {
                                    tbEquipmentMoalefeValue obj22 = new tbEquipmentMoalefeValue();
                                    if (int.TryParse(Name.Value.ToString(), out int vahed2))
                                    {
                                        var find2 = pymn.Where(p => p.pec_ProjectCode == vahed2).FirstOrDefault();

                                        var findusr = user.Where(p => p.usr_Personal_ID == vahed2).FirstOrDefault();
                                        if (find2 != null)
                                        {
                                            obj22.Fk_pymn = find2.pec_ID;
                                        }
                                        else
                                        {
                                            obj22.Fk_user = findusr.usr_ID;

                                        }
                                    }
                                    else
                                    {

                                        db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                                        await db.SaveChangesAsync();
                                        return Content(" فایل را دوباره بررسی بفرمایید ");
                                        //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                    }

                                    var Value = row.Cells[shomarande];
                                    var value = Value.Value ?? null;
                                    float? countDays;
                                    if (float.TryParse(value, out float parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }
                                    else
                                    {
                                        countDays = null; // Or handle the case where the value cannot be parsed
                                    }
                                    int? FK_Equipment = null;
                                    if (int.TryParse(it, out int FK_Equipment2))
                                    {
                                        FK_Equipment = FK_Equipment2;
                                    }
                                    else
                                    {
                                        FK_Equipment = null; // Or handle the case where the value cannot be parsed
                                    }

                                    obj22.CountDays = countDays;
                                    obj22.FK_Equipment = FK_Equipment;
                                    obj22.FK_tbEquipmentMoalefeValueReffrenceSave = objsave.ID;
                                    obj22.Month = Month;
                                    obj22.Year = Year;
                                    objectsToAdd.Add(obj22);
                                    shomarande++;

                                    //db.tbEquipmentMoalefeValue.Add(obj22);
                                    //db.SaveChanges();
                                    //shomarande++;





                                }
                                db.tbEquipmentMoalefeValue.AddRange(objectsToAdd);
                                await db.SaveChangesAsync();
                                if (await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.Final_Accept == true && p.FK_Baste == FK_BASTE && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).FirstOrDefaultAsync() != null)
                                {
                                    foreach (var item in await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.tbEquipmentMoalefeValue.Any(m => m.Month == Month && m.Year == Year)).ToListAsync())
                                    {
                                        item.is_submit = 1;
                                        await db.SaveChangesAsync();
                                    }
                                }


                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                    await db.SaveChangesAsync();
                    return Content(" فایل را دوباره بررسی بفرمایید ");
                    return Content(ex.Message);
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return Content("True");









        }




        public ActionResult helppview()
        {
            return View();
        }






        public async Task<ActionResult> GetDataFromExcelMachinsOrTools4(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate,
         int Year, int Month, int Type, int peymanID, int FK_BASTE)
        {

            var pyy = db.tbPeymanContracts.Where(p => p.pec_ID == peymanID && p.Inactive != true && p.Inactive != true).FirstOrDefault();
            var tool = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == Type).FirstOrDefault();
            var pymanname = pyy.pec_Title;
            var tools = tool.Eqpgrp_Name;
            var ty = db.tbmarahelsabt.Where(p => p.FK_Group == Type).Select(s => s.ID).FirstOrDefault();
            var us = db.tbReffrenceSave
                .Where(p => p.FK_PeymanID == peymanID && p.tbPeymanContracts.Inactive != true && p.IsFor == ty)


                 .OrderByDescending(p => p.ID)
                 .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
     .Where(p => p.FK_RRSave == us.ID)
     .OrderByDescending(p => p.ID)
     .FirstOrDefault();

            var rus3 = db.tbReffrenceSaveLevelUser
          .Where(p => p.FK_LevelID == us2.ID)
          .OrderByDescending(p => p.ID)
          .Select(p => p.FK_UserID)
          .FirstOrDefault();
            var User = new tbUsers();
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User != null)
                {

                    userid = User.usr_ID;

                }
            }
            var sabtt = db.tbEquipmentMoalefeValueReffrenceSave
.Where(p => p.FK_Baste == FK_BASTE &&

          p.FK_User == userid &&

          p.tbEquipmentMoalefeValue
              .Any(m => m.Year == Year && m.Month == Month))
.FirstOrDefault();
            var find = await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefaultAsync();
            if (find != null)
            {
                return Content("این بسته در این ماه و سال  تایید نهایی شده است ");
            }
            if (sabtt != null)
            {
                return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
            }

            List<string> lstMoalefe = new List<string>();
            List<int> intt = new List<int>();

            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;
            tbEquipmentSpecificationData obj2 = new tbEquipmentSpecificationData();
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();
            List<tbEquipmentMoalefeValue> list = new List<tbEquipmentMoalefeValue>();

            tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
            objsave.FK_User = userid;
            objsave.DateTime = DateTime.Now;
            objsave.FK_Baste = FK_BASTE;


            //var firstFilter = Filters.First();
            //var fkBASTE = firstFilter.FK_BASTE;

            var rus33 = await db.tbReffrenceSaveLevelUser
.Where(p => p.FK_LevelID == FK_BASTE)
.OrderByDescending(p => p.ID)
.Select(p => p.FK_UserID)
.FirstOrDefaultAsync();
            if (userid == rus3)
            {
                objsave.Final_Sabt = true;
            }

            if (MyExcelStream != null)
            {

                if (MyExcelStream.ContentLength > 0)
                {
                    var segment = MyExcelStream.FileName.Split('.');
                    string file_type = segment[segment.Length - 1];
                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                    MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                    objsave.File_SystemNameForAccept = filename;
                    objsave.FileNameForAccept = MyExcelStream.FileName;
                }
            }



            db.tbEquipmentMoalefeValueReffrenceSave.Add(objsave);
            db.SaveChanges();




            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();


                            if (count == 1)
                            {
                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                            }

                            for (int i = 2; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }



                            var user = await db.tbUsers.Where(p => p.usr_Personal_ID != null).ToListAsync();
                            var pymn = await db.tbPeymanContracts.Where(p => p.Inactive != true).ToListAsync();

                            for (int i = 3; i < count; i++)
                            {
                                tbEquipmentMoalefeValue list2 = new tbEquipmentMoalefeValue();

                                var row = workbook.Worksheets[0].Rows[i];
                                var Name = row.Cells[1];


                                //int savedItemId = db.tbEquipments
                                //                    .OrderByDescending(p => p.ID)
                                //                    .Select(s => s.ID)
                                //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                                int shomarande = 2;
                                List<tbEquipmentMoalefeValue> objectsToAdd = new List<tbEquipmentMoalefeValue>();

                                foreach (var it in lstMoalefe)
                                {
                                    tbEquipmentMoalefeValue obj22 = new tbEquipmentMoalefeValue();
                                    if (int.TryParse(Name.Value.ToString(), out int vahed2))
                                    {
                                        var find2 = pymn.Where(p => p.pec_ProjectCode == vahed2).FirstOrDefault();
                                        var findusr = user.Where(p => p.usr_Personal_ID == vahed2).FirstOrDefault();

                                        if (find2 != null)
                                        {
                                            obj22.Fk_pymn = find2.pec_ID;
                                        }
                                        else
                                        {
                                            obj22.Fk_user = findusr.usr_ID;
                                        }
                                    }
                                    else
                                    {
                                        db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                                        await db.SaveChangesAsync();
                                        return Content(" فایل را دوباره بررسی بفرمایید ");
                                    }

                                    var Value = row.Cells[shomarande];
                                    var value = Value.Value ?? null;
                                    float? countDays;

                                    if (float.TryParse(value, out float parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }
                                    else
                                    {
                                        countDays = null; // Or handle the case where the value cannot be parsed
                                    }

                                    int? FK_Equipment = null;

                                    if (int.TryParse(it, out int FK_Equipment2))
                                    {
                                        FK_Equipment = FK_Equipment2;
                                    }
                                    else
                                    {
                                        FK_Equipment = null; // Or handle the case where the value cannot be parsed
                                    }

                                    obj22.CountDays = countDays;
                                    obj22.FK_Equipment = FK_Equipment;
                                    obj22.FK_tbEquipmentMoalefeValueReffrenceSave = objsave.ID;
                                    obj22.Month = Month;
                                    obj22.Year = Year;

                                    objectsToAdd.Add(obj22);
                                    shomarande++;
                                }

                                // Add all objects in the list to the database using AddRange
                                db.tbEquipmentMoalefeValue.AddRange(objectsToAdd);
                                await db.SaveChangesAsync();


                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    db.tbEquipmentMoalefeValueReffrenceSave.Remove(objsave);
                    await db.SaveChangesAsync();
                    return Content(ex.Message);
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return Content("True");









        }


        public async Task<ActionResult> GetDataFromExcelMachinsOrTools4edit(HttpPostedFileBase MyExcelStream,
 int Year, int Month, int peymanID)
        {
            List<string> lstMoalefe = new List<string>();

            List<int> intt = new List<int>();
            List<int> lstMoalefeint = new List<int>();

            string name = ""; string family = "";
            float? vahed, countvahed, numbervahed, colmablagh, numbermondareg, typetazmin, typemalk, numbertazmin, typecar, codepersenly, soght, egareh;
            tbEquipmentSpecificationData obj2 = new tbEquipmentSpecificationData();
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();
            List<tbEquipmentMoalefeValue> list = new List<tbEquipmentMoalefeValue>();

            //tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
            var exi = await db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == peymanID).FirstOrDefaultAsync();
            var ttt = await db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == exi.ID).ToListAsync();
            //foreach (var item in ttt)
            //{
            //    db.tbEquipmentMoalefeValue.Remove(item);
            //}
            //    db.tbEquipmentMoalefeValue.Remove(item);
            db.tbEquipmentMoalefeValue.RemoveRange(ttt);
            await db.SaveChangesAsync();
            if (MyExcelStream != null)
            {

                if (MyExcelStream.ContentLength > 0)
                {
                    var segment = MyExcelStream.FileName.Split('.');
                    string file_type = segment[segment.Length - 1];
                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                    MyExcelStream.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                    exi.File_SystemNameForAccept = filename;
                    exi.FileNameForAccept = MyExcelStream.FileName;
                    await db.SaveChangesAsync();

                }
            }







            using (ExcelEngine xl = new ExcelEngine())
            {
                try
                {
                    IApplication app = xl.Excel;
                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);
                    var Count_Sheets = workbook.Worksheets.Count;
                    if (Count_Sheets > 0)
                    {

                        var sheet = workbook.Worksheets[0];
                        var Rows = sheet.Rows;
                        //var count = workbook.Sheets[0].Rows.Count();

                        if (Rows.Length >= 1)
                        {
                            var title = workbook.Worksheets[0].Rows[0].Cells;
                            var count = workbook.Worksheets[0].Rows.Count();
                            //var title2 = workbook.Worksheets[0].Rows[0].Cells;


                            if (count == 1)
                            {
                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                            }

                            for (int i = 2; i < title.Length; i++)
                            {
                                lstMoalefe.Add(title[i].Value.ToString());
                            }
                            //for (int i = 3; i < title2.Length; i++)
                            //{
                            //    if (int.TryParse(title2[i].Value.ToString(), out int cellValue))
                            //    {
                            //        lstMoalefeint.Add(cellValue);
                            //    }

                            //}


                            var user = await db.tbUsers.Where(p => p.usr_Personal_ID != null).ToListAsync();
                            var pymn = await db.tbPeymanContracts.Where(p => p.Inactive != true).ToListAsync();

                            for (int i = 3; i < count; i++)
                            {
                                tbEquipmentMoalefeValue list2 = new tbEquipmentMoalefeValue();

                                var row = workbook.Worksheets[0].Rows[i];
                                var Name = row.Cells[1];


                                //int savedItemId = db.tbEquipments
                                //                    .OrderByDescending(p => p.ID)
                                //                    .Select(s => s.ID)
                                //                    .FirstOrDefault(); // فراخوانی متد FirstOrDefault() برای دریافت مقدار

                                // savedItemId حالا مقدار اولین شناسه (ID) است که با استفاده از LINQ دریافت شده است
                                int shomarande = 2;
                                List<tbEquipmentMoalefeValue> objectsToAdd = new List<tbEquipmentMoalefeValue>();

                                foreach (var it in lstMoalefe)
                                {
                                    tbEquipmentMoalefeValue obj22 = new tbEquipmentMoalefeValue();
                                    if (int.TryParse(Name.Value.ToString(), out int vahed2))
                                    {
                                        var find2 = pymn.Where(p => p.pec_ProjectCode == vahed2).FirstOrDefault();
                                        var findusr = user.Where(p => p.usr_Personal_ID == vahed2).FirstOrDefault();

                                        if (find2 != null)
                                        {
                                            obj22.Fk_pymn = find2.pec_ID;
                                        }
                                        else
                                        {
                                            obj22.Fk_user = findusr.usr_ID;
                                        }
                                    }
                                    else
                                    {
                                        db.tbEquipmentMoalefeValueReffrenceSave.Remove(exi);
                                        await db.SaveChangesAsync();
                                        return Content(" فایل را دوباره بررسی بفرمایید ");                                        //return Content(" لطفا ستون نام  سطر " + " ( " + i + " ) " + "را پر کنید ");
                                    }

                                    var Value = row.Cells[shomarande];
                                    var value = Value.Value ?? null;
                                    float? countDays;

                                    if (float.TryParse(value, out float parsedValue))
                                    {
                                        countDays = parsedValue;
                                    }
                                    else
                                    {
                                        countDays = null; // Or handle the case where the value cannot be parsed
                                    }

                                    int? FK_Equipment = null;

                                    if (int.TryParse(it, out int FK_Equipment2))
                                    {
                                        FK_Equipment = FK_Equipment2;
                                    }
                                    else
                                    {
                                        FK_Equipment = null; // Or handle the case where the value cannot be parsed
                                    }

                                    obj22.CountDays = countDays;
                                    obj22.FK_Equipment = FK_Equipment;
                                    obj22.FK_tbEquipmentMoalefeValueReffrenceSave = exi.ID;
                                    obj22.Month = Month;
                                    obj22.Year = Year;

                                    objectsToAdd.Add(obj22);
                                    shomarande++;
                                }

                                // Add all objects in the list to the database using AddRange
                                db.tbEquipmentMoalefeValue.AddRange(objectsToAdd);
                                await db.SaveChangesAsync();


                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    db.tbEquipmentMoalefeValueReffrenceSave.Remove(exi);
                    await db.SaveChangesAsync();
                    return Content(ex.Message);
                }
            }
            //var Model3 = db.tbEquipments.Where(p => p.FK_Bunch == Basteh && p.FK_Peyman == PeymanID).ToList();

            return Content("True");









        }

        public ActionResult GetDataFromExcelMachinsOrTools2(HttpPostedFileBase MyExcelStream, DateTime FromDate, DateTime ToDate,
          int Year, int Month, int Type, int peymanID, int FK_BASTE)
        {
            var pyy = db.tbPeymanContracts.Where(p => p.pec_ID == peymanID && p.Inactive != true).FirstOrDefault();
            var tool = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == Type).FirstOrDefault();
            var pymanname = pyy.pec_Title;
            var tools = tool.Eqpgrp_Name;
            var us = db.tbReffrenceSave
                .Where(p => p.FK_PeymanID == peymanID && p.tbPeymanContracts.Inactive != true && p.IsFor == Type)


                 .OrderByDescending(p => p.ID)
                 .FirstOrDefault();
            var us2 = db.tbReffrenceSaveLevel
     .Where(p => p.FK_RRSave == us.ID)
     .OrderByDescending(p => p.ID)
     .FirstOrDefault();

            var rus3 = db.tbReffrenceSaveLevelUser
          .Where(p => p.FK_LevelID == us2.ID)
          .OrderByDescending(p => p.ID)
          .Select(p => p.FK_UserID)
          .FirstOrDefault();
            var User = new tbUsers();
            int userid = 0;

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User != null)
                {

                    userid = User.usr_ID;

                }
            }
            var sabtt = db.tbEquipmentMoalefeValueReffrenceSave
     .Where(p => p.FK_Baste == FK_BASTE &&

                 p.FK_User == userid &&

                 p.tbEquipmentMoalefeValue
                     .Any(m => m.Year == Year && m.Month == Month))
     .FirstOrDefault();

            using (SaabEntities db = new SaabEntities())
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    using (ExcelEngine xl = new ExcelEngine())
                    {

                        if (Type == 3)
                        {
                            {

                                Virwforexcel equipment_obj = new Virwforexcel();


                                equipment_obj.SabtHeader = new SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment.Header();  // اضافه کردن این خط برای مقداردهی اولیه
                                equipment_obj.informationToolsss = new List<informationTools>();  // اضافه کردن این خط برای مقداردهی اولیه
                                equipment_obj.FK_BASTE = FK_BASTE;

                                equipment_obj.SabtHeader.FromDate = FromDate;
                                equipment_obj.SabtHeader.ToDate = System.Convert.ToDateTime(ToDate);
                                equipment_obj.SabtHeader.MoalfeVal_Year = Year;
                                equipment_obj.SabtHeader.MoalfeVal_Month = Month;
                                equipment_obj.SabtHeader.PeymanName = pymanname;
                                equipment_obj.SabtHeader.Type = tools;
                                var find = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefault();



                                try
                                {
                                    IApplication app = xl.Excel;
                                    var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                                    var Count_Sheets = workbook.Worksheets.Count;
                                    if (Count_Sheets > 1)
                                    {
                                        if (find != null)
                                        {
                                            return Content("این بسته در این ماه و سال  تایید نهایی شده است ");
                                        }
                                        if (sabtt != null)
                                        {
                                            return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
                                        }

                                        var sheet = workbook.Worksheets[0];
                                        var Rows = sheet.Rows;
                                        int Cols = 0;
                                        if (Rows.Length >= 1)
                                        {
                                            Cols = Rows[0].Cells.Length;

                                            //تعداد ستون ها
                                            if (Count_Sheets != 3 || Cols != 7)
                                            {
                                                return Content("اکسل وارد شده با اکسل نمونه  متفاوت است");
                                            }

                                            var title = workbook.Worksheets[0].Rows[0].Cells;
                                            var count = workbook.Worksheets[0].Rows.Count();
                                            if (count == 1)
                                            {
                                                return Content("فایل اکسل فاقد اطلاعات می باشد");
                                            }

                                            Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                            int counter = 0;
                                            foreach (var item in title)
                                            {
                                                dicTitles.Add(item.Value.ToString(), counter);
                                                counter++;
                                            }
                                            tbEquipmentMoalefeValueReffrenceSave obj_EquipmentMoalefeValueReffrenceSav = new tbEquipmentMoalefeValueReffrenceSave();
                                            obj_EquipmentMoalefeValueReffrenceSav.DateTime = DateTime.Now;
                                            obj_EquipmentMoalefeValueReffrenceSav.FK_User = User.usr_ID;


                                            List<tbEquipmentMoalefeValue> myList = new List<tbEquipmentMoalefeValue>();
                                            for (int i = 1; i < count; i++)
                                            {
                                                var infor = new informationTools();
                                                var row = workbook.Worksheets[0].Rows[i];
                                                var cell_code = row.Cells[0];

                                                if (cell_code != null)
                                                {
                                                    if (cell_code.Value != null || cell_code.Value != "")
                                                    {
                                                        infor.ID = System.Convert.ToInt32(cell_code.Value);
                                                    }
                                                    else
                                                    {
                                                        return Content("اکسل وارد شده با اکسل نمونه تفاوت دارد");
                                                    }
                                                }
                                                else
                                                {
                                                    return Content("اکسل وارد شده با اکسل نمونه تفاوت دارد");
                                                }

                                                var cell_countDays = row.Cells[4];
                                                if (cell_countDays != null)
                                                {
                                                    if (!string.IsNullOrEmpty(cell_countDays.Value))
                                                    {
                                                        infor.days_number = System.Convert.ToInt32(cell_countDays.Value);
                                                    }
                                                    else
                                                    {
                                                        return Content(" لطفا ستون تعداد روز سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                    }
                                                }
                                                var Other_tool_information = row.Cells[3];
                                                if (Other_tool_information != null)
                                                {
                                                    if (!string.IsNullOrEmpty(Other_tool_information.Value))
                                                    {
                                                        infor.Other_tool_information = System.Convert.ToString(Other_tool_information.Value);
                                                    }
                                                    else
                                                    {
                                                        return Content(" لطفا ستون تعداد روز سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                    }
                                                }


                                                var Condition = row.Cells[5];

                                                if (Condition != null)
                                                {
                                                    if (!string.IsNullOrEmpty(Condition.Value))
                                                    {
                                                        infor.Condition = System.Convert.ToInt32(Condition.Value);
                                                    }
                                                    else
                                                    {
                                                        return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                    }
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                                var cell_status = row.Cells[2];

                                                if (cell_status != null)
                                                {
                                                    if (!string.IsNullOrEmpty(cell_status.Value))
                                                    {
                                                        infor.name_tool = System.Convert.ToString(cell_status.Value);
                                                    }
                                                    else
                                                    {
                                                        return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                    }
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                                var cell_description = row.Cells[6];
                                                if (cell_description != null)
                                                {
                                                    if (!string.IsNullOrEmpty(cell_description.Value))
                                                    {

                                                        infor.Description = cell_description.Value;

                                                    }
                                                }


                                                equipment_obj.informationToolsss.Add(infor);


                                            }



                                            //db.tbEquipmentMoalefeValueReffrenceSave.Add(obj_EquipmentMoalefeValueReffrenceSav);
                                            //db.SaveChanges();
                                            //transaction.Commit();
                                        }
                                        else
                                        {
                                            return Content("این شیت فاقد سطر می باشد");
                                        }

                                        return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/GetDataFromExcelMachinsOrTools2.cshtml", equipment_obj);
                                    }
                                    else
                                    {
                                        return Content("هیچ شیتی در این اکسل وجود ندارد");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    return Content("falsde");
                                }

                            }
                        }
                        else if (Type == 4)
                        {
                            Viewexceleabzar equipment_obj = new Viewexceleabzar();
                            equipment_obj.SabtHeader = new SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment.Header();  // اضافه کردن این خط برای مقداردهی اولیه
                            equipment_obj.informationtool = new List<information>();  // اضافه کردن این خط برای مقداردهی اولیه
                            equipment_obj.FK_BASTE = FK_BASTE;
                            equipment_obj.SabtHeader.FromDate = FromDate;
                            equipment_obj.SabtHeader.ToDate = System.Convert.ToDateTime(ToDate);
                            equipment_obj.SabtHeader.MoalfeVal_Year = Year;
                            equipment_obj.SabtHeader.MoalfeVal_Month = Month;
                            equipment_obj.SabtHeader.PeymanName = pymanname;
                            equipment_obj.SabtHeader.Type = tools;
                            var find = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_Baste == FK_BASTE && p.Final_Accept == true && p.tbEquipmentMoalefeValue.Any(m => m.Year == Year && m.Month == Month)).FirstOrDefault();

                            try
                            {
                                IApplication app = xl.Excel;
                                var workbook = app.Workbooks.Open(MyExcelStream.InputStream);

                                var Count_Sheets = workbook.Worksheets.Count;
                                if (Count_Sheets > 1)
                                {
                                    if (find != null)
                                    {
                                        return Content("این بسته در این ماه و سال  تایید نهایی شده است ");
                                    }
                                    if (sabtt != null)
                                    {
                                        return Content("شما قبلا این بسته را در این سال و ماه ثبت کرده اید  ");
                                    }

                                    var sheet = workbook.Worksheets[0];
                                    var Rows = sheet.Rows;
                                    int Cols = 0;
                                    if (Rows.Length >= 1)
                                    {
                                        Cols = Rows[0].Cells.Length;

                                        //تعداد ستون ها
                                        if (Count_Sheets != 3 || Cols != 9)
                                        {
                                            return Content("اکسل وارد شده با اکسل نمونه  متفاوت است");
                                        }

                                        var title = workbook.Worksheets[0].Rows[0].Cells;
                                        var count = workbook.Worksheets[0].Rows.Count();
                                        if (count == 1)
                                        {
                                            return Content("فایل اکسل فاقد اطلاعات می باشد");
                                        }

                                        Dictionary<string, int> dicTitles = new Dictionary<string, int>();
                                        int counter = 0;
                                        foreach (var item in title)
                                        {
                                            dicTitles.Add(item.Value.ToString(), counter);
                                            counter++;
                                        }
                                        tbEquipmentMoalefeValueReffrenceSave obj_EquipmentMoalefeValueReffrenceSav = new tbEquipmentMoalefeValueReffrenceSave();
                                        obj_EquipmentMoalefeValueReffrenceSav.DateTime = DateTime.Now;
                                        obj_EquipmentMoalefeValueReffrenceSav.FK_User = User.usr_ID;

                                        List<tbEquipmentMoalefeValue> myList = new List<tbEquipmentMoalefeValue>();
                                        for (int i = 1; i < count; i++)
                                        {


                                            var infor = new information();


                                            var row = workbook.Worksheets[0].Rows[i];
                                            var cell_code = row.Cells[0];

                                            if (cell_code != null)
                                            {
                                                if (cell_code.Value != null || cell_code.Value != "")
                                                {
                                                    infor.ID = System.Convert.ToInt32(cell_code.Value);
                                                }
                                                else
                                                {
                                                    return Content("اکسل وارد شده با اکسل نمونه تفاوت دارد");
                                                }
                                            }
                                            else
                                            {
                                                return Content("اکسل وارد شده با اکسل نمونه تفاوت دارد");
                                            }
                                            var cell_countDays = row.Cells[6];
                                            if (cell_countDays != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_countDays.Value))
                                                {
                                                    infor.days_number = System.Convert.ToInt32(cell_countDays.Value);
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون تعداد روز سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                            }

                                            var cell_status = row.Cells[7];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_status.Value))
                                                {
                                                    infor.Condition = System.Convert.ToInt32(cell_status.Value);
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                            }
                                            else
                                            {
                                                return Content(" لطفا ستون وضعیت سطر " + " ( " + i + " ) " + "را پر کنید ");
                                            }

                                            var cell_Accept = row.Cells[5];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_Accept.Value))
                                                {
                                                    infor.Approved_number = System.Convert.ToInt32(cell_Accept.Value);
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون تعداد مورد تایید سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                            }
                                            var name_tool = row.Cells[2];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(name_tool.Value))
                                                {
                                                    infor.name_tool = name_tool.Value;
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون تعداد مورد تایید سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                            }
                                            else
                                            {
                                                return Content(" لطفا ستون تعداد موردتایید سطر " + " ( " + i + " ) " + "را پر کنید ");
                                            }

                                            var contractNumber = row.Cells[4];

                                            if (cell_status != null)
                                            {
                                                if (!string.IsNullOrEmpty(contractNumber.Value))
                                                {
                                                    infor.number_recorded = System.Convert.ToInt32(contractNumber.Value);
                                                }
                                                else
                                                {
                                                    return Content(" لطفا ستون تعداد ثبت شده درقرارداد سطر " + " ( " + i + " ) " + "را پر کنید ");
                                                }
                                            }
                                            else
                                            {
                                                return Content(" لطفا ستون تعداد ثبت شده در قراداد سطر " + " ( " + i + " ) " + "را پر کنید ");
                                            }

                                            var cell_description = row.Cells[8];
                                            if (cell_description != null)
                                            {
                                                if (!string.IsNullOrEmpty(cell_description.Value))
                                                {

                                                    infor.Description = cell_description.Value;

                                                }
                                            }




                                            equipment_obj.informationtool.Add(infor);

                                        }

                                        //db.tbEquipmentMoalefeValue.AddRange(myList);




                                        var userid2 = User.usr_ID;
                                        if (userid2 == rus3)
                                        {

                                            var matchedRows = db.tbMoalefeValuePishkhan
                                                .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان")
                                                .ToList();

                                            foreach (var row in matchedRows)
                                            {
                                                row.mlfval_Value = "1";

                                            }

                                            db.SaveChanges();

                                        }
                                        transaction.Commit();
                                    }

                                    else
                                    {
                                        return Content("این شیت فاقد سطر می باشد");
                                    }

                                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/View_Excel_fortools.cshtml", equipment_obj);

                                }
                                else
                                {
                                    return Content("هیچ شیتی در این اکسل وجود ندارد");
                                }

                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                return Content("false");
                            }
                        }

                        else
                        {
                            return Content("ثبت برای این دسته از کارکرد ها تعریف نشده است");
                        }
                    }
                }
            }
        }
        public ActionResult sendtools(List<Viewexceleabzar> Filters)
        {
            var User = new tbUsers();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User != null)
                {

                    userid = User.usr_ID;










                    tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
                    objsave.FK_User = userid;
                    objsave.DateTime = DateTime.Now;
                    objsave.FK_Baste = Filters.First().FK_BASTE;






                    var firstFilter = Filters.First();
                    var fkBASTE = firstFilter.FK_BASTE;

                    var rus3 = db.tbReffrenceSaveLevelUser
                        .Where(p => p.FK_LevelID == fkBASTE)
                        .OrderByDescending(p => p.ID)
                        .Select(p => p.FK_UserID)
                        .FirstOrDefault();
                    if (userid == rus3)
                    {
                        objsave.Final_Sabt = true;
                        var matchedRows = db.tbMoalefeValuePishkhan
                            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ابزار پیمان")
                            .ToList();

                        foreach (var row in matchedRows)
                        {
                            row.mlfval_Value = "1";

                        }

                        db.SaveChanges();

                    }




                    db.tbEquipmentMoalefeValueReffrenceSave.Add(objsave);
                    db.SaveChanges();
                    foreach (var filter in Filters)
                    {
                        List<information> infore = filter.informationtool;
                        Header SabtHeader = filter.SabtHeader;

                        foreach (var userMoalefe in infore)
                        {
                            tbEquipmentMoalefeValue obj = new tbEquipmentMoalefeValue(); // Create a new instance for each iteration
                            obj.Year = SabtHeader.MoalfeVal_Year;
                            obj.Month = SabtHeader.MoalfeVal_Month;
                            obj.FromDate = SabtHeader.FromDate;
                            obj.ToDate = SabtHeader.ToDate;
                            obj.NumberSavedInContracts = userMoalefe.number_recorded;
                            obj.NumberSavedInContractsThatAccepted = userMoalefe.Approved_number;

                            obj.FK_Equipment = userMoalefe.ID;
                            obj.Description = userMoalefe.Description;
                            obj.Status = userMoalefe.Condition;
                            obj.CountDays = userMoalefe.days_number;
                            obj.FK_tbEquipmentMoalefeValueReffrenceSave = objsave.ID;

                            db.tbEquipmentMoalefeValue.Add(obj);
                        }
                    }

                    db.SaveChanges();
                    return Content("true");

                }
            }

            return Content("true");
        }
        public ActionResult sendtoo(List<Virwforexcel> Filters)
        {
            var User = new tbUsers();
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                int userid = 0;
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                }
                if (User != null)
                {

                    userid = User.usr_ID;



                    tbEquipmentMoalefeValueReffrenceSave objsave = new tbEquipmentMoalefeValueReffrenceSave();
                    objsave.FK_User = userid;
                    objsave.DateTime = DateTime.Now;
                    objsave.FK_Baste = Filters.First().FK_BASTE;


                    var firstFilter = Filters.First();
                    var fkBASTE = firstFilter.FK_BASTE;

                    var rus3 = db.tbReffrenceSaveLevelUser
.Where(p => p.FK_LevelID == fkBASTE)
.OrderByDescending(p => p.ID)
.Select(p => p.FK_UserID)
.FirstOrDefault();
                    if (userid == rus3)
                    {
                        objsave.Final_Sabt = true;

                        var matchedRows = db.tbMoalefeValuePishkhan
                            .Where(p => p.tbContractMoalefeDastmozdi.md_Title == "پیشخوان-پیمان-مستندات وضعیت-ثبت ماشین آلات پیمان")
                            .ToList();

                        foreach (var row in matchedRows)
                        {
                            row.mlfval_Value = "1";

                        }

                        db.SaveChanges();

                    }

                    db.tbEquipmentMoalefeValueReffrenceSave.Add(objsave);
                    db.SaveChanges();
                    foreach (var filter in Filters)
                    {
                        List<informationTools> infore = filter.informationToolsss;
                        Header SabtHeader = filter.SabtHeader;

                        foreach (var userMoalefe in infore)
                        {
                            tbEquipmentMoalefeValue obj = new tbEquipmentMoalefeValue(); // Create a new instance for each iteration
                            obj.Year = SabtHeader.MoalfeVal_Year;
                            obj.Month = SabtHeader.MoalfeVal_Month;
                            obj.FromDate = SabtHeader.FromDate;
                            obj.ToDate = SabtHeader.ToDate;
                            obj.FK_Equipment = userMoalefe.ID;
                            obj.Description = userMoalefe.Description;
                            obj.Status = userMoalefe.Condition;
                            obj.CountDays = userMoalefe.days_number;
                            obj.FK_tbEquipmentMoalefeValueReffrenceSave = objsave.ID;

                            db.tbEquipmentMoalefeValue.Add(obj);
                        }
                    }

                    db.SaveChanges();
                    return Content("true");

                }
            }

            return Content("true");
        }


        public ActionResult pavastrools(IEnumerable<HttpPostedFileBase> files = null, int id = 0)
        {
            tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();

            // Retrieve the record by ID
            var record = db.tbEquipmentMoalefeValueReffrenceSave.Find(id);

            if (record != null)
            {
                if (files != null)
                {
                    foreach (var file in files)
                    {
                        if (file.ContentLength > 0)
                        {
                            var segment = file.FileName.Split('.');
                            string file_type = segment[segment.Length - 1];
                            var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                            file.SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));

                            // Save the values to the existing record
                            record.FileName = filename;
                            record.File_SystemName = file.FileName;
                        }
                    }

                    // Save changes to the database
                    db.SaveChanges();
                    return Content("True");

                }
            }
            return View();
        }



        public ActionResult _ListEquipmentMoalefeValue()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var a = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_User == User.usr_ID).ToList();
                    foreach (var item in a)
                    {
                        item.tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item.ID).ToList();
                        foreach (var item2 in item.tbEquipmentMoalefeValue.ToList())
                        {
                            item2.tbEquipments = db.tbEquipments.Find(item2.FK_Equipment);
                            item2.tbEquipments.tbPeymanContracts = db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                            item2.tbEquipments.tbEquipmentBunch = db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                            item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                        }
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValue.cshtml", a);

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValue.cshtml", new List<tbEquipmentMoalefeValue>());
        }




        public ActionResult _ListEquipmentMoalefeValuefor()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();

                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var today = DateTime.Now.GetShamsiDayOfMonth();
                    var levels = db.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth >= today && today <= (p.StartDayInMonth + p.Duration)).ToList();

                    foreach (var lvl in levels)
                    {
                        var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;

                        if (baste_duration != 1)
                        {

                        }
                    }
                    var a = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.FK_User == User.usr_ID).ToList();

                    foreach (var item in a)
                    {
                        item.tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item.ID).ToList();
                        foreach (var item2 in item.tbEquipmentMoalefeValue.ToList())
                        {
                            item2.tbEquipments = db.tbEquipments.Find(item2.FK_Equipment);
                            item2.tbEquipments.tbPeymanContracts = db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                            item2.tbEquipments.tbEquipmentBunch = db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                            item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                        }
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValue.cshtml", a);





                }

            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValue.cshtml", new List<tbEquipmentMoalefeValue>());
        }









        public ActionResult ShowUsersFuctionforPy()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    if (User != null)
                    {
                        SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository rep_accept = new SaabWebProject.Models.Repositories.UserFunctions.ReffrenceAcceptRepository();
                        var result = rep_accept.Listt2(User.usr_ID);
                        List<tbReffrenceAcceptLevel> baste_id2 = new List<tbReffrenceAcceptLevel>();
                        List<tbReffrenceAcceptLevelUsers> baste_id3 = new List<tbReffrenceAcceptLevelUsers>();
                        List<tbEquipmentMoalefeValueReffrenceSave> list = new List<tbEquipmentMoalefeValueReffrenceSave>();
                        List<tbEquipmentMoalefeValueReffrenceSave> list2 = new List<tbEquipmentMoalefeValueReffrenceSave>();
                        List<tbEquipmentMoalefeValueReffrenceSave> list3 = new List<tbEquipmentMoalefeValueReffrenceSave>();

                        foreach (var item in result)
                        {
                            var acc = db.tbReffrenceAcceptLevel.Where(p => p.FK_ReffrenceAccept == item.ID).FirstOrDefault();
                            baste_id2.Add(acc);
                        }
                        foreach (var item in baste_id2)
                        {
                            var acc = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_ReffrenceAcceptLevel == item.ID && p.FK_User == User.usr_ID).FirstOrDefault();
                            baste_id3.Add(acc);
                        }
                        foreach (var item in baste_id3)
                        {
                            var b = db.tbEquipments.Where(p => p.FK_Peyman == item.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID && p.tbPeymanContracts.Inactive != true).ToList();
                            //list.Add(b.FirstOrDefault().tbEquipmentMoalefeValue.FirstOrDefault().tbEquipmentMoalefeValueReffrenceSave);
                            foreach (var item2 in b)
                            {
                                if (item2.tbEquipmentMoalefeValue.Count() != null)
                                {
                                    foreach (var item3 in item2.tbEquipmentMoalefeValue.ToList())
                                    {
                                        var exist = item3.tbEquipmentMoalefeValueReffrenceSave;
                                        if (exist != null && exist.Final_Accept == true)
                                        {
                                            list2.Add(exist);
                                        }
                                        else
                                        {
                                            list3.Add(exist);

                                        }
                                    }
                                    if (list2.Count == 0)
                                    {

                                        foreach (var item3 in list3)
                                        {
                                            list.Add(item3);

                                        }
                                    }




                                }
                            }
                            list = list.Distinct().ToList();
                            foreach (var item3 in list)
                            {
                                item3.tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item3.ID).ToList();
                                foreach (var item2 in item3.tbEquipmentMoalefeValue.ToList())
                                {
                                    item2.tbEquipments = db.tbEquipments.Find(item2.FK_Equipment);
                                    item2.tbEquipments.tbPeymanContracts = db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                                    item2.tbEquipments.tbEquipmentBunch = db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                                    item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                                }
                            }
                        }
                        return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", list.Distinct().ToList());
                    }

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());

        }









        public ActionResult _ListForAcceptTools2()
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    List<tbEquipmentMoalefeValueReffrenceSave> list = new List<tbEquipmentMoalefeValueReffrenceSave>();
                    var a = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == User.usr_ID).ToList();
                    foreach (var item in a)
                    {
                        var b = db.tbEquipments.Where(p => p.FK_Peyman == item.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID && p.tbPeymanContracts.Inactive != true).ToList();
                        //list.Add(b.FirstOrDefault().tbEquipmentMoalefeValue.FirstOrDefault().tbEquipmentMoalefeValueReffrenceSave);

                        foreach (var item2 in b)
                        {
                            if (item2.tbEquipmentMoalefeValue.Count() != null)
                            {
                                foreach (var item3 in item2.tbEquipmentMoalefeValue.ToList())
                                {
                                    list.Add(item3.tbEquipmentMoalefeValueReffrenceSave);
                                }
                            }
                        }
                        list = list.Distinct().ToList();
                        foreach (var item3 in list)
                        {
                            item3.tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item3.ID).ToList();
                            foreach (var item2 in item3.tbEquipmentMoalefeValue.ToList())
                            {
                                item2.tbEquipments = db.tbEquipments.Find(item2.FK_Equipment);
                                item2.tbEquipments.tbPeymanContracts = db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                                item2.tbEquipments.tbEquipmentBunch = db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                                item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                            }
                        }
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", list.Distinct().ToList());

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        }












        //public async Task<ActionResult> _ListForAcceptTools3()
        //{
        //    var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
        //    if (cookie_user != null)
        //    {
        //        var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);

        //        using (var db = new SaabEntities())
        //        {
        //            var User = await db.tbUsers.FirstOrDefaultAsync(p => p.usr_NationalCode.ToString() == nationalcode);
        //            if (User != null)
        //            {
        //                var today = DateTime.Now.GetShamsiDayOfMonth();

        //                var list = await db.tbReffrenceAcceptLevelUsers
        //                    .Where(p => p.FK_User == User.usr_ID &&
        //                                p.tbReffrenceAcceptLevel.StartDayInMonth <= today &&
        //                                today <= (p.tbReffrenceAcceptLevel.StartDayInMonth + p.tbReffrenceAcceptLevel.Duration) &&
        //                                p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor != 1 &&
        //                                p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
        //                    .SelectMany(p => db.tbEquipments.Where(e => e.FK_Peyman == p.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID)
        //                                                      .SelectMany(e => e.tbEquipmentMoalefeValue.Where(ev => ev.is_submit != 1)
        //                                                                                               .Select(ev => ev.tbEquipmentMoalefeValueReffrenceSave)))
        //                    .Distinct()
        //                    .ToListAsync();

        //                foreach (var item in list)
        //                {
        //                    item.tbEquipmentMoalefeValue = await db.tbEquipmentMoalefeValue
        //                        .Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item.ID)
        //                        .ToListAsync();

        //                    foreach (var item2 in item.tbEquipmentMoalefeValue.ToList())
        //                    {
        //                        item2.tbEquipments = await db.tbEquipments.FindAsync(item2.FK_Equipment);
        //                        item2.tbEquipments.tbPeymanContracts = await db.tbPeymanContracts.FindAsync(item2.tbEquipments.FK_Peyman);
        //                        item2.tbEquipments.tbEquipmentBunch = await db.tbEquipmentBunch.FindAsync(item2.tbEquipments.FK_Bunch);
        //                        item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = await db.tbEquipmentGroup.FindAsync(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
        //                    }
        //                }

        //                return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", list);
        //            }
        //        }
        //    }

        //    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        //}



        public ActionResult viewsayet2(int IsFor = 0)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                //var tbEquipments =  db.tbEquipments.ToList();
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var today = DateTime.Now.GetShamsiDayOfMonth();

                    List<tbEquipmentMoalefeValueReffrenceSave> list = new List<tbEquipmentMoalefeValueReffrenceSave>();
                    var a = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == User.usr_ID && p.tbReffrenceAcceptLevel.StartDayInMonth <= today && today <= (p.tbReffrenceAcceptLevel.StartDayInMonth + p.tbReffrenceAcceptLevel.Duration) && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == IsFor && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true).ToList();

                    foreach (var item in a)
                    {

                        var b = db.tbEquipments.Where(p => p.FK_Peyman == item.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID && p.tbPeymanContracts.Inactive != true).ToList();
                        //list.Add(b.FirstOrDefault().tbEquipmentMoalefeValue.FirstOrDefault().tbEquipmentMoalefeValueReffrenceSave);

                        foreach (var item2 in b)
                        {
                            if (item2.tbEquipmentMoalefeValue.Count() != 0)
                            {
                                foreach (var item3 in item2.tbEquipmentMoalefeValue.ToList())
                                {
                                    if (item3.tbEquipmentMoalefeValueReffrenceSave.is_submit != 1)
                                    {
                                        list.Add(item3.tbEquipmentMoalefeValueReffrenceSave);

                                    }
                                }
                            }
                        }
                        list = list.Distinct().ToList();
                        //foreach (var item3 in list)
                        //{
                        //    item3.tbEquipmentMoalefeValue =  db.tbEquipmentMoalefeValue.Where(p =>
                        //    p.FK_tbEquipmentMoalefeValueReffrenceSave == item3.ID).ToList();
                        //    foreach (var item2 in item3.tbEquipmentMoalefeValue.ToList())
                        //    {
                        //        item2.tbEquipments =  db.tbEquipments.Find(item2.FK_Equipment);
                        //        item2.tbEquipments.tbPeymanContracts =  db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                        //        item2.tbEquipments.tbEquipmentBunch =  db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                        //        item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup =  db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                        //    }
                        //}
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewsayet.cshtml", list.Distinct().ToList());

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewsayet.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        }


        public ActionResult viewnew()
        {
            var f = db.tbmarahelsabt.Where(p => p.FK_Group != null).ToList();
            List<tbEquipmentGroup> Model = new List<tbEquipmentGroup>();
            foreach (var it in f)
            {
                var t = db.tbEquipmentGroup.Where(p => p.Eqpgrp_ID == it.FK_Group).FirstOrDefault();
                if (t != null)
                {
                    Model.Add(t);
                }
            }
            if (f != null)
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewnew.cshtml", f);

            }
            else
            {
                return View("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewnew.cshtml", new tbmarahelsabt());

            }
        }


        public ActionResult viewsayet(int IsFor = 0)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                //var tbEquipments =  db.tbEquipments.ToList();
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var tbref = db.tbEquipmentMoalefeValueReffrenceSave.ToList();
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    var today = DateTime.Now.GetShamsiDayOfMonth();

                    List<tbEquipmentMoalefeValueReffrenceSave> list = new List<tbEquipmentMoalefeValueReffrenceSave>();
                    var a = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == User.usr_ID && p.tbReffrenceAcceptLevel.StartDayInMonth <= today && today <= (p.tbReffrenceAcceptLevel.StartDayInMonth + p.tbReffrenceAcceptLevel.Duration) && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == IsFor && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true).ToList();
                    //var aa=a.Where(=)
                    foreach (var item in a)
                    {
                        var find = db.tbmarahelsabt.Where(p => p.ID == IsFor).FirstOrDefault();
                        var b = db.tbEquipments.Where(p => p.FK_Peyman == item.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID && p.tbPeymanContracts.Inactive != true && p.tbEquipmentBunch.FK_EqpgrpID == find.FK_Group).ToList();
                        //list.Add(b.FirstOrDefault().tbEquipmentMoalefeValue.FirstOrDefault().tbEquipmentMoalefeValueReffrenceSave);

                        foreach (var item2 in b)
                        {
                            if (item2.tbEquipmentMoalefeValue.Count() != 0)
                            {
                                foreach (var item3 in item2.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave != null).GroupBy(p => p.FK_tbEquipmentMoalefeValueReffrenceSave).ToList())
                                {
                                    var ty = tbref.Where(p => p.ID == item3.Key).FirstOrDefault();
                                    if (ty.is_submit != 1)
                                    {
                                        //var y = tbref.Where(p => p.ID == item3.Key).FirstOrDefault();
                                        list.Add(ty);

                                    }
                                }
                            }
                        }
                        list = list.Distinct().ToList();
                        //foreach (var item3 in list)
                        //{
                        //    item3.tbEquipmentMoalefeValue =  db.tbEquipmentMoalefeValue.Where(p =>
                        //    p.FK_tbEquipmentMoalefeValueReffrenceSave == item3.ID).ToList();
                        //    foreach (var item2 in item3.tbEquipmentMoalefeValue.ToList())
                        //    {
                        //        item2.tbEquipments =  db.tbEquipments.Find(item2.FK_Equipment);
                        //        item2.tbEquipments.tbPeymanContracts =  db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                        //        item2.tbEquipments.tbEquipmentBunch =  db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                        //        item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup =  db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                        //    }
                        //}
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewsayet.cshtml", list.Distinct().ToList());

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/viewsayet.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        }





        public async Task<ActionResult> _ListForAcceptTools()
        {

            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListForAcceptTools.cshtml");
        }





        public ActionResult _ListEquipmenForAcceper(int Month = 0, int Year = 0, int peyman_ID = 0, int type = 0)

        {
            var equipments = db.tbEquipments.Where(p => p.FK_Peyman == peyman_ID && p.tbEquipmentBunch.FK_EqpgrpID == type && p.tbPeymanContracts.Inactive != true).ToList();



            if (equipments.Count() > 0)
            {
                var firstEquipmentId = equipments.First().ID;
                var exist = db.tbEquipmentMoalefeValue.Where(p => p.FK_Equipment == firstEquipmentId && p.Month == Month && p.Year == Year).Select(p => p.FK_tbEquipmentMoalefeValueReffrenceSave).ToList();            // or equipments[0]
                List<tbEquipmentMoalefeValueReffrenceSave> exist_use = new List<tbEquipmentMoalefeValueReffrenceSave>();  // Replace YourType with the actual type of tbEquipmentMoalefeValueReffrenceSave

                foreach (var itme3 in exist)
                {
                    var itemsFromDatabase = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == itme3).ToList();
                    exist_use.AddRange(itemsFromDatabase);

                }


























                return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmenForAcceper.cshtml", exist_use);

            }

            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmenForAcceper.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        }









        public ActionResult _ListEquipmentMoalefeValueForAcceper(int IsFor = 0, int year_pY1 = 0, int month_pY1 = 0, int pymnID = 0)
        {
            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var tbref = db.tbEquipmentMoalefeValueReffrenceSave.ToList();

                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();
                    List<tbEquipmentMoalefeValueReffrenceSave> list = new List<tbEquipmentMoalefeValueReffrenceSave>();

                    var today = DateTime.Now.GetShamsiDayOfMonth();

                    var a = db.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == User.usr_ID && p.tbReffrenceAcceptLevel.StartDayInMonth <= today && today <= (p.tbReffrenceAcceptLevel.StartDayInMonth + p.tbReffrenceAcceptLevel.Duration) && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == IsFor && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true && p.tbReffrenceAcceptLevel.tbReffrenceAccept.tbReffrenceSaveLevel.ID == pymnID).ToList();
                    foreach (var item in a)
                    {
                        var b = db.tbEquipments.Where(p => p.FK_Peyman == item.tbReffrenceAcceptLevel.tbReffrenceAccept.FK_PeymanID && p.tbPeymanContracts.Inactive != true).ToList();
                        //list.Add(b.FirstOrDefault().tbEquipmentMoalefeValue.FirstOrDefault().tbEquipmentMoalefeValueReffrenceSave);

                        foreach (var item2 in b)
                        {
                            if (item2.tbEquipmentMoalefeValue.Where(p => p.Month == month_pY1 && p.Year == year_pY1 && p.tbEquipmentMoalefeValueReffrenceSave.FK_Baste == pymnID).Count() != 0)
                            {
                                foreach (var item3 in item2.tbEquipmentMoalefeValue.GroupBy(p => p.FK_tbEquipmentMoalefeValueReffrenceSave).ToList())
                                {
                                    var ty = tbref.Where(p => p.ID == item3.Key).FirstOrDefault();
                                    if (ty.is_submit != 1)
                                    {
                                        //var y = tbref.Where(p => p.ID == item3.Key).FirstOrDefault();
                                        list.Add(ty);

                                    }
                                    else if (ty.Final_Accept == true)
                                    {
                                        list.Add(ty);


                                    }
                                }
                                //foreach (var item3 in item2.tbEquipmentMoalefeValue.ToList())
                                //{
                                //    if (item3.tbEquipmentMoalefeValueReffrenceSave.is_submit != 1)
                                //    {
                                //        list.Add(item3.tbEquipmentMoalefeValueReffrenceSave);

                                //    }
                                //}
                            }
                        }
                        list = list.Distinct().ToList();
                        //foreach (var item3 in list)
                        //{
                        //    item3.tbEquipmentMoalefeValue = db.tbEquipmentMoalefeValue.Where(p => p.FK_tbEquipmentMoalefeValueReffrenceSave == item3.ID).ToList();
                        //    foreach (var item2 in item3.tbEquipmentMoalefeValue.ToList())
                        //    {
                        //        item2.tbEquipments = db.tbEquipments.Find(item2.FK_Equipment);
                        //        item2.tbEquipments.tbPeymanContracts = db.tbPeymanContracts.Find(item2.tbEquipments.FK_Peyman);
                        //        item2.tbEquipments.tbEquipmentBunch = db.tbEquipmentBunch.Find(item2.tbEquipments.FK_Bunch);
                        //        item2.tbEquipments.tbEquipmentBunch.tbEquipmentGroup = db.tbEquipmentGroup.Find(item2.tbEquipments.tbEquipmentBunch.FK_EqpgrpID);
                        //    }
                        //}
                    }
                    return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValueForAcceper.cshtml", list.Distinct().ToList());

                }
            }
            return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListEquipmentMoalefeValueForAcceper.cshtml", new List<tbEquipmentMoalefeValueReffrenceSave>());
        }
        public bool AcceptEquipmentMoalefeValue(int IDd = 0, string Description = "", IEnumerable<HttpPostedFileBase> files = null, FormCollection data = null)
        {

            var cookie_user = Request.Cookies[Utility.Base64.Base64Encode("CodeMeli")];
            if (cookie_user != null)
            {
                var nationalcode = Utility.Base64.Base64Decode(cookie_user.Value);
                using (SaabEntities db = new SaabEntities())
                {
                    var User = db.tbUsers.Where(p => p.usr_NationalCode.ToString() == nationalcode).FirstOrDefault();

                    tbEquipmentMoalefeValueReffrenceSave obj = new tbEquipmentMoalefeValueReffrenceSave();
                    obj.DateTime = DateTime.Now;
                    obj.FK_User = User.usr_ID;
                    obj.Description = Description;
                    obj.Parent_ID = IDd;

                    var segment = files.First().FileName.Split('.');
                    string file_type = segment[segment.Length - 1];
                    var filename = (DateTime.Now.ToString("yyyy_MM_dd_mm_ss") + '.' + file_type).ToString();
                    files.First().SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/Areas/Contracts/Contents/DefFacilitiesEquipment/UserUpload/" + filename));
                    obj.File_SystemName = filename;
                    obj.FileName = files.First().FileName;
                    db.tbEquipmentMoalefeValueReffrenceSave.Add(obj);

                    return System.Convert.ToBoolean(db.SaveChanges());
                }
            }
            else
            {
                return false;
            }
        }
        public ActionResult _ListAccepter(int id = 0)
        {
            using (var db = new SaabEntities())
            {
                var a = db.tbEquipmentMoalefeValueReffrenceSave.Where(p => p.ID == id || p.Parent_ID == id).ToList();
                foreach (var item in a)
                {
                    item.tbUsers = db.tbUsers.Find(item.FK_User);
                }
                return PartialView("~/Areas/Contracts/Views/DefFacilitiesEquipment/_ListAccepter.cshtml", a);
            }

        }
        #endregion

    }

}