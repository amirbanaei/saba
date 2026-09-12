using DocumentFormat.OpenXml.Bibliography;
using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
using static SaabWebProject.Areas.Setting.Controllers.OrdersController;
using static SaabWebProject.Models.ViewModels.Contracts.Function.FunctionModel;

namespace SaabWebProject.Models.ViewModels.Contracts.Function
{
    public class FunctionModel
    {
      
        public List<ValuunicPriorityResultrej> PriorityResultrej { get; set; }
        public long countgadval1 { get; set; }

        public long countgadval2 { get; set; }

        public long countgadval3 { get; set; }
        public long countgadval4 { get; set; }
        public long countgadval6 { get; set; }

        public long countgadval7 { get; set; }

        public long countgadval8 { get; set; }
        public long countgadval9 { get; set; }
        public long countgadval10 { get; set; }
        public long countgadval11{ get; set; }
        public long countgadval12 { get; set; }
        public long countgadval13 { get; set; }

        public decimal percent2 { get; set; }
        public decimal percent3 { get; set; }
        public decimal percent4 { get; set; }
        public decimal percent5 { get; set; }

        public decimal percent6 { get; set; }
        public decimal percent7 { get; set; }
        public decimal percent8 { get; set; }
        public decimal percent9 { get; set; }
        public decimal percent10 { get; set; }
        public decimal percent11 { get; set; }
        public decimal percent12 { get; set; }
        public decimal percent13 { get; set; }
        public decimal percent14 { get; set; }
        public decimal percent15 { get; set; }
        public decimal percent16 { get; set; }
        public decimal percent17 { get; set; }


        public decimal percent2c { get; set; }
        public decimal percent3c { get; set; }
        public decimal percent4c { get; set; }
        public decimal percent5c { get; set; }

        public decimal percent6c { get; set; }
        public decimal percent7c { get; set; }
        public decimal percent8c { get; set; }
        public decimal percent9c { get; set; }
        public decimal percent10c { get; set; }
        public decimal percent11c { get; set; }
        public decimal percent12c { get; set; }
        public decimal percent13c { get; set; }
        public decimal percent14c { get; set; }
        public decimal percent15c { get; set; }
        public decimal percent16c { get; set; }
        public decimal percent17c { get; set; }

        public long countgadval5 { get; set; }
        public decimal mablaghtable1 { get; set; }
        public decimal mablaghtable2 { get; set; }
        public decimal mablaghtable3 { get; set; }
        public decimal mablaghtable4 { get; set; }
        public decimal mablaghtable5 { get; set; }
        public decimal mablaghtable6 { get; set; }
        public decimal mablaghtable7 { get; set; }
        public decimal mablaghtable8 { get; set; }
        public decimal mablaghtable9 { get; set; }
        public decimal mablaghtable10 { get; set; }
        public decimal mablaghtable11 { get; set; }
        public decimal mablaghtable12 { get; set; }
        public decimal mablaghtable13 { get; set; }
        public long countmablaghtabletelephon { get; set; }
        public long countbletelephonnot { get; set; }
        public long countbletelephonnot3 { get; set; }

        public long countmablaghtabletelephon2 { get; set; }
        public long countbletelephonnot2 { get; set; }
        public string shenasehReason { get; set; }
        public string shenasehh { get; set; }

        public long counthozor { get; set; }
        public long counthozornot { get; set; }
        public decimal mablaghtabletelephon { get; set; }
        public decimal mablaghtabletelephonnot { get; set; }
        public decimal mablaghtabletelephon2 { get; set; }
        public decimal mablaghtabletelephon3 { get; set; }
    

        public decimal mablaghtabletelephonnot2 { get; set; }
        public decimal mablaghtabletehozor { get; set; }
        public decimal mablaghtabletehozornot { get; set; }



        public long countreason1 { get; set; }
        public long countreason2 { get; set; }
        public long countreason3 { get; set; }
        public long countreason4 { get; set; }
        public decimal mablagreson1 { get; set; }
        public decimal mablagreson2 { get; set; }
        public decimal mablagreson3 { get; set; }
        public decimal mablagreson4 { get; set; }
        public decimal persentrerason1 { get; set; }
        public decimal persentrerason2 { get; set; }
        public decimal persentrerason3 { get; set; }
        public decimal persentrerason4 { get; set; }






        public List<ChartSeries> ChartSeries2 { get; set; }

        public List<dbdashbord1> dbdashbord1 { get; set; }
        public class ChartSeries
        {
            public string name { get; set; }
            public List<double> data { get; set; }
        }
        public List<UserRow> Users { get; set; }
        public string fullname { get; set; }
        public string name_moadel { get; set; }
        public string jobname { get; set; }

        public int cod { get; set; }
        public int usr_ID { get; set; }
        public int usr_ID_person { get; set; }
        public int monthh { get; set; }
        public int yearr { get; set; }
        public int daykol { get; set; }

        public string namecity { get; set; }
        public double moadelkar { get; set; }
        public double moadelkarkol { get; set; }
        public double moadelkarsum { get; set; }

        public List<string> MoalefehTitles { get; set; }
        public List<tbkarkard_notsystem> tbkarkard_notsystem { get; set; }
        public List<tbmadel_kar1> tbmadel_kar1 { get; set; }
     
        public class UserRow
        {

            // برای نمودار اول (درصدی)

            // برای نمودار دوم (مقادیر واقعی job1..job10)
            public List<double> RawValues { get; set; }
            public int UserId { get; set; }
            public string FullName { get; set; }
            public int PersonalId { get; set; }
            public string CityName { get; set; }
            public double Moadel { get; set; }
            public double? MoadelGheirSystem { get; set; }
            public string JobName { get; set; }
            public int datasl { get; set; }
            public string FileSystemName { get; set; }

            public Dictionary<string, decimal> Values { get; set; }
        }
        public List<listmode> listmode1 { get; set; }
        public List<dashbord_talafat> dashbord_talafat { get; set; }

        public class listmode
        {
            public string name { get; set; }
            public long count_0 { get; set; }
            public long count_1 { get; set; }
            public long count_2 { get; set; }
            public long count_3 { get; set; }
            public long personalid { get; set; }
            public string name_person { get; set; }
            public long id { get; set; }


        }
        public Ticket_Barnamerizi3 Ticket_Barnamerizi33 { get; set; }
        public Ticket_TimBarnamenervisi4 Ticket_TimBarnamenervisi44 { get; set; }
        public List<Ticket_TimBarnamenervisi4> Ticket_TimBarnamenervisi4 { get; set; }
        public List<Ticket_motaghazi1> Ticket_motaghazi1 { get; set; }
        public List<UserEstedadViewModel> SubUsersData { get; set; }
        public List<EstedadNiroEnsani> EstedadRecords { get; set; }
        public List<erja_omomi_edit> erja_omomi_edit { get; set; }
        public List<erja_like_link> erja_like_link { get; set; }

        public List<estedad_khodro> estedad_khodro { get; set; }
        public List<Ticket_Barnamerizi3> Ticket_Barnamerizi31 { get; set; }
        public motagayer motagayer2 { get; set; }
        public DateTime? SelectedDate { get; set; }
        public List<valueokala> valueokala { get; set; }
        public List<valueokala1> valueokala1 { get; set; }
        public List<tbCities> Citys { get; set; }
        public List<tbPeymanContracts> tbPeymanContracts { get; set; }

        public List<tbUsers> SubUsers { get; set; }
        public string name_gozaresh { get; set; }
        public int start { get; set; }
        public DateTime fromDate { get; set; }


        public int personID { get; set; }
            public int? FK_usr { get; set; }
            public long? TaghsitBedehi { get; set; }
            public DateTime? Miladi_date1 { get; set; }
            public bool PardakhtShod { get; set; }
        public List<madarek_sabtdoreh_4> madarek_sabtdoreh_4 { get; set; }
        public List<tbSabtEstedadKhodro> tbSabtEstedadKhodro { get; set; }
        public List<Table_Khodro> Table_Khodro { get; set; }
        public Ticket_motaghazi1 Ticket_motaghazi11 { get; set; }
        //public Ticket_IT_2 Ticket_IT_2 { get; set; }
        public List<Ticket_Barnamerizi3> Ticket_Barnamerizi3 { get; set; }
        public List<Ticket_IT_2> Ticket_IT_2 { get; set; }
        public int var1 { get; set; }
        public int var2 { get; set; }
        public int var3 { get; set; }
        public int var4 { get; set; }
        public int var5 { get; set; }
        public int var6 { get; set; }
        public int var7 { get; set; }
        public int var8 { get; set; }
        public int var9 { get; set; }
        public int var10 { get; set; }
        public int var11 { get; set; }
        public int var12 { get; set; }
        public long IDIDID { get; set; }
        public List<tbUsers> tbUsers { get; set; }
        public bool DarbBasteh { get; set; }
            public bool MamanetAzAnjamKar { get; set; }
            public bool EnsheabGhatMovaghat { get; set; }
            public bool EnsheabJamAvaryShodeh { get; set; }
            public bool MoghayratFazBaAmper { get; set; }
            public bool DastkariDarKontor { get; set; }
            public bool KontorKharabAst { get; set; }
            public bool BarghGheyrMojazAzShabakeh { get; set; }
            public bool AzGhalamOftadeh { get; set; }
            public int EslahShomareHamrah { get; set; }
        public string unicc { get; set; }
        public int unicc2 { get; set; }

        public long fkdb12 { get; set; }
        public List<LinkDto> Create_linkrelation2job_3 { get; set; }
        public List<tbAdamAdam> tbAdamAdam23 { get; set; }
        public List<dbergharelimsertfk1_7> dbergharelimsertfk1_7 { get; set; }

        public List<dbergapilike> dbergapilike { get; set; }
        public List<tbTarifKalayeAsasi1> tbTarifKalayeAsasi1 { get; set; }
        public List<tbTarifKalayeAsasi2Estedad> tbTarifKalayeAsasi2Estedad { get; set; }
        public List<tbTarifAbzarKar1> tbTarifAbzarKar1 { get; set; }
        public List<tbTarifAbzarKar2Estedad> tbTarifAbzarKar2Estedad { get; set; }

        public class LinkDto
        {
            public long ID { get; set; }
            public string unic { get; set; }
        }
        public class MarhaleMahdodayatViewModel
        {
            public int FK_usr { get; set; }  // شناسه کاربر برای ساخت id کنترل‌ها
            public string FullName { get; set; }
            public string PeymanTitle { get; set; }
            public string CityName { get; set; }
            public List<MoalefeValueViewModel> MoalefeValues { get; set; }
        }
        public class motagayer
        {
            public int var1 { get; set; }
            public int var2 { get; set; }
            public int var3 { get; set; }
            public int var4 { get; set; }
            public int var5 { get; set; }
            public int var6 { get; set; }
            public int var7 { get; set; }
            public int var8 { get; set; }

        }

        public class MoalefeValueViewModel
        {
            public int MoalefeId { get; set; }
            public string MoalefeTitle { get; set; }
            public long Value { get; set; }
        }
        public List<SoalWithGozinehViewModel> SoalList { get; set; }
        public List<EquipmentSavedValueDto> savedValues { get; set; }

        public string Job { get; set; }
        public string Tahsilat { get; set; }
        public string SabghehJob { get; set; }
        public string PhonNumber { get; set; }
        public string fullnamesabtusr { get; set; }
        public int sokonat { get; set; }
        public int  usrID { get; set; }

        public Header SabtHeader { get; set; }
        public year year { get; set; }
        public number number { get; set; }
        public id id { get; set; }
        public tim tim { get; set; }
        public long IDbank { get; set; }
        public long sumentezar { get; set; }
        public long sumtaeed { get; set; }
        public long sumersal { get; set; }

        public string namebank { get; set; }
        public string hesabbank { get; set; }
        public DateTime timeroz { get; set; }
        public class SoalWithGozinehViewModel
        {
            public long SoalId { get; set; }
            public string SoalText { get; set; }
            public List<GozinehItem> GozinehList { get; set; }
            public string SharhSoal { get; set; } // اضافه شده

        }

        public class GozinehItem
        {
            public long GozinehId { get; set; }
            public string GozinehText { get; set; }
            public bool Selected { get; set; }
        }
       
        public int ID { get; set; }
        public string Moalfe { get; set; }
        public int name2 { get; set; }
        public int idmaolfeh { get; set; }
        public string namemoalfeh { get; set; }
        public int  countsabt { get; set; }
        public int notsabt { get; set; }
        public int month { get; set; }
        public int year333 { get; set; }
        public int Month1 { get; set; }
        public int Year1 { get; set; }
        public List<detailjob> detailjob { get; set; }
        public List<detailjob1> detailjob1 { get; set; }
        public List<MoalefeUserInfo> UsersMoalefe { get; set; }
        public List<MoalefeUserInfo22> fortariftatil { get; set; }
        public List<dbAsnafsherkatsanad> dbAsnafsherkatsanad { get; set; }
        public List<tbmarahesabt41> tbmarahesabt4 { get; set; }
        public arzyzbi8_edit arzyzbi8_edit { get; set; }
        public int cityusr { get; set; }

        public string shomarehheshterak { get; set; }
        public string ComputerPassword { get; set; }



        public List<tbdetail> tbdetail { get; set; }
        public List<mahdoday> mahdoday { get; set; }
        public List<mahdadatballla> mahdadatballla { get; set; }
        public List<tickmomahdodat> tickmomahdodat { get; set; }
        public List<tickmomahdodatkenar> tickmomahdodatkenar { get; set; }
        public List<tickmomahdodatkenar2> tickmomahdodatkenar2 { get; set; }
        public List<DataListItem> DataList { get; set; }
        public List<santnashodeh> santnashodeh1 { get; set; }

        public List<sabttaeed> sabttaeedview { get; set; }
        public List<listinttb4> listinttb4 { get; set; }
        public List<Agarinvaziat> Agarinvaziat { get; set; }
        public List<sathdastresi> sathdastresi { get; set; }
        public List<gozaresh> gozaresh { get; set; }
        public List<dbergharel4_5> dbergharel4_5 { get; set; }

        public List<dbergharel_4_4> dbergharel_4_4 { get; set; }

        public List<valuuarzyabi> valuuarzyabi { get; set; }
        public List<arzyzbi5_time> arzyzbi5_time { get; set; }
        public List<tbAdamAdam> tbAdamAdam { get; set; }
        public List<arzyzbi5_time2> arzyzbi5_time2 { get; set; }
        public List<tbAdamAdam3> tbAdamAdam3 { get; set; }
        public List<sabttaeedviewsoton> sabttaeedviewsoton { get; set; }
        public List<sabttaeedviewAdama> sabttaeedviewAdama { get; set; }

        public List<tbAdamAdam2> tbAdamAdam2 { get; set; }

        public List<gozareshmally> gozareshmally { get; set; }
        public List<MarhaleViewModel> MarhaleViewModel { get; set; }
        public EditEvaluationViewModel EvaluationData { get; set; } // ← اضافه کن

        public List<Dastehfish> Dastehfish { get; set; }
        public class EditEvaluationViewModel
        {
            public long ID { get; set; }
            public string Title { get; set; }
            public int typegozarwsh { get; set; }

            public List<EditQuestionViewModel> Questions { get; set; }
        }

        public class EditQuestionViewModel
        {
            public long? ID { get; set; }
            public string Text { get; set; }
            public string Description { get; set; }
            public List<EditOptionViewModel> Options { get; set; }
        }

        public class EditOptionViewModel
        {
            public long? ID { get; set; }
            public string Text { get; set; }
            public int? Weight { get; set; }
        }
        public LinkafradmoalfehViewModel LinkafradmoalfehData { get; set; }

        public class LinkafradmoalfehViewModel
        {
            public List<MahdodayViewModel> MahdodayList { get; set; }
            public List<TickMoadelViewModel> TickList { get; set; }
            public List<ContractUserViewModel> ContractUsers { get; set; }
        }
  
        public class MahdodayViewModel
        {
            public int ID { get; set; }
            public int FK_moalfe { get; set; }
            public string Title { get; set; }
        }
        public class EquipmentSavedValueDto
        {
            public int FK_Equipment { get; set; }
            public int Fk_user { get; set; }
            public double CountDays { get; set; }
        }
        public class TickMoadelViewModel
        {
            public int FK_name { get; set; }
            public int FK_moalfeh { get; set; }
            public bool value { get; set; }
        }

        public class ContractUserViewModel
        {
            public int FK_usr { get; set; }
            public string FullName { get; set; }
            public int PersonalID { get; set; }
        }
    }
    public class UserBaleViewModel
    {
        public int UserId { get; set; }
        public string FullName { get; set; }
        public long ChatId { get; set; }
        public string BaleUsername { get; set; }
    }
    public class detailjob1
    {

        public string IDst { get; set; }
        public long ID { get; set; }
        public long FK_usr { get; set; }
        public long vaziat { get; set; }


    }
    public class detailjob
    {
        public long ID { get; set; }
        public string name { get; set; }
        public string IDst { get; set; }
        public double Lat { get; set; }
        public double Lng { get; set; }

        public double PrevLat { get; set; }
        public double PrevLng { get; set; }

        public string detal { get; set; }
        public long ID_link4 { get; set; }
        public long fild { get; set; }
        public long counttt { get; set; }

    }
    public class Marhale4ViewModel
    {
        public List<tbUsers> Users { get; set; }
        public List<tbCities> Cities { get; set; }
        public List<Link_User_And_Peyman> UserPeymans { get; set; }
        public List<dbtarifmahdodayt3> Marhale3 { get; set; }
        public List<dbtarifmahdodayt4> Values { get; set; }
    }
    public class Marhale5ViewModel
    {
        public List<tbPeymanContracts> Peymans { get; set; }
        public List<tbpeymancities> Cities { get; set; }
        public List<dbtarifmahdodayt3> Marhale3 { get; set; }
        public List<dbtarifmahdodayt1> Moalefs { get; set; }
        public List<dbtarifmahdodayt5> Values { get; set; }
    }
    public class Marhale2ViewModel
    {
        public List<UserRowVM> Users { get; set; }
        public List<MoalefeVM> Moalefeh { get; set; }
        public List<ValueVM> Values { get; set; }
    }

    public class UserRowVM
    {
        public int UserId { get; set; }
        public string FullName { get; set; }
        public string City { get; set; }
        public string Peyman { get; set; }
    }

    public class MoalefeVM
    {
        public int Id { get; set; }
        public string Title { get; set; }
    }

    public class ValueVM
    {
        public int? FK_usr { get; set; }
        public int? FK_moalfe { get; set; }

        public decimal? Value { get; set; }
    }
    //public class SoalWithGozinehViewModel
    //{
    //    public int SoalId { get; set; }
    //    public string SoalText { get; set; }
    //    public List<GozinehItem> GozinehList { get; set; }
    //}

    //public class GozinehItem
    //{
    //    public int GozinehId { get; set; }
    //    public string GozinehText { get; set; }
    //    public bool Selected { get; set; }
    //}
    public class valueokala
    {
        public long ID { get; set; }
        public long count { get; set; }

        public long countentezar { get; set; }
        public long countersal { get; set; }
        public long countersaltaeed { get; set; }

    }
    public class valueokala1
    {
        public long ID { get; set; }
        public long count { get; set; }

        public long countentezar { get; set; }
        public long countersal { get; set; }
        public long countersaltaeed { get; set; }

    }
    public class MarhaleViewModel
    {
        public int ID { get; set; }
        public string PeymanTitle { get; set; }
        public string CityName { get; set; }
        public string Naem { get; set; }
        public int Number { get; set; }
        public bool Sabt { get; set; }
        public bool Control { get; set; }
        public string vaziat { get; set; }

    }



    public class madod2
    {
        public tbAdamAdam tbAdamAdam { get; set; }
        public int personalID { get; set; }

        public string fullname { get; set; }
        public int ID { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public int ID4 { get; set; }
        public int fK_pymn { get; set; }

        public int ID1 { get; set; }
    }
    public class santnashodeh
    {
        public int personalID { get; set; }
        public int ID { get; set; }
        public string fullname { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public int vaziat { get; set; }
        


    }
    public class listinttb4
    {
        public int fkcity { get; set; }
        public int fkpymn { get; set; }

        public int IDBasteh { get; set; }
        public int IDTB4 { get; set; }
    }
    public class tbAdamAdam2
    {
        public string fullname { get; set; }
        public int ID { get; set; }

   
    }
    public class arzyzbi5_time2
    {
        public arzyzbi5_time arzyzbi5_Time { get; set; }
        
        public string fullname { get; set; }
        public int ID { get; set; }


    }
    public class tbAdamAdam3
    {
        public tbAdamAdam tbAdamAdam { get; set; }

        public string fullname { get; set; }
        public int ID { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public int ID4 { get; set; }

        public int ID1 { get; set; }
    }
    public class sabttaeedviewAdama
    {
        public tbAdamAdam tbAdamAdam { get; set; }
        public int personalID { get; set; }

        public string fullname { get; set; }
        public int ID { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public int ID4 { get; set; }
        public int fK_pymn { get; set; }
       public int valuekarbar { get; set; }
        public int ID1 { get; set; }
    }
    public class sabttaeedviewsoton
    {
        public tbAdamAdam tbAdamAdam { get; set; }

        public string fullname { get; set; }
        public int ID { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public int ID4 { get; set; }
        public int personalID { get; set; }

        public int ID1 { get; set; }
    }

    public class gozareshmally
    {
        public double SUM { get; set; }
        public double SUMBEDEHKAR { get; set; }

        public double SUMBESTENKAR { get; set; }

        public double SUMBEDEHKAR2 { get; set; }

        public double SUMBESTENKAR2 { get; set; }
        public double SUMBEDEHKAR4 { get; set; }

        public double SUMBESTENKAR4 { get; set; }

        public long  FK_ID { get; set; }
        public long ID { get; set; }

        public string coding { get; set; }
        public string namecoding { get; set; }

        public int IDTB4 { get; set; }
    }
    public class Agarinvaziat
    {
        public int IDPymn { get; set; }
        public string Namepymn { get; set; }
        public int countend { get; set; }
        public int countdargaryan { get; set; }

    }
    public class sathdastresi
    {
        public int IDPymn { get; set; }
        public string fullname { get; set; }
        public string  job { get; set; }
        public string tab { get; set; }
        public string zarmagmetab { get; set; }
        public string zarzermagmetab { get; set; }

        public int codpersanly { get; set; }
        public string namepymn { get; set; }
        public int IDyusr { get; set; }
        public int tabint { get; set; }
        public int zarmagmetabint { get; set; }
        public int zarmazaergmetabint { get; set; }


    }
    public class sabttaeed
    {
        public int IDcit { get; set; }
        public int IDPymn { get; set; }
        public string Namepymn { get; set; }
        public string namecity { get; set; }
        public int sabt { get; set; }
        public int taeed { get; set; }
        public int control { get; set; }
        public int IDBasteh { get; set; }
        public int IDTB4 { get; set; }
        public int month { get; set; }
        public int year { get; set; }

        public int vaziattt { get; set; }
        public int number { get; set; }

    }
    public class valuuarzyabi
    {
        //public long? ID { get; set; }
        //public string Text { get; set; }
        public List<arzyzbi4_link> arzyzbi4_link { get; set; }
        public List<arzyzbi5_time> arzyzbi5_time { get; set; }
    }
    public class gozaresh1detail
    {
        public int IDcit { get; set; }
        public int IDPymn { get; set; }
        public string Namepymn { get; set; }
        public string namecity { get; set; }
        public int sabt { get; set; }
        public int taeed { get; set; }
        public int control { get; set; }
        public int IDBasteh { get; set; }
        public int IDTB4 { get; set; }
        public int month { get; set; }
        public int year { get; set; }

        public int vaziattt { get; set; }
        public int number { get; set; }


    }

    public class DataListItem
    {
        public int CityID { get; set; }
        public int PymnID { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int IDPersonal { get; set; }
        public int FkMoalfe { get; set; }
        public int sabt { get; set; }

        public double value { get; set; }
    }
    public class gozaresh
    {
        public tbUsers tbusr { get; set; }

    }
    public class Dastehfish
    {
        public int ID { get; set; }
        public int CityID { get; set; }
        public int PymnID { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int IDPersonal { get; set; }
        public int FkMoalfe { get; set; }
        public int sabt { get; set; }
        public string city { get; set; }
        public string pumn { get; set; }
        public string fullname { get; set; }
        public string shogl { get; set; }
        public string stringshogk { get; set; }

        public double value { get; set; }
    }



    public  class tbmarahesabt41
    {
     public tbmarahesabt4 tbmarahesabt4 { get; set; }
        public string vazeiat { get; set; }
        public int day {  get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public string DAY1 { get; set; }
        public int insabt { get; set; }
        public int nadedh { get; set; }
        public int kol { get; set; }

    }
    public class tbdetail
    {
        public tbEquipmentMoalefeValueReffrenceSave tbEquipmentMoalefeValueReffrenceSave { get; set; }
        public tbsaveSoratBasteh tbsaveSoratBasteh { get; set; }

        public tbSavedFunctions tbSavedFunctions { get; set; }
        public tbmoalfefishexcel tbmoalfefishexcel { get; set; }
        public tbmarahesabt4 tbmarahesabt4 { get; set; }
        public string fullname { get; set; }
        public string personalID { get; set; }
        public string marhaleh { get; set; }
        public string vaziat { get; set; }
        public int vazrt { get; set; }
        public int day { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public string daysabt { get; set; }
        public string namebasteh { get; set; }
        public string namemoalfeh { get; set; }
        public int idmoalfeh { get; set; }
        public int IDusr { get; set; }

        public int daymondeh { get; set; }

    }
    public class year
    {
        public int MoalfeVal_Year { get; set; }
        public int month { get; set; }

    }
    public class mahdadatballla
    {
        public int IDbasteh { get; set; }
        public string nameVBasteh { get; set; }
        public int value { get; set; }

    }
    public class tickmomahdodat
    {
        public int IDMOalfeh { get; set; }
        public string namemoalfeh { get; set; }
        public string nomoalfeh { get; set; }
        public double valueMax { get; set; }
        public double valueMiyang { get; set; }

        public int value { get; set; }

    }


    public class tickmomahdodatkenar
    {
        public int IDMOalfeh { get; set; }
        public string namemoalfeh { get; set; }
        public string namebasteh { get; set; }
        public int IDBasteh { get; set; }

        public double value { get; set; }

    }
    public class tickmomahdodatkenar2
    {
        public int IDMOalfeh { get; set; }
        public string namemoalfeh { get; set; }
        public string namebasteh { get; set; }
        public int IDBasteh { get; set; }

        public double value { get; set; }

    }

    public class mahdoday
    {
        public int IDpersonal { get; set; }
        public int PersonalID { get; set; }
        public string fullnam { get; set; }
        public string pymanname { get; set; }
        public int PymnID { get; set; }
        public int cityID { get; set; }
        public string CITynam { get; set; }
        public string bastehname { get; set; }
        public int bastehID { get; set; }
        public double seghfBasteh { get; set; }
        public int seghfBastehID { get; set; }

        public int fkmoalfe { get; set; }
        public int IDtackBasteh { get; set; }
        public string nametackBasteh { get; set; }
        public double moadel { get; set; }


        public double value22 { get; set; }
        public long value { get; set; }
        public double seghf { get; set; }
        public long bigsaghf { get; set; }
        public int IDmarhaleh { get; set; }
        public int month { get; set; }
        public int year { get; set; }

    }
    public class id
    {
        public string id2 { get; set; }
        public DateTime month { get; set; }
        public int MoalfeVal_Year { get; set; }
        public string bindValue { get; set; }
    }
    public class tim
    {
        public int id2 { get; set; }
        public int month { get; set; }
        public int MoalfeVal_Year { get; set; }

    }
    public class number
    {
        public string numberstr { get; set; }

    }
    public class Header  {      
        public int MoalfeVal_Month { get; set; }   
        public int MoalfeVal_Year { get; set; }   
        public string PeymanName { get; set; }    
        public string BastehName { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
    } 
    public class MoalefeUserInfo  {     
        public string FullName { get; set; } 
        public int PersonalCode { get; set; }   
        public List<MoalefeInfo> listMoalefe{ get; set; } 
    } 
    public class MoalefeInfo  {      
        public string MoalefeTitle { get; set; }  
        public string MoalefeValue { get; set; }
        public int Moalefenum { get; set; }

    }
    public class MoalefeUserInfo2
    {
        public string FullName { get; set; }
        public int PersonalCode { get; set; }
        public List<MoalefeInfo2> listMoalefe { get; set; }
    }
    public class MoalefeInfo2
    {
        public string MoalefeTitle { get; set; }
        public string MoalefeValue { get; set; }
        public int Moalefenum { get; set; }
    }
    public class FunctionModel2
    {
        public Header SabtHeader { get; set; }
        public List<MoalefeUserInfo2> UsersMoalefe2 { get; set; }
    }
    public class MoalefeUserInfo22
    {
        public string FullName { get; set; }
        public int PersonalCode { get; set; }
        public int anjam { get; set; }
        public int notanjam { get; set; }
        public int tagher { get; set; }
        public int darjaryan { get; set; }
        public int adam { get; set; }
        public int cant { get; set; }
        public int ID { get; set; }
        public int pyam { get; set; }
        public DateTime inmonth { get; set; }
        public DateTime nextmonth { get; set; }
    }
}
