using System.Collections.Generic;

namespace SaabWebProject.Models.ViewModels.Statements
{
    public class ReportStatmentsViewModel
    {
        public RSHeader Header_0 { get; set; }
        public RSContract Contracs_1 { get; set; }
        public RSAmalkerdEntezar AmalkerdEntezar_2 { get; set; }
        public RSNetFee NetFee_3 { get; set; }
        public RSControlBasteks ControlBasteks_4 { get; set; }
        public Description decriptionsabad { get; set; }
        public Comparisonbasket Comparison6 { get; set; }
        public basket basket7 { get; set; }
        public basket9 basket9 { get; set; }
        public basket10 basket10 { get; set; }
        public sorat basss { get; set; }
        public int month {  get; set; }
        public int year { get; set; }
        public long final { get; set; }
        public long  numbersorat { get; set; }
        public long padash { get; set; }
        public long grameh { get; set; }
        public double gareme2 { get; set; }

        public double padash2 { get; set; }
        public double padashkol { get; set; }
        public double garemehkol { get; set; }
       public int numbersoratt { get; set; }

        public long eydipadash { get; set; }
        public long zarib { get; set; }
        public long bemehtakmily { get; set; }
        public long bemehtakmilyzarib { get; set; }
        public long zarib22 { get; set; }
        public long restrictions_A { get; set; }
        public long restrictions_B { get; set; }

        public long restrictions_C { get; set; }
        public int vaziat { get; set; }
        public long subkoljob { get; set; }
        public long subkoljobarzesh { get; set; }
        public long eydipadashkol { get; set; }
        public long sanavatkol { get; set; }

        //public long restrictions_C { get; set; }
        //public long restrictions_C { get; set; }

        public long restrictions_D { get; set; }
        public float sumrestrictions { get; set; }
        public List<jobSubCirties> jobSubCirties { get; set; }
        public List<arzaeshsubCirties> arzaeshsubCirties { get; set; }

        public List<JobSumViewModel> JobSumViewModel { get; set; }

    }
    public class JobSumViewModel
    {
        public int Id { get; set; } // کد از 1 شروع میشه
        public string JobTitle { get; set; } // نام شغل
        public decimal TotalValue { get; set; } // جمع مقدار
    }
    public class RsBaskets
    {
        public string Number { get; set; }
        public string Name { get; set; }
    }

    public class RsBas
    {
        public string Name { get; set; }
        public long Sumsabad {  get; set; }
    }
    #region Header_0
    public class RSHeader
    {
        public int numberOfSooratVaziat { get; set; }
        public string FromDate { get; set; }
        public string toDate { get; set; }
        public string Peymankar { get; set; }
        public string KargozariName { get; set; }
        public string NameOfcubCities { get; set; }
        public int KargozariCode { get; set; }
        public string kargoz { get; set; }

    }
    #endregion

    #region Contracs_1
    public class RSContract
    {
        public RSContractInformation EnergyInformation { get; set; }
        public RSContractInformation PriceContractInformation { get; set; }
        public List<RSContractBasketsInformation> Baskets { get; set; }
    }

    public class RSContractInformation
    {
        public string Base { get; set; }
        public string Elhaghie { get; set; }
        public string Final { get; set; }
        public string TaeedShode { get; set; }
        public string Mande { get; set; }
    }
    public class RSContractInformation2
    {
        public string Base { get; set; }
        public string Elhaghie { get; set; }
        public string Final { get; set; }
        public string TaeedShode { get; set; }
        public string Mande { get; set; }
    }
    public class RSContractBasketsInformation
    {
        public string basket { get; set; }
        public long Price { get; set; }
    }
    #endregion 

    #region AmalkerdEntezar_2
    public class RSAmalkerdEntezar
    {
        public List<RSAmalkerdEntezarSubCirties> amalkerdEntezarSubCirties { get; set; }
    }
    public class RSAmalkerdEntezarSubCirties
    {
        public string SubCityName { get; set; }
        public long amalkerd_EnergyTahvili { get; set; }
        public long amalkerd_EnergyToziee { get; set; }
        public long amalkerd_EnergyToziee2 { get; set; }
        public long amalkerd_Foroosh { get; set; }
        public float amalkerd_Vosool { get; set; }
        public float amalkerd_VosoolPercent { get; set; }
        //------------------------------------------------
        public float Entezar_EnergyTahvili { get; set; }
        public float Entezar_EnergyToziee { get; set; }
        public float Entezar_EnergyTozieeEslahi { get; set; }
        public float Entezar_VosoolPercent { get; set; }
        public float Entezar_talafatPercent_Paziresh { get; set; }
        public float Entezar_talafatPercnt_hadaf { get; set; }

        //------------------------------------------------
        public float Enheraft_Energy { get; set; }
        public double Enheraft_Energy2 { get; set; }
        public double Enheraft_VosoolPercent { get; set; }
    }
    public class jobSubCirties
    {
        public string SubCityName { get; set; }
        public long shogl1 { get; set; }
        public long shogl2 { get; set; }
        public long shogl3 { get; set; }
        public long shogl4 { get; set; }
        public long shogl5 { get; set; }
        public long shogl6 { get; set; }
        //------------------------------------------------
        public long shogl7 { get; set; }
        public long shogl8 { get; set; }
        public long shogl9 { get; set; }
        public long shogl10 { get; set; }
        public long arzeshafzodeh { get; set; }
        public long eydi { get; set; }

        //------------------------------------------------
        public long sanavat { get; set; }
        public long sumjob { get; set; }
        public double Enheraft_VosoolPercent { get; set; }
    }
    public class arzaeshsubCirties
    {
        public string SubCityName { get; set; }
        public long shogl1 { get; set; }
        public long shogl2 { get; set; }
        public long shogl3 { get; set; }
        public long shogl4 { get; set; }
        public long shogl5 { get; set; }
        public long shogl6 { get; set; }
        //------------------------------------------------
        public long shogl7 { get; set; }
        public long shogl8 { get; set; }
        public long shogl9 { get; set; }
        public long shogl10 { get; set; }
        public float arzeshafzodeh { get; set; }
        public float eydi { get; set; }

        //------------------------------------------------
        public float sanavat { get; set; }
        public double moraghasi { get; set; }
        public double Enheraft_VosoolPercent { get; set; }
    }

    #endregion

    #region NetFee_3
    public class RSNetFee
    {
        public List<RSNetFeeSooratVaziatHistory> SooratVaziatHistory { get; set; }
        public string SUM_Energy { get; set; }
        public string SUM_PriceKham { get; set; }
    }
    public class RSNetFeeSooratVaziatHistory
    {
        public string Name { get; set; }
        public string Energy { get; set; }
        public string Nerkh { get; set; }
        public string PriceKham { get; set; }

    }
    #endregion

    #region سبد های هزینه ای
    public class RSControlBasteks
    {
        public List<RsSooratVaziat> list_SooratVaziat { get; set; }
        public string EydiPadashSanavatTajamoee { get; set; }
        public long Sumsabad { get; set; }
        public long Sumamal { get; set; }
        public long Sumhadd { get; set; }
        public long tafavot { get; set; }


        //این را باید بپرسیم 
    }
    public class RsSooratVaziat
    {
        public int number { get; set; }

        public List<RSHazineBaskets> listHzineBasket { get; set; }
        public long SUM_Mamoorin { get; set; }
        public long SUM_NamayandeMoqim { get; set; }
        public long SUM_Karkerd { get; set; }
        public long SUM_Total { get; set; }
        public long SUM_amalfard { get; set; }
        public long SUM_had { get; set; }
        public long SUM_tafavot { get; set; }



    }
    public class sorat
    {
        public List<soart1> sorat2 { get; set; }
    }
    public class soart1
    {
        public int number { get; set; }
        public long Energy { get; set; }
        public long Nerkh { get; set; }
        public long PriceKham { get; set; }


        public List<soaratvass> sorat2 { get; set; }
    }
    public class soaratvass
    {
        public RsBas basket { get; set; }
        public long Mamoorin { get; set; }
        public long NamayandeMoqim { get; set; }
        public long KadrEdari
        {
            get; set;
        }
        public long SumMamoorin { get; set; }
        public long SumNamayandeMoqim { get; set; }
        public long SumKadrEdari
        {
            get; set;
        }
        public long sumsabad { get; set; }


    }
    public class RSHazineBaskets
    {
        public RsBas basket { get; set; }
        public long Mamoorin { get; set; }
        public long NamayandeMoqim { get; set; }
        public long KadrEdari
        {
            get; set;
        }
        public long chom
        {
            get; set;
        }
    }
    public class Description
    {
        public List<ListDescrip> List_Descrip { get; set; }
        public long SumDescrip { get; set; }
    }
    public class ListDescrip
    {
        public double value {  get; set; }
        public string Description {  get; set; }
        public string name { get; set; }
    }
    public class Comparisonbasket
    {
        public List<RsSooratVaziat> list_SooratVaziat { get; set; }
        public long kole { get; set; } //این را باید بپرسیم 
    }

    public class Soratdetail
    {
        public string name { get; set; }
        public long value { get; set; }
    }

    public class basket
    {
        public long cham { get; set; }
        public float zarib { get; set; }
        public long aftereslahat { get; set; }

        public long chabli { get; set; }
        public long aval { get; set; }



        //این را باید بپرسیم 
    }
    public class basket9
    {
        public long colpadash { get; set; }
        public float garame { get; set; }
        public long mande { get; set; }

        public long mahdodeatA { get; set; }
        public long avaleh { get; set; }
        public long mahdodeatB { get; set; }
        public long mahdodeatC { get; set; }
        public long mahdodeatD { get; set; }

        public long summahdod { get; set; }
        public long finalsorat { get; set; }


        //این را باید بپرسیم 
    }
    public class basket10
    {
        public long mabladhcol { get; set; }
        public float coltaeed { get; set; }
        public long gess { get; set; }

        public long enhraf { get; set; }
        public long save { get; set; }
        //public long mahdodeatB { get; set; }
        //public long mahdodeatC { get; set; }
        //public long mahdodeatD { get; set; }

        //public long summahdod { get; set; }
        //public long finalsorat { get; set; }


        //این را باید بپرسیم 
    }
    #endregion
}