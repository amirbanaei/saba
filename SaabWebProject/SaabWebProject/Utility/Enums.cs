namespace SaabWebProject.Utility
{
    public enum Type_Moalefe
    {
        Gharardadi = 1,
        Amalkardi = 2,
        Karbari = 3,
        Sayer = 4,
        Controlli = 5,
        Calculational =6,

        Dastmozdi=7,
        Karkardi=8,
        Pishkhan=9,
        Fish=10,
        sorat = 11
    }
    public enum UnitTime
    {
        
        Hour=1,
        Day=2,
        Month=3,
        Year=4
    }
    public enum TazminType
    {
        check=1,
        safte,
        zemanatnamebanki,
        vajgnaghd,
        namezemanat
    }
    public enum SalaryDetail
    {
        Null=0,
        HaveContract,
        Consolidation,
        PerformanceSave,
        PerformanceAccept
    }
    
}
namespace System
{
    public static class Gender_Moalefe
    {
        public static string Int { get => "int"; }
        public static string String { get => "string"; }
        public static string Date { get => "date"; }
        public static string Time { get => "time"; }
        public static string Tik { get => "Tik"; }
    }

    
}