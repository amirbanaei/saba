using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbUserContractsMetaData
    {
    }
}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbUserContractsMetaData))]
    public partial class tbUserContracts : ViewModelBase, IDisposable
    {
        public int ShamsiStartTime_day { get; set; }
        public int ShamsiStartTime_month
        {
            get; set;

        }
        public int ShamsiStartTime_year
        {
            get; set;
        }
        public int ShamsiENDTime_day
        {
            get; set;
        }
        //public int ShamsiENDTime_month
        //{
        //    get
        //    {
        //        PersianCalendar pc = new PersianCalendar();
        //        DateTime dt = usc_EndTime ?? DateTime.Now;
        //       return pc.GetMonth(dt);
               
        //    }
        //}
        public int ShamsiENDTime_month
        {
            get; set;

        }
        public int ShamsiENDTime_year
        {
            get; set;
        }
        public string usc_ShmasiStartTime
        {
            get
            {
                DateTime dt = usc_StartTime ?? DateTime.Now;
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(dt);
            }
        }
        public string usc_ShmasiENDTime
        {
            get 
            { 
                DateTime dt = usc_EndTime ?? DateTime.Now;
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(dt);
            }
        }
        public void Dispose()
        {
            // throw new NotImplementedException();
        }
    }
}