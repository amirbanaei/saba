using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbPeymanContractsMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbPeymanContractsMetaData))]
    public partial class tbPeymanContracts : ViewModelBase, IDisposable
    {
        public int pec_StartTime_day {get;set;}
        public int pec_StartTime_month
        {
            get; set;

        }
        public int pec_StartTime_year
        {
            get; set;

        }
        public int pec_ENDTime_day
        {
            get; set;

        }
        public int pec_ENDTime_month {
            get; set;

        }
        public int pec_ENDTime_year {
            get; set;

        }
        public string pec_ShmasiStartTime
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(pec_StartTime);
            }
        }
        public string pec_ShmasiENDTime
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(pec_EndTime);
            }
        }
        public void Dispose()
        {
            // throw new NotImplementedException();
        }
    }
}