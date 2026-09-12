using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbPeymanElhaghieMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbPeymanElhaghieMetaData))]
    public partial class tbPeymanElhaghie : ViewModelBase, IDisposable
    {
        public string pec_Price { get; set; }
        public int EndTime_day
        {
            get; set;
        }
        public int EndTime_month
        {
            get; set;
        }
        public int EndTime_year
        { get; set;
        }

        public string ShmasiENDTime
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(EndTime);
            }
        }
        public void Dispose()
        {
            // throw new NotImplementedException();
        }
    }
}