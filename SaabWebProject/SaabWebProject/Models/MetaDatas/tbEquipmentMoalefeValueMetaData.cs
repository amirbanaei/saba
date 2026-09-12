using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbEquipmentMoalefeValueMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbEquipmentMoalefeValueMetaData))]
    public partial class tbEquipmentMoalefeValue : ViewModelBase
    {
        public string Persian_FromDate
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(FromDate);
            }
        }
        public string Persian_ToDate
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(ToDate);
            }
        }

    }
}