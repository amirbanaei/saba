using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbVosoolMotalebatMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbVosoolMotalebatMetaData))]
    public partial class tbVosoolMotalebat : ViewModelBase
    {
       public string peymanBrief
        {
            get
            {
                using (var db = new SaabEntities())
                {
                    return db.tbPeymanContracts.Find(FK_PeymanID).pec_BriefTitle;
                }
                 
            }
        }

    }
}