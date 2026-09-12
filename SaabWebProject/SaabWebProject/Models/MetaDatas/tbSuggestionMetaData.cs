using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbSuggestionMetaData
    {
    }
}
namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbSuggestionMetaData))]
    public partial class tbSuggestion : ViewModelBase, IDisposable
    {

        public string Persian_Date
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(Datetime);
            }
        }
        public string UserFullName
        {
            get
            {
                using(var db = new SaabEntities())
                {
                    return db.tbUsers.Find(FK_UserID).FullName;
                }
            }
        }
        public int HasNotSeen { get; set; }
        public void Dispose()
        {
        }
    }
}