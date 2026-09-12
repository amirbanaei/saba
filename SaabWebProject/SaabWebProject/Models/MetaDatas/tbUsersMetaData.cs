
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbUsersMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbUsersMetaData))]
    public partial class tbUsers : ViewModelBase, IDisposable
    {
        public string FullName
        {
            get => usr_Name + " " + usr_Family;
        }
        public bool usr_RememberMe { get; set; }
        public string shamsiDateOfBirth
        {

            get
            {
                DateTime dt = usr_DateOfBrith ?? DateTime.Now;
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(dt);
            }
        }
        public void Dispose()
        {
            // throw new NotImplementedException();
        }
    }
}