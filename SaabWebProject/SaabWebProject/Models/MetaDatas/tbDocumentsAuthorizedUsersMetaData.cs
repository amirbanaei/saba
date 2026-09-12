using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbDocumentsAuthorizedUsersMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbDocumentsAuthorizedUsersMetaData))]
    public partial class tbDocumentsAuthorizedUsers : ViewModelBase, IDisposable
    {
       
        public List<tbUsers> listUsers
        {
            get
            {
                List<tbUsers> users = new List<tbUsers>();
                using(SaabEntities db= new SaabEntities())
                {
                    foreach(var item in db.tbDocumentsAuthorizedUsers.Where(p => p.FK_PeymanID == FK_PeymanID).ToList())
                    {
                        users.Add(item.tbUsers);
                    }
                }
                return users;
            }
        }
        public void Dispose()
        {
        }
    }
}