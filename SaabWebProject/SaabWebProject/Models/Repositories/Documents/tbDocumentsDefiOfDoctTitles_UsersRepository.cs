using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Documents
{
    public class tbDocumentsDefiOfDoctTitles_UsersRepository : IInterFace<tbDocumentsDefiOfDoctTitles_Users>
    {
        #region متغیرها
        SaabEntities db;
        #endregion

        #region سازنده ها
        public tbDocumentsDefiOfDoctTitles_UsersRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDocumentsDefiOfDoctTitles_UsersRepository()
        {
            db = new SaabEntities();
        }
        #endregion

        public string Create(tbDocumentsDefiOfDoctTitles_Users obj)
        {
            try
            {
                db.tbDocumentsDefiOfDoctTitles_Users.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public string Create(List<tbDocumentsDefiOfDoctTitles_Users> obj)
        {
            try
            {
                var new_list = new List<tbDocumentsDefiOfDoctTitles_Users>();
                foreach(var item in obj)
                {
                    if(db.tbDocumentsDefiOfDoctTitles_Users.Where(p=>p.FK_DefinitionOfDocTitle == item.FK_DefinitionOfDocTitle &&
                    p.FK_User == item.FK_User
                    && p.IsSabt == item.IsSabt && p.IsTaeed==item.IsTaeed).FirstOrDefault() == null)
                    {
                        db.tbDocumentsDefiOfDoctTitles_Users.Add(item);
                    }
                }
                //var peyman_ID = db.tbDocumentsDefinitionOfDocumentTitles.Find(obj.FirstOrDefault().FK_DefinitionOfDocTitle).FK_PeymanID;
                //var ff = db.tbDocumentsDefiOfDoctTitles_Users.Where(p => p.tbDocumentsDefinitionOfDocumentTitles.FK_PeymanID == peyman_ID);
                //if (ff.FirstOrDefault() != null)
                //{
                //    db.tbDocumentsDefiOfDoctTitles_Users.RemoveRange(db.tbDocumentsDefiOfDoctTitles_Users.Where(p => p.tbDocumentsDefinitionOfDocumentTitles.FK_PeymanID == peyman_ID).ToList());
                //}
                //db.tbDocumentsDefiOfDoctTitles_Users.AddRange(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public bool Disable(int ID)
        {
            db.tbDocumentsDefiOfDoctTitles_Users.Remove(db.tbDocumentsDefiOfDoctTitles_Users.Find(ID));
            return SaveChanges();
        }

        public tbDocumentsDefiOfDoctTitles_Users Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbDocumentsDefiOfDoctTitles_Users> Update()
        {
            return db.tbDocumentsDefiOfDoctTitles_Users.ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbDocumentsDefiOfDoctTitles_Users obj)
        {
            throw new NotImplementedException();
        }

        public List<tbPeymanContracts> GetPeymansForThisUser_Saver(int user_ID)
        {
          return  db.tbDocumentsDefiOfDoctTitles_Users.Where(p => p.IsSabt == true && p.FK_User == user_ID).Select(p => p.tbDocumentsDefinitionOfDocumentTitles.tbPeymanContracts).ToList().Distinct().ToList();
        }
        public IEnumerable<tbPeymanContracts> GetPeymansForThisUser_Accepter(int user_ID)
        {
          var a =  db.tbDocumentsDefiOfDoctTitles_Users.Where(p => p.IsTaeed == true && p.FK_User == user_ID).ToList();
            return a.Select(p => p.tbDocumentsDefinitionOfDocumentTitles.tbPeymanContracts);
        }


        
    }
}