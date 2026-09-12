using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Documents
{
    public class tbDocumentsRepository : IInterFace<tbDocuments>
    {
        #region متغیرها
        SaabEntities db;

        #endregion

        #region سازنده ها
        public tbDocumentsRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDocumentsRepository()
        {
            db = new SaabEntities();
        }

        #endregion
        public string Create(tbDocuments obj)
        {
            try
            {
                db.tbDocuments.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbDocuments Find(int ID)
        {
            return db.tbDocuments.Find(ID);
        }

        public List<tbDocuments> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbDocuments obj)
        {
            try
            {
                var old = Find(obj.ID);
                if (old != null)
                {
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }

        public string Accepting(int doc_ID , bool IsAccepted )
        {
            var ff = Find(doc_ID);
            ff.IsAccepted = IsAccepted;
            return Update(ff);
        }
        public List<tbDocuments> List_userSaver(int saver_id)
        {
            return db.tbDocuments.Where(p => p.FK_UserID_Saver == saver_id).ToList();
        }
    }
}