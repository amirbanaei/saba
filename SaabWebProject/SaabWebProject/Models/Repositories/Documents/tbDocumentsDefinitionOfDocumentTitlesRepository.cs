using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Documents
{
    public class tbDocumentsDefinitionOfDocumentTitlesRepository : IInterFace<tbDocumentsDefinitionOfDocumentTitles>
    {
        #region متغیرها
        SaabEntities db;

        #endregion

        #region سازنده ها
        public tbDocumentsDefinitionOfDocumentTitlesRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDocumentsDefinitionOfDocumentTitlesRepository()
        {
            db = new SaabEntities();
        }

        #endregion
        public string Create(tbDocumentsDefinitionOfDocumentTitles obj)
        {
            throw new NotImplementedException();
        }
        public string Create(List<tbDocumentsDefinitionOfDocumentTitles> obj)
        {
            try
            {
                db.tbDocumentsDefinitionOfDocumentTitles.AddRange(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(tbDocumentsDefinitionOfDocumentTitles obj)
        {
            try
            {
                var old = Find(obj.ID);
                obj.FK_PeymanID = old.FK_PeymanID;
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

        public tbDocumentsDefinitionOfDocumentTitles Find(int ID)
        {
            return db.tbDocumentsDefinitionOfDocumentTitles.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbDocumentsDefinitionOfDocumentTitles> Update()
        {
            return db.tbDocumentsDefinitionOfDocumentTitles.ToList();
        }
        public List<tbDocumentsDefinitionOfDocumentTitles> Listt(List<tbPeymanContracts> contracts)
        {
            List<tbDocumentsDefinitionOfDocumentTitles> list_obj = new List<tbDocumentsDefinitionOfDocumentTitles>();
            foreach (var item in contracts)
            {
                list_obj.AddRange(db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.FK_PeymanID == item.pec_ID).ToList());
            }
            return list_obj;
        }
        public List<tbDocumentsDefinitionOfDocumentTitles> Listt(int PeymanID)
        {
            return db.tbDocumentsDefinitionOfDocumentTitles.Where(p => p.FK_PeymanID == PeymanID).ToList(); 
        }
        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
    }
}