using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Documents
{
    public class tbDocumentsAuthorizedUsersRepository : IInterFace<tbDocumentsAuthorizedUsers>
    {
        #region متغیرها
        SaabEntities db;

        #endregion

        #region سازنده ها
        public tbDocumentsAuthorizedUsersRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDocumentsAuthorizedUsersRepository()
        {
            db = new SaabEntities();
        }
        #endregion
        public string Create(tbDocumentsAuthorizedUsers obj)
        {
            try
            {
                db.tbDocumentsAuthorizedUsers.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public string Create(List<tbDocumentsAuthorizedUsers> obj)
        {
            try
            {
             
                db.tbDocumentsAuthorizedUsers.AddRange(obj);
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

        public tbDocumentsAuthorizedUsers Find(int ID)
        {
            return db.tbDocumentsAuthorizedUsers.Find(ID);
        }

        public List<tbDocumentsAuthorizedUsers> Update()
        {
            var result =  db.tbDocumentsAuthorizedUsers.OrderByDescending(p => p.FK_PeymanID).ToList();
            foreach(var item in result)
            {
                item.tbPeymanContracts = db.tbPeymanContracts.Find(item.FK_PeymanID);
            }
            return result;
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbDocumentsAuthorizedUsers obj)
        {
            return "false";
        }
        public string Update(List<tbDocumentsAuthorizedUsers> obj)
        {
            try
            {
                using (DbContextTransaction transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        var ff = obj.FirstOrDefault().FK_PeymanID;
                        var olds = db.tbDocumentsAuthorizedUsers.Where(p => p.FK_PeymanID == ff).ToList();
                        if(olds != null)
                             db.tbDocumentsAuthorizedUsers.RemoveRange(olds);
                        db.tbDocumentsAuthorizedUsers.AddRange(obj);
                        db.SaveChanges();

                        transaction.Commit();
                        return "True";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return ex.Message;
                    }
                }

            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public List<tbPeymanContracts> GePeymansOfThisUSer(int UserID)
        {
            List<tbPeymanContracts> listPeyman = new List<tbPeymanContracts>();
            foreach(var item in db.tbDocumentsAuthorizedUsers.Where(p => p.FK_UserID == UserID).ToList())
            {
                listPeyman.Add(item.tbPeymanContracts);
            }
            listPeyman = listPeyman.Distinct().ToList();
            return listPeyman;
        }
        public List<tbUsers> GetUserInThisPeymans(int PeymanID)
        {
            return db.tbDocumentsAuthorizedUsers.Where(p => p.FK_PeymanID == PeymanID).Select(p => p.tbUsers).Distinct().ToList();
        }

    }
}