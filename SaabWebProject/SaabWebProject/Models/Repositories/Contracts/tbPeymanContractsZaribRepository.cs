using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbPeymanContractsZaribRepository : IInterFace<tbPeymanContractsZarib>
    {
        private SaabEntities db;
        public tbPeymanContractsZaribRepository()
        {
            db = new SaabEntities();
        }
        public tbPeymanContractsZaribRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbPeymanContractsZarib obj)
        {
            try
            {
                db.tbPeymanContractsZarib.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();


            }
            catch (Exception ex)
            {
                return "Error";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbPeymanContractsZarib Find(int ID)
        {
            return db.tbPeymanContractsZarib.Find(ID);
        }

        public List<tbPeymanContractsZarib> Update()
        {
            return db.tbPeymanContractsZarib.ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbPeymanContractsZarib obj)
        {
            if (obj != null)
            {
                db.Entry(obj).State = EntityState.Modified;
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }
        public bool AddRange(List<tbPeymanContractsZarib> lstEntry)
        {
            try
            {
                db.tbPeymanContractsZarib.AddRange(lstEntry);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }
    }
}