using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbControllyPeymanRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbControllyPeymanRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbControllyPeymanRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public string Create(tbControllyPeyman obj)
        {
            try
            {
                Context.tbControllyPeyman.Add(obj);
                return SaveChanges().ToString();
            }
            catch
            {
                return "False";
            }
        }
        public string Update(tbControllyPeyman obj)
        {
            try
            {
                var old = Find(obj.Id);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    Context.Entry(old).CurrentValues.SetValues(obj);
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
        public tbControllyPeyman Find(int ID)
        {
            try
            {
                return Context.tbControllyPeyman.Find(ID);
            }
            catch
            {
                return null;
            }
        }
        public bool SaveChanges()
        {
            return Convert.ToBoolean(Context.SaveChanges());
        }
        public List<tbControllyPeyman> Listt()
        {
            return Context.tbControllyPeyman.ToList();
        }
    }
}