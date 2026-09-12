using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories
{
    public class tbUserRentContractsRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbUserRentContractsRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbUserRentContractsRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public bool Create(tbUserRentContracts obj)
        {
            try
            {
                Context.tbUserRentContracts.Add(obj);
                return SaveChanges();
            }
            catch
            {
                return false;
            }
        }
        public bool Update(tbUserRentContracts obj)
        {
            try
            {
                var old = Find(obj.urc_ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    Context.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges();
                }
                else
                    return true;
            }
            catch
            {
                return false;
            }
        }
        public tbUserRentContracts Find(int ID)
        {
            try
            {
                return Context.tbUserRentContracts.Find(ID);
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
        public List<tbUserRentContracts> Listt()
        {
            return Context.tbUserRentContracts.ToList();
        }
        public List<tbUserRentContracts> Listt(int user_ID)
        {
            return Context.tbUserRentContracts.Where(p => p.FK_UserID == user_ID).ToList();
        }

        public string Disable(int ID)
        {
            var old = Find(ID);
            if (old != null)
            {
                if (old.urc_Accepted != true)
                {
                    Context.Entry(old).State = EntityState.Deleted;
                    return SaveChanges().ToString();
                }
                else
                {
                    return "false_Accepted";
                }
            }
            else
            {
                return "true_notFound";
            }
        }
        public bool AcceptUserContract(int userContractID)
        {
            var old = Find(userContractID);
            try
            {
                old.urc_Accepted = true;
                return SaveChanges();
            }
            catch (Exception)
            {

                return false;
            }
        }
    }
}