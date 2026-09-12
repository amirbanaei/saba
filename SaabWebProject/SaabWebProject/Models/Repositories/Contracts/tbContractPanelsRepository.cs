using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbContractPanelsRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbContractPanelsRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbContractPanelsRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public string Create(tbContractPanels obj)
        {
            try
            {
                Context.tbContractPanels.Add(obj);
                return SaveChanges().ToString();
            }
            catch
            {
                return "False";
            }
        }
        public string Update(tbContractPanels obj)
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
        public tbContractPanels Find(int ID)
        {
            try
            {
                return Context.tbContractPanels.Find(ID);
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
        public List<tbContractPanels> Listt()
        {
            return Context.tbContractPanels.ToList();
        }

    }
}
