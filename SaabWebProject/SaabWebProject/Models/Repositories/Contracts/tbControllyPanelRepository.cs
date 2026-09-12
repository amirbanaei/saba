using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbControllyPanelRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbControllyPanelRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbControllyPanelRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public string Create(tbControllyPanel obj)
        {
            try
            {
                Context.tbControllyPanel.Add(obj);
                return SaveChanges().ToString();
            }
            catch
            {
                return "False";
            }
        }
        public string Update(tbControllyPanel obj)
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
        public tbControllyPanel Find(int ID)
        {
            try
            {
                return Context.tbControllyPanel.Find(ID);
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
        public List<tbControllyPanel> Listt()
        {
            return Context.tbControllyPanel.ToList();
        }

    }

}