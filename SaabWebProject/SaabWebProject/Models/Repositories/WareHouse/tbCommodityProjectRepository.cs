using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.Repositories.WareHouse
{
    public class tbCommodityProjectRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbCommodityProjectRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbCommodityProjectRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public string Create(tbCommodityProject obj)
        {
            try
            {
                Context.tbCommodityProject.Add(obj);
                return SaveChanges().ToString();
            }
            catch
            {
                return "False";
            }
        }
        public string Create(List<tbCommodityProject> list_obj)
        {
            try
            {
                Context.tbCommodityProject.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }




        public string Update(tbCommodityProject obj)
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
            catch(Exception ex)
            {
                return "False";
            }
        }
        public tbCommodityProject Find(int ID)
        {
            try
            {
                return Context.tbCommodityProject.Find(ID);
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
        public List<tbCommodityProject> Listt()
        {
            return Context.tbCommodityProject.ToList();
        }
    }
}