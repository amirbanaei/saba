using System;
using System.Collections.Generic;
using System.Linq;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
namespace SaabWebProject.Models.Repositories
{
    public class tbPeymanContractsRepository : IInterFace<tbPeymanContracts>
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbPeymanContractsRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbPeymanContractsRepository()
        {
            Context = new SaabEntities();
        }
        #endregion
        public string Create(tbPeymanContracts obj)
        {

            try
            {
                Context.tbPeymanContracts.Add(obj);
                if (SaveChanges() == true)
                {
                    return obj.pec_ID.ToString();
                }
                else
                {
                    return "0";
                }
            }
            catch
            {
                return "0";
            }
        }
        public string Update(tbPeymanContracts obj)
        {
            try
            {
                var old = Find(obj.pec_ID);
                if (old != null)
                {
                    old.pec_Price = obj.pec_Price;
                    foreach (var item in (obj.tbPeymanContractPrice.ToList()))
                    {
                        item.FKPeymanID = old.pec_ID;
                    }
                    
                    Context.tbPeymanContractPrice.AddRange(obj.tbPeymanContractPrice.ToList());
                   
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
        public tbPeymanContracts Find(int ID)
        {
            try
            {
                return Context.tbPeymanContracts.Find(ID);
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
        public List<tbPeymanContracts> Update()
        {
            try { return Context.tbPeymanContracts.Where(p => p.Inactive != true).ToList(); }
            catch(Exception ex)
            {
                return new List<tbPeymanContracts>();
            }
         
        }
        public List<tbPeymanContracts> Listt(int peymanID)
        {
            return Context.tbPeymanContracts.Where(p => p.pec_ID == peymanID || p.FK_PecFatherElhaghie == peymanID && p.Inactive != true).ToList();
        }
        public bool Disable(int ID) //deleted
        {

            var find = Find(ID);
            if(find != null)
            {
                var a = Context.tbPeymanContractPrice.Where(p => p.FKPeymanID == ID).ToList();
                foreach(var item in a)
                {
                    Context.tbPeymanContractPrice.Remove(item);
                }

                var b = Context.tbPeymanElhaghie.Where(p => p.FK_PeymanID == ID).ToList();
                foreach (var item in b)
                {
                    Context.tbPeymanElhaghie.Remove(item);
                }

                Context.tbPeymanContracts.Remove(find);
                return SaveChanges();

            }
            return false ;
        }
    }
}