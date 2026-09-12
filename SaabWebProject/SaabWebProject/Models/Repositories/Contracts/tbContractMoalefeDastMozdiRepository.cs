using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;


namespace SaabWebProject.Models.Repositories
{
    public class tbContractMoalefeDastMozdiRepository : IInterFace<tbContractMoalefeDastmozdi>
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbContractMoalefeDastMozdiRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbContractMoalefeDastMozdiRepository()
        {
            Context = new SaabEntities();
        }
        #endregion
       
        public string Create(tbContractMoalefeDastmozdi obj)
        {
            try
            {
                Context.tbContractMoalefeDastmozdi.Add(obj);
                return SaveChanges().ToString();
            }
            catch
            {
                return "False";
            }
        }
        public string Update(tbContractMoalefeDastmozdi obj)
        {
            try
            {
                var old = Find(obj.md_ID);
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
        public tbContractMoalefeDastmozdi Find(int ID)
        {
            try
            {
                return Context.tbContractMoalefeDastmozdi.Find(ID);
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
        public List<tbContractMoalefeDastmozdi> Update()
        {
            return Context.tbContractMoalefeDastmozdi.ToList(); 
        }
        /// <summary>
        /// لیست مولفه های فعال
        /// </summary>
        /// <returns></returns>
        public List<tbContractMoalefeDastmozdi> ActivateList()
        {
            return Context.tbContractMoalefeDastmozdi.Where(p => p.md_IsActive == true && p.md_Type==1).ToList();
        }
        public bool Disable(int ID)
        {
            var old = Find(ID);
            if (old != null)
            {
                if ((bool)old.md_IsActive)
                {
                    old.md_IsActive = false;
                }
                else
                {
                    old.md_IsActive = true;
                }
                var result = SaveChanges();
                return result;
            }
            else
            {
                return true;
            }
        }


        public bool AddRange(List<tbContractMoalefeDastmozdi> lstEntry)
        {
            try
            {
                Context.tbContractMoalefeDastmozdi.AddRange(lstEntry);
                return Convert.ToBoolean(Context.SaveChanges());
            }
            catch
            {
                return false;
            }
        }

        public List<tbContractMoalefeDastmozdi> MoalefeTypeDateAndTime_Lists()
        {
            return Context.tbContractMoalefeDastmozdi.Where(p => p.md_variable == "date" || p.md_variable == "time").ToList();
        }


    }
}