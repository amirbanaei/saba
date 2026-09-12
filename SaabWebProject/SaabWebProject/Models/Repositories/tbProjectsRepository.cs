using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories
{
    public class tbProjectsRepository : IInterFace<tbProjects> ,IDisposable
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbProjectsRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbProjectsRepository()
        {
            Context = new SaabEntities();
        }
        #endregion

        public string Create(tbProjects obj)
        {
            try
            {
                //var findold = Context.tbContractJobGroup.Where(p => p.jg_Year == obj.jg_Year && p.jg_Code == obj.jg_Code).FirstOrDefault();
                //if (findold != null)
                //    return "False_Availabel";
                //Context.tbContractJobGroup.Add(obj);
                //return SaveChanges().ToString();
                return "False";
            }
            catch
            {
                return "False";
            }
        }
        public string Update(tbProjects obj)
        {
            try
            {
                //var findold = Context.tbContractJobGroup.Where(p => p.jg_Year == obj.jg_Year && p.jg_Code == obj.jg_Code && p.jg_ID != obj.jg_ID).FirstOrDefault();
                //if (findold != null)
                //    return "False_Availabel";
                //var old = Find(obj.jg_ID);
                //if (old != null)
                //{
                //    old.jg_Title = obj.jg_Title;
                //    old.jg_Year = obj.jg_Year;
                //    old.jg_Code = obj.jg_Code;
                //    foreach (var item in old.tbRelMoalefeJobGroup)
                //    {
                //        item.rmg_Value = obj.tbRelMoalefeJobGroup.Where(p => p.rmg_ID == item.rmg_ID).FirstOrDefault().rmg_Value;
                //    }
                //    return SaveChanges().ToString();
                //}
                //else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }
        public tbProjects Find(int ID)
        {
            try
            {
                return Context.tbProjects.Find(ID);
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
        public List<tbProjects> Update()
        {
            return Context.tbProjects.ToList();
        }
    

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
    }
}