using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbMoalefeValueRepository : IInterFace<tbMoalefeValue>
    {
        SaabEntities db;

        public tbMoalefeValueRepository(SaabEntities context)
        {
            db = context;
        }

        public tbMoalefeValueRepository()
        {
            db = new SaabEntities();
        }

        public string Create(tbMoalefeValue obj)
        {
            try
            {
                db.tbMoalefeValue.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception)
            {

                return "false";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbMoalefeValue Find(int ID)
        {
            return db.tbMoalefeValue.Find(ID);
        }

        public List<tbMoalefeValue> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbMoalefeValue obj)
        {
            try
            {
                var old = Find(obj.mlfval_ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
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
        public bool AddRange(List<tbMoalefeValue> lstEntry)
        {
            db.tbMoalefeValue.AddRange(lstEntry);
            return Convert.ToBoolean(db.SaveChanges());
        }
    }
}