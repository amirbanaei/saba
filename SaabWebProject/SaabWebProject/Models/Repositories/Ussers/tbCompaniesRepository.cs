using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Ussers
{
    public class tbCompaniesRepository : IInterFace<tbCompanies>
    {
        SaabEntities db = new SaabEntities();

        public string Create(tbCompanies obj)
        {
            try
            {
                db.tbCompanies.Add(obj);
                return SaveChanges().ToString();
            }
            catch(Exception ex)
            {
                return "False";
            }
        }

        public tbCompanies Find(int ID)
        {
            return db.tbCompanies.Find(ID);
        }

        public List<tbCompanies> Update()
        {
            return db.tbCompanies.ToList();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbCompanies obj)
        {
            try
            {
                var old = Find(obj.ID);
                if (old != null)
                {
                    obj.tbUsers.usr_IsActive = old.tbUsers.usr_IsActive;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    //db.Entry(old.tbUsers).CurrentValues.SetValues(obj.tbUsers);

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

        public bool ActiveOrNotActiveCompany(int id)
        {
            var res = db.tbCompanies.Find(id);
            if (res != null && res.IsActive == true)
            {
                res.IsActive = false;
            }
            else
            {
                if (res != null) res.IsActive = true;
            }

            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbCompanies> ActiveList()
        {
            return db.tbCompanies.Where(p => p.IsActive == true).ToList();
        }
    }
}