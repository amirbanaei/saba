using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbMoalefeValueFishRepositories : IInterFace<tbMoalefeValueFish>
    {
        SaabEntities db;
        public tbMoalefeValueFishRepositories()
        {
            db = new SaabEntities();
        }
        public tbMoalefeValueFishRepositories(SaabEntities _Context)
        {
            db = _Context;
        }
        public string Create(tbMoalefeValueFish obj)
        {
            try
            {
                db.tbMoalefeValueFish.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbMoalefeValueFish Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());

        }

        public string Update(tbMoalefeValueFish obj)
        {
            throw new NotImplementedException();
        }

        public List<tbMoalefeValueFish> Update()
        {
            throw new NotImplementedException();
        }

        public bool AddRange(List<tbMoalefeValueFish> lstEntry)
        {
            try
            {
                db.tbMoalefeValueFish.AddRange(lstEntry);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }

    }
}