using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbVisibleFishRepository : IInterFace<tbVisibleFish>
    {
        SaabEntities db;
        public tbVisibleFishRepository()
        {
            db= new SaabEntities();
        }
        public tbVisibleFishRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbVisibleFish obj)
        {
            try
            {
                db.tbVisibleFish.Add(obj);
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

        public tbVisibleFish Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbVisibleFish obj)
        {
            throw new NotImplementedException();
        }

        public List<tbVisibleFish> Update()
        {
            throw new NotImplementedException();
        }
    }
}