using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbsaveSoratBastehRepository : IInterFace<tbsaveSoratBasteh>
    {
        SaabEntities db;
        public tbsaveSoratBastehRepository()
        {
            db = new SaabEntities();
        }
        public tbsaveSoratBastehRepository(SaabEntities Context)
        {
            db = Context;
        }


        public string Create(tbsaveSoratBasteh obj)
        {
            try
            {
                db.tbsaveSoratBasteh.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception e)
            {

                return "false";
            }
        }

      

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbsaveSoratBasteh Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return System.Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbsaveSoratBasteh obj)
        {
            throw new NotImplementedException();
        }

        public List<tbsaveSoratBasteh> Update()
        {
            throw new NotImplementedException();
        }
    }
}