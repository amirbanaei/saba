using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbStepTaxRepository : IInterFace<tb_Step_tax>
    {
        SaabEntities db;
        public tbStepTaxRepository()
        {
            db = new SaabEntities();
        }
        public tbStepTaxRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tb_Step_tax obj)
        {
            try
            {
                db.tb_Step_tax.Add(obj);
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

        public tb_Step_tax Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tb_Step_tax obj)
        {
            throw new NotImplementedException();
        }

        public List<tb_Step_tax> Update()
        {
            throw new NotImplementedException();
        }
        public List<tb_Step_tax> List() 
        {
            return db.tb_Step_tax.ToList();
        }
    }
}