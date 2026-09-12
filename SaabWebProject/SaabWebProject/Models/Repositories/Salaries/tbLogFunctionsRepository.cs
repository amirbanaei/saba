using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbLogFunctionsRepository : IInterFace<tbLogFunctions>
    {
        SaabEntities db;
        public tbLogFunctionsRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbLogFunctionsRepository()
        {
            db = new SaabEntities();
        }
        public string Create(tbLogFunctions obj)
        {
            try
            {
                db.tbLogFunctions.Add(obj);
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

        public tbLogFunctions Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbLogFunctions> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbLogFunctions obj)
        {
            throw new NotImplementedException();
        }
    }
}