using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbHighLowMoalefeValRepository : IInterFace<tbHighLowMoalefeVal>
    {
        SaabEntities db;
        public tbHighLowMoalefeValRepository()
        {
            db= new SaabEntities();
        }
        public tbHighLowMoalefeValRepository(SaabEntities Context)
        {
            db= Context;
        }

        public string Create(tbHighLowMoalefeVal obj)
        {
            try
            {
                db.tbHighLowMoalefeVal.Add(obj);
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

        public tbHighLowMoalefeVal Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbHighLowMoalefeVal obj)
        {
            throw new NotImplementedException();
        }

        public List<tbHighLowMoalefeVal> Update()
        {
            throw new NotImplementedException();
        }
    }
}