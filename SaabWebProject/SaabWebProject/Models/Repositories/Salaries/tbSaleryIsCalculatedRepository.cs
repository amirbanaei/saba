using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbSaleryIsCalculatedRepository : IInterFace<tbSaleryIsCalculated>
    {
        SaabEntities db;
        public tbSaleryIsCalculatedRepository()
        {
            db = new SaabEntities();
        }
        public tbSaleryIsCalculatedRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbSaleryIsCalculated obj)
        {
            try
            {
                db.tbSaleryIsCalculated.Add(obj);
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

        public tbSaleryIsCalculated Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbSaleryIsCalculated obj)
        {
            throw new NotImplementedException();
        }

        public List<tbSaleryIsCalculated> Update()
        {
            throw new NotImplementedException();
        }
    }
}