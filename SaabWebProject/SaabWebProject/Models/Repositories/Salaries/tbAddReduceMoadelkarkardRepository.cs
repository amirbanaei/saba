using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbAddReduceMoadelkarkardRepository : IInterFace<tbAddReduceMoadelkarkard>
    {
        SaabEntities db;
        public tbAddReduceMoadelkarkardRepository()
        {
            db = new SaabEntities();
        }
        public tbAddReduceMoadelkarkardRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbAddReduceMoadelkarkard obj)
        {
            try
            {
                db.tbAddReduceMoadelkarkard.Add(obj);
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

        public tbAddReduceMoadelkarkard Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbAddReduceMoadelkarkard obj)
        {
            throw new NotImplementedException();
        }

        public List<tbAddReduceMoadelkarkard> Update()
        {
            throw new NotImplementedException();
        }
    }
}