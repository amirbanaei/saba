using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbUserSalaryDetailRepositories : IInterFace<tbUserSalaryDetail>
    {
        SaabEntities db;

        public tbUserSalaryDetailRepositories(SaabEntities _Context)
        {
            db = _Context;
        }
        public tbUserSalaryDetailRepositories()
        {
            db = new SaabEntities();
        }
        public string Create(tbUserSalaryDetail obj)
        {
            try
            {
                db.tbUserSalaryDetail.Add(obj);
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

        public tbUserSalaryDetail Find(int ID)
        {
            return db.tbUserSalaryDetail.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbUserSalaryDetail obj)
        {
            throw new NotImplementedException();
        }

        public List<tbUserSalaryDetail> Update()
        {
            throw new NotImplementedException();
        }
    }
}