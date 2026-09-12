using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbMoalefeValueSoratRepositories : IInterFace<tbMoalefeValueSorat>
    {
        SaabEntities db;
        public tbMoalefeValueSoratRepositories()
        {
            db = new SaabEntities();
        }
        public tbMoalefeValueSoratRepositories(SaabEntities _Context)
        {
            db = _Context;
        }

        
     
        public string Create(tbMoalefeValueSorat obj)
        {
            try
            {
                db.tbMoalefeValueSorat.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }

        public bool AddRange(List<tbMoalefeValueSorat> lstEntry)
        {
            try
            {
                db.tbMoalefeValueSorat.AddRange(lstEntry);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }
        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbMoalefeValueSorat Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbMoalefeValueSorat obj)
        {
            throw new NotImplementedException();
        }

        public List<tbMoalefeValueSorat> Update()
        {
            throw new NotImplementedException();
        }
    }
}