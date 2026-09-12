using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbMoalefeValueFishTestRepositories : IInterFace<tbMoalefeValeFishTest>
    {
        SaabEntities db;
        public tbMoalefeValueFishTestRepositories()
        {
            db = new SaabEntities();
        }
        public tbMoalefeValueFishTestRepositories(SaabEntities _Context)
        {
            db = _Context;
        }
        public string Create(tbMoalefeValeFishTest obj)
        {
            try
            {
                db.tbMoalefeValeFishTest.Add(obj);
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

        public tbMoalefeValeFishTest Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());

        }

        public string Update(tbMoalefeValeFishTest obj)
        {
            throw new NotImplementedException();
        }

        public List<tbMoalefeValeFishTest> Update()
        {
            throw new NotImplementedException();
        }

        public bool AddRange(List<tbMoalefeValeFishTest> lstEntry)
        {
            try
            {
                db.tbMoalefeValeFishTest.AddRange(lstEntry);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }

    }
}