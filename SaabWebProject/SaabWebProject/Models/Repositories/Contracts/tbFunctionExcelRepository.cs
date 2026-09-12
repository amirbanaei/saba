using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbFunctionExcelRepository : IInterFace<tbFucntionExcel>
    {
        SaabEntities db;
        public tbFunctionExcelRepository()
        {
            db = new SaabEntities();
        }
        public tbFunctionExcelRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbFucntionExcel obj)
        {
            try
            {
                db.tbFucntionExcel.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
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

        public tbFucntionExcel Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbFucntionExcel> Update()
        {
            return db.tbFucntionExcel.ToList();
        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbFucntionExcel obj)
        {
            throw new NotImplementedException();
        }
    }
}