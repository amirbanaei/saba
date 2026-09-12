using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbmoalfeExcelRepository: IInterFace<tbmoalfefishexcel>
    {
        #region متغیرها
        SaabEntities db;


        public tbmoalfeExcelRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbmoalfeExcelRepository()
        {
            db = new SaabEntities();
        }

        public string Create(tbmoalfefishexcel obj)
        {
            try
            {
                db.tbmoalfefishexcel.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbmoalfefishexcel> list_obj)
        {
            try
            {
                db.tbmoalfefishexcel.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(tbmoalfefishexcel obj)
        {
            throw new NotImplementedException();
        }

        public tbmoalfefishexcel Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbmoalfefishexcel> Update()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
        #endregion

    }
}