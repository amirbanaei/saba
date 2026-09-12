using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbRial_restrictions_pymnRepository
    {
        #region متغیرها
        SaabEntities db;


        public tbRial_restrictions_pymnRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbRial_restrictions_pymnRepository()
        {
            db = new SaabEntities();
        }
        //----------------------------------------------------------------------------
        public string Create(tbRial_restrictions_pymn obj)
        {
            try
            {
                db.tbRial_restrictions_pymn.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbRial_restrictions_pymn> list_obj)
        {
            try
            {
                db.tbRial_restrictions_pymn.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(tbRial_restrictions_pymn obj)
        {
            throw new NotImplementedException();
        }

        public tbRial_restrictions_pymn Find(int ID)
        {
            //throw new NotImplementedException();
            return db.tbRial_restrictions_pymn.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbRial_restrictions_pymn> Update()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        //-------------------------------------------------------------------------------
        public string Update2(tbRial_restrictions_pymn obj)
        {
            try
            {
                var old = Find(obj.ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        #endregion
    }
}