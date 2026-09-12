using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbNergh_PymnRepository
    {
        #region متغیرها
        SaabEntities db;


        public tbNergh_PymnRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbNergh_PymnRepository()
        {
            db = new SaabEntities();
        }
        //----------------------------------------------------------------------------
        public string Create(tbNergh_Pymn obj)
        {
            try
            {
                db.tbNergh_Pymn.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbNergh_Pymn> list_obj)
        {
            try
            {
                db.tbNergh_Pymn.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(tbNergh_Pymn obj)
        {
            throw new NotImplementedException();
        }

        public tbNergh_Pymn Find(int ID)
        {
            //throw new NotImplementedException();
            return db.tbNergh_Pymn.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbNergh_Pymn> Update()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException(); 
        }

        //-------------------------------------------------------------------------------
        public string Update2(tbNergh_Pymn obj)
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