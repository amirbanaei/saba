using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tblink_moalfe_city_valuRepository : IInterFace<tblink_moalfe_city_valu>
    {
        #region متغیرها
        SaabEntities db;


        public tblink_moalfe_city_valuRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tblink_moalfe_city_valuRepository()
        {
            db = new SaabEntities();
        }
        #endregion

        public string Create(tblink_moalfe_city_valu obj)
        {
            try
            {
                db.tblink_moalfe_city_valu.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tblink_moalfe_city_valu> list_obj)
        {
            try
            {
                db.tblink_moalfe_city_valu.AddRange(list_obj);
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

        public tblink_moalfe_city_valu Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tblink_moalfe_city_valu obj)
        {
            throw new NotImplementedException();
        }

        public List<tblink_moalfe_city_valu> Update()
        {
            throw new NotImplementedException();
        }
    }
}