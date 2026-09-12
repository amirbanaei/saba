using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tblink_moalfe_cityRepository : IInterFace<tblink_moalfe_city>
    {
        SaabEntities db;


        public tblink_moalfe_cityRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tblink_moalfe_cityRepository()
        {
            db = new SaabEntities();
        }
        public string Create(tblink_moalfe_city obj)
        {
            try
            {
                db.tblink_moalfe_city.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tblink_moalfe_city> list_obj)
        {
            try
            {
                db.tblink_moalfe_city.AddRange(list_obj);
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

        public tblink_moalfe_city Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tblink_moalfe_city obj)
        {
            throw new NotImplementedException();
        }

        public List<tblink_moalfe_city> Update()
        {
            throw new NotImplementedException();
        }
    }
}