using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.BaseInformation
{
    public class tbUnitParameterRepository : IInterFace<tbUnitParameter>
    {
        SaabEntities db;
        public tbUnitParameterRepository()
        {
            db = new SaabEntities();
        }
        public tbUnitParameterRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbUnitParameter obj)
        {
            if (obj != null)
            {
                db.tbUnitParameter.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public bool Disable(int ID)
        {
            if (ID != 0)
            {
                var res = Find(ID);
                if (res != null)
                {
                    if (res.Status==true)
                    {
                        res.Status = false;
                    }
                    else
                    {
                        res.Status = true;
                    }

                    return Convert.ToBoolean(db.SaveChanges());
                }
            }

            return false;
        }

        public tbUnitParameter Find(int ID)
        {
            return db.tbUnitParameter.Find(ID);
        }

        public List<tbUnitParameter> Update()
        {
            return db.tbUnitParameter.ToList();
        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbUnitParameter obj)
        {
            if (obj != null)
            {
                db.Entry(obj).State = EntityState.Modified;
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        } 
        public List<tbUnitParameter> ListActiveUnit()
        {
            return db.tbUnitParameter.Where(p => p.Status != false).ToList();
        }
    }
}