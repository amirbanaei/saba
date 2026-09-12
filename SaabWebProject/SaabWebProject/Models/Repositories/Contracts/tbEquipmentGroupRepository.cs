using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{


    public class tbEquipmentGroupRepository : IInterFace<tbEquipmentGroup>
    {
        private SaabEntities db;
        public tbEquipmentGroupRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbEquipmentGroup obj)
        {
            if (obj != null)
            {
                db.tbEquipmentGroup.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public bool Disable(int ID)
        {
            if (ID != 0)
            {
                //var res = Find(ID);
                //if (res != null)
                //{
                //    if (res.Status)
                //    {
                //        res.Status = false;
                //    }
                //    else
                //    {
                //        res.Status = true;
                //    }

                //    return Convert.ToBoolean(db.SaveChanges());
                //}
            }

            return false;
        }

        public tbEquipmentGroup Find(int ID)
        {
            return db.tbEquipmentGroup.Find(ID);
        }

        public List<tbEquipmentGroup> Listt()
        {
            return db.tbEquipmentGroup.OrderBy(p => p.Eqpgrp_ID).ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbEquipmentGroup obj)
        {
            if (obj != null)
            {
                db.Entry(obj).State = EntityState.Modified;
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public List<tbEquipmentGroup> GetAllEquipmentGroup()
        {
            return db.tbEquipmentGroup.ToList();
        }

        tbEquipmentGroup IInterFace<tbEquipmentGroup>.Find(int ID)
        {
            throw new NotImplementedException();
        }

        List<tbEquipmentGroup> IInterFace<tbEquipmentGroup>.Update()
        {
            throw new NotImplementedException();
        }
    }
}
