using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbEquipmentRepository : IInterFace<tbEquipments>
    {
        SaabEntities db;
        public tbEquipmentRepository()
        {
            db = new SaabEntities();
        }
        public tbEquipmentRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbEquipments obj)
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbEquipments Find(int ID)
        {
            return db.tbEquipments.Find(ID);
        }

        public List<tbEquipments> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbEquipments obj)
        {
            throw new NotImplementedException();
        }

        public bool AddRange(List<tbEquipments> lstEntry)
        {
            db.tbEquipments.AddRange(lstEntry);
            return Convert.ToBoolean(db.SaveChanges());
        }
     

        public bool Delete(int ID)
        {
            try
            {
                db.tbEquipments.Remove(Find(ID));
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }
          
        }
    }
}