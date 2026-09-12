using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbEquipmentBucnhRepository : IInterFace<tbEquipmentBunch>
    {
        SaabEntities db;
        public tbEquipmentBucnhRepository()
        {
            db = new SaabEntities();
        }
        public tbEquipmentBucnhRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbEquipmentBunch obj)
        {
            if (obj != null)
            {
                db.tbEquipmentBunch.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbEquipmentBunch Find(int ID)
        {
            return db.tbEquipmentBunch.Find(ID); ;
        }
        public List<tbEquipmentBunch> FindByGRP(int ID)
        {
            return db.tbEquipmentBunch.Where(p=>p.FK_EqpgrpID==ID).ToList() ;
        }
        public List<tbEquipmentBunch> Update()
        {
            return db.tbEquipmentBunch.OrderBy(p => p.Eqpbnch_ID).ToList();

        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbEquipmentBunch obj)
        {
            throw new NotImplementedException();
        }
    }
}