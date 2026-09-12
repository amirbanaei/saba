using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;
namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbEquipmentSpecificationDataRepository : IInterFace<tbEquipmentSpecificationData>
    {
        private SaabEntities db;
        public tbEquipmentSpecificationDataRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbEquipmentSpecificationData obj)
        {
            try
            {
                db.tbEquipmentSpecificationData.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }
            catch
            {
                return "False";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbEquipmentSpecificationData Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbEquipmentSpecificationData> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbEquipmentSpecificationData obj)
        {
            throw new NotImplementedException();
        }
    }
}