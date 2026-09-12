using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbEquipmentSpecificationsRepository : Interfaces.IInterFace<tbEquipmentSpecifications>
    {
        private SaabEntities db;
        public tbEquipmentSpecificationsRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbEquipmentSpecifications obj)
        {
            try
            {

                db.tbEquipmentSpecifications.Add(obj);
                return "True";

            }
            catch (Exception ex)
            {
                return "Error";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbEquipmentSpecifications Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbEquipmentSpecifications> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbEquipmentSpecifications obj)
        {
            throw new NotImplementedException();
        }

        public List<tbEquipmentSpecifications> GetSpecGroupAndBunch(int FK_EqpBunch_ID)
        {
            return db.tbEquipmentSpecifications.Where(p => p.FK_EqpBunch_ID == FK_EqpBunch_ID).ToList();
        }
    }
}