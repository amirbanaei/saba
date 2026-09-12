using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    //public class tbEquipmentFinantialRepository : IInterFace<tbEquipmentFinantial>
    //{
    //    SaabEntities db;
    //    public tbEquipmentFinantialRepository()
    //    {
    //        db = new SaabEntities();
    //    }
    //    public tbEquipmentFinantialRepository(SaabEntities Context)
    //    {
    //        db = Context;

    //    }
    //    public string Create(tbEquipmentFinantial obj)
    //    {
    //        if (obj != null)
    //        {
    //            db.tbEquipmentFinantial.Add(obj);
    //            return Convert.ToBoolean(db.SaveChanges()).ToString();
    //        }

    //        return false.ToString();
    //    }

    //    public bool Disable(int ID)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public tbEquipmentFinantial Find(int ID)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public List<tbEquipmentFinantial> Listt()
    //    {
    //        return db.tbEquipmentFinantial.ToList();
    //    }

    //    public bool SaveChanges()
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public string Update(tbEquipmentFinantial obj)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}
}