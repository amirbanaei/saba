using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.DomainModels;
namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbCaransSettingRepository : IInterFace<tbCaranSettings>
    {
        private SaabEntities db;
        public tbCaransSettingRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbCaranSettings obj)
        {
            try
            {
                db.tbCaranSettings.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();


            }
            catch(Exception ex)
            {
                return "Error";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbCaranSettings Find(int ID)
        {
            return db.tbCaranSettings.Find(ID);
        }

        public List<tbCaranSettings> Update()
        {
           return db.tbCaranSettings.ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbCaranSettings obj)
        {
            throw new NotImplementedException();
        }
        public bool AddRange(List<tbCaranSettings> lstEntry)
        {
            try
            {
                db.tbCaranSettings.AddRange(lstEntry);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }
    }
}