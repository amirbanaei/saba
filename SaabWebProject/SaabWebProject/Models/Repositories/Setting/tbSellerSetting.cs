using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using SaabWebProject.Models.DomainModels;
using System.Data.Entity.Validation;

namespace SaabWebProject.Models.Repositories.Setting
{
    public class SellerSettingRepository
    {
        SaabEntities db = new SaabEntities();
        public bool Create(tb_SellerSetting sellerSetting)
        {
            try
            {
                db.tb_SellerSetting.Add(sellerSetting);
                db.SaveChanges();
                return true;
            }
            catch (DbEntityValidationException e)
            {
                foreach (var eve in e.EntityValidationErrors)
                {
                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                    foreach (var ve in eve.ValidationErrors)
                    {
                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                            ve.PropertyName, ve.ErrorMessage);
                    }
                }
                throw;
            }
        }
        public tb_SellerSetting Find(int id)
        {
            return db.tb_SellerSetting.Find(id);
        }
        public List<tb_SellerSetting> List()
        {
            return db.tb_SellerSetting.ToList();
        }

        public bool Update(tb_SellerSetting entity)
        {

            if (entity != null)
            {
                try
                {
                    db.Entry(entity).State = EntityState.Modified;
                    return Convert.ToBoolean(db.SaveChanges());
                }
                catch
                {

                    return false;
                }
            }
            else
            {
                return true;
            }

        }
    }
}