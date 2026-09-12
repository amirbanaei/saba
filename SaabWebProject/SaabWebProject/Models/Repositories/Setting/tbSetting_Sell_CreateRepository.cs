using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Setting
{
    public class tbSetting_Sell_CreateRepository : IInterFace<tbSetting_Sell_Create>
    {
        #region متغیرها
        SaabEntities db;


        public tbSetting_Sell_CreateRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbSetting_Sell_CreateRepository()
        {
            db = new SaabEntities();
        }
        //----------------------------------------------------------------------------
        public string Create(tbSetting_Sell_Create obj)
        {
            try
            {
                db.tbSetting_Sell_Create.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbSetting_Sell_Create> list_obj)
        {
            try
            {
                db.tbSetting_Sell_Create.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(tbSetting_Sell_Create obj)
        {
            throw new NotImplementedException();
        }

        public tbSetting_Sell_Create Find(int ID)
        {
            //throw new NotImplementedException();
            return db.tbSetting_Sell_Create.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public List<tbSetting_Sell_Create> Update()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        //-------------------------------------------------------------------------------
        public string Update2(tbSetting_Sell_Create obj)
        {
            try
            {
                var old = Find(obj.ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        #endregion
    }
}