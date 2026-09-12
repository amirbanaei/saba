using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Budget
{
    public class tbDetermining_creditlineRepository : IInterFace<tbDetermining_creditline>
    {
        SaabEntities db;


        public tbDetermining_creditlineRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDetermining_creditlineRepository()
        {
            db = new SaabEntities();
        }


        public string Create(tbDetermining_creditline obj)
        {
            try
            {
                db.tbDetermining_creditline.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }


        public string Create(List<tbDetermining_creditline> list_obj)
        {
            try
            {
                db.tbDetermining_creditline.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }




        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbDetermining_creditline Find(int ID)
        {
            return db.tbDetermining_creditline.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbDetermining_creditline obj)
        {
            try
            {
                var old = db.tbDetermining_creditline.Find(obj.Determining_creditline_ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }


        public string Update(List<tbDetermining_creditline> newRecords)
        {
            try
            {
           
                foreach (var record in newRecords)
                {
                    var old = db.tbDetermining_creditline.Find(record.Determining_creditline_ID);

                    if (old != null)
                    {
                     
                        db.Entry(old).CurrentValues.SetValues(record);
                    }
                    else
                    {
                       
                        db.tbDetermining_creditline.Add(record);
                    }
                }

               
                SaveChanges();

                return "True"; // اگر همه چیز موفقیت‌آمیز بود
            }
            catch (Exception ex)
            {
                // در صورت بروز خطا، اطلاعات خطا را به عنوان نتیجه برگردانید
                return ex.Message;
            }
        }


        public List<tbDetermining_creditline> List()
        {
            return db.tbDetermining_creditline.ToList();
        }
        public bool Delete(int ID)
        {
            try
            {
                db.tbDetermining_creditline.Remove(Find(ID));
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }

        }

        public List<tbDetermining_creditline> Update()
        {
            throw new NotImplementedException();
        }
    }
}