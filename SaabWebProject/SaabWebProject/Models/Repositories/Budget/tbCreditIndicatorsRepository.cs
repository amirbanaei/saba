using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Budget
{
    public class tbCreditIndicatorsRepository : IInterFace<tbCreditIndicators>
    {
        SaabEntities db;
        public tbCreditIndicatorsRepository()
        {
            db = new SaabEntities();
        }
        public tbCreditIndicatorsRepository(SaabEntities Context)
        {
            db = Context;
        }
        public string Create(tbCreditIndicators obj)
        {
            if (obj != null)
            {
                db.tbCreditIndicators.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbCreditIndicators Find(int ID)
        {

            return db.tbCreditIndicators.Find(ID);
        }

        public List<tbCreditIndicators> Update()
        {
            return db.tbCreditIndicators.ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbCreditIndicators obj)
        {
            throw new NotImplementedException();
        }
        public List<tbCreditIndicators> List()
        {
            return db.tbCreditIndicators.ToList();
        }

        public string Update(List<tbCreditIndicators> newRecords)
        {
            try
            {

                foreach (var record in newRecords)
                {
                    var old = db.tbCreditIndicators.Find(record.ID);

                    if (old != null)
                    {

                        db.Entry(old).CurrentValues.SetValues(record);
                    }
                    else
                    {

                        db.tbCreditIndicators.Add(record);
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
    }
}