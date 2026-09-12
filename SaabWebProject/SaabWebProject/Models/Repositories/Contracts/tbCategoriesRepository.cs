using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbCategoriesRepository : IInterFace<tbCategories>
    {
        private SaabEntities db;
        public tbCategoriesRepository(SaabEntities _db)
        {
            db = _db;
        }
        public string Create(tbCategories obj)
        {
            if (obj != null)
            {
                db.tbCategories.Add(obj);
                return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public bool Disable(int ID)
        {
            if (ID != 0)
            {
                var res = Find(ID);
                if (res != null)
                {
                    if (res.Status)
                    {
                        res.Status = false;
                    }
                    else
                    {
                        res.Status = true;
                    }

                    return Convert.ToBoolean(db.SaveChanges());
                }
            }

            return false;
        }

        public tbCategories Find(int ID)
        {
            return db.tbCategories.Find(ID);
        }

        public List<tbCategories> Update()
        {
            return db.tbCategories.OrderBy(p=> p.Category_ID).ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbCategories obj)
        {
            if (obj != null)
            {
                db.Entry(obj).State = EntityState.Modified;
               return Convert.ToBoolean(db.SaveChanges()).ToString();
            }

            return false.ToString();
        }

        public List<tbCategories> GetAllActiveCategory()
        {
            return db.tbCategories.Where(p => p.Status == true).ToList();
        }
    }
}