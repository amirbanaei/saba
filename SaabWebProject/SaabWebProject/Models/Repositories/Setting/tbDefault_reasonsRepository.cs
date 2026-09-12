using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Web;




namespace SaabWebProject.Models.Repositories.Settings
{
    public class tbDefault_reasonsRepository : IInterFace<tbDefault_reasons>
    {

        SaabEntities db;
        public tbDefault_reasonsRepository()//سازنده
        {
            db = new SaabEntities(); // obj database
        }
        public string Create(tbDefault_reasons obj)
        {
            try
            {
                db.tbDefault_reasons.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Create(List<tbDefault_reasons> list_obj)
        {
            try
            {
                db.tbDefault_reasons.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public bool Disable(int ID)
        {
            try
            {
                var find = db.tbDefault_reasons.Find(ID);
                db.tbDefault_reasons.Remove(find);
                return SaveChanges();
            }
            catch (Exception)
            {

                return false;
            }
        }

        public tbDefault_reasons Find(int ID)
        {
            return db.tbDefault_reasons.Find(ID);
        }

        public List<tbDefault_reasons> Update()
        {
            return db.tbDefault_reasons.ToList();
        }

        public bool SaveChanges()
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    db.SaveChanges();
                    transaction.Commit();
                    return true;
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    Console.WriteLine(e.Message);
                    return false;
                }
            }
           
        }


        

        public string Update(List<tbDefault_reasons> obj)
        {
            if (obj != null)
            {
                try
                {
                    foreach (var entity in obj)
                    {
                        db.Entry(entity).State = EntityState.Modified; // Update each entity in the list
                    }

                    SaveChanges(); // Save changes to the database

                    return "Success";
                }
                catch (Exception e)
                {
                    return "False " + e.Message;

                }
            }

            return "False";
        }

        public string Update(tbDefault_reasons obj)
        {
            try
            {
                var old = Find(obj.id);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch (Exception e)
            {
                return "False " + e.Message;

            }
        }
    }
}

     