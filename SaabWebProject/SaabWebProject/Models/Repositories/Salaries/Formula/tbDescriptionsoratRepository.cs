using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries.Formula
{
    public class tbDescriptionsoratRepository : IInterFace<Descriptionsorat>
    {
        SaabEntities db;


        public tbDescriptionsoratRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbDescriptionsoratRepository()
        {
            db = new SaabEntities();
        }

        public int Create(Descriptionsorat obj)
        {
            try
            {
                db.Descriptionsorat.Add(obj);
                 SaveChanges();
                return obj.ID;
            }
            catch (Exception ex)
            {

                return 1;

            }
        }
        public string Create3(TbDescriptionDetail obj)
        {
            try
            {
                db.TbDescriptionDetail.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<Descriptionsorat> list_obj)
        {
            try
            {
                db.Descriptionsorat.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public string Create3(List<TbDescriptionDetail> list_obj)
        {
            try
            {
                db.TbDescriptionDetail.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Update(Descriptionsorat obj)
        {
            throw new NotImplementedException();
        }

        public Descriptionsorat Find(int ID)
        {
            try
            {
                return db.Descriptionsorat.Find(ID);


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding FinancialDocument with ID {ID}: {ex.Message}");
                // You might want to throw or return a default value here, depending on your requirements
                throw;
            }
        }

        public bool SaveChanges()
        {
            try
            {
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public List<Descriptionsorat> Update()
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
        public string Update2(Descriptionsorat obj)
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

        string IInterFace<Descriptionsorat>.Create(Descriptionsorat obj)
        {
            throw new NotImplementedException();
        }
    }
}