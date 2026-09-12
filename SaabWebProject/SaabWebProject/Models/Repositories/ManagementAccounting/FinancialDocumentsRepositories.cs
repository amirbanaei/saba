using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.ManagementAccounting
{
    public class FinancialDocumentsRepositories : IInterFace<FinancialDocuments>
    {
        #region متغیرها
        SaabEntities db;


        public FinancialDocumentsRepositories(SaabEntities Context)
        {
            db = Context;
        }
        public FinancialDocumentsRepositories()
        {
            db = new SaabEntities();
        }

        public string Create(FinancialDocuments obj)
        {

            try
            {
                db.FinancialDocuments.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<FinancialDocuments> list_obj)
        {
            try
            {
                db.FinancialDocuments.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }



        public string Update(FinancialDocuments obj)
        {
            try
            {
                var formul = db.FinancialDocuments.Where(p => p.ID == obj.ID).FirstOrDefault();
                if (formul != null)
                {

                    formul.Title = obj.Title;
                    formul.Creditor = obj.Creditor;
                    formul.Creditlamdicatros_ID = obj.Creditlamdicatros_ID;
                    formul.Debtore = obj.Debtore;

                    var res = SaveChanges();
                    return res.ToString();
                }
                else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }

        public FinancialDocuments Find(int ID)
        {
            try
            {
                return db.FinancialDocuments.Find(ID);


            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error finding FinancialDocument with ID {ID}: {ex.Message}");
                // You might want to throw or return a default value here, depending on your requirements
                throw;
            }
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }



        public List<FinancialDocuments> Update()
        {
            return db.FinancialDocuments.ToList();
        }


        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }



        public void Remove(FinancialDocuments FinancialDocuments)
        {
            db.FinancialDocuments.Remove(FinancialDocuments);
        }



        public string Create2(tbDetailedtitle obj)
        {

            try
            {
                db.tbDetailedtitle.Add(obj);
                return SaveChanges().ToString();

            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create2(List<tbDetailedtitle> list_obj)
        {
            try
            {
                db.tbDetailedtitle.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public List<tbDetailedtitle> List()
        {
            return db.tbDetailedtitle.ToList();
        }
        public string Update2(FinancialDocuments obj)
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
            catch(Exception ex)
            {
                return "False";
            }
        }
        #endregion




    }
}