using EntityFramework.BulkInsert.Extensions;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{

    public class tbSavedFunctionsRepositories : IInterFace<tbSavedFunctions>
    {
        SaabEntities db;
        public tbSavedFunctionsRepositories()
        {
            db = new SaabEntities();
        }
        public tbSavedFunctionsRepositories(SaabEntities Context)
        {
            db = Context;
        }

        public string Create4(List<tbSavedFunctions> list_obj)
        {
            try
            {


                // Assuming 'db' is a valid and initialized database context
                // Check if 'db' is not null and BulkInsert method is accessible
                if (db != null)
                {
                    db.BulkInsert(list_obj);
                    return "True";
                }
                else
                {
                    return "False - Database context is not initialized.";
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                Console.WriteLine($"Error in Create method: {ex.Message}");
                return "False - An error occurred.";
            }
        }
        public string Create5(List<tbSavedFunctions> list_obj)
        {
            try
            {


                // Assuming 'db' is a valid and initialized database context
                // Check if 'db' is not null and BulkInsert method is accessible
                if (db != null)
                {
                    db.tbSavedFunctions.AddRange(list_obj);
                    return SaveChanges().ToString();
                }
                else
                {
                    return "False - Database context is not initialized.";
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                Console.WriteLine($"Error in Create method: {ex.Message}");
                return "False - An error occurred.";
            }
        }


        public string Create(tbSavedFunctions obj)
        {
            try
            {
                db.tbSavedFunctions.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception e)
            {

                return "false";
            }
        }

        public string Create5tial(List<tbmoalfefishexcel> list_obj)
        {
            try
            {


                // Assuming 'db' is a valid and initialized database context
                // Check if 'db' is not null and BulkInsert method is accessible
                if (db != null)
                {
                    db.tbmoalfefishexcel.AddRange(list_obj);
                    return SaveChanges().ToString();
                }
                else
                {
                    return "False - Database context is not initialized.";
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                Console.WriteLine($"Error in Create method: {ex.Message}");
                return "False - An error occurred.";
            }
        }


        public string Createtoal(tbmoalfefishexcel obj)
        {
            try
            {
                db.tbmoalfefishexcel.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception e)
            {

                return "false";
            }
        }

        public int Create2(tbSavedFunctions obj)
        {
            try
            {
                db.tbSavedFunctions.Add(obj);
              SaveChanges().ToString();
                return obj.svdfunc_ID;

            }
            catch (Exception e)
            {

                return 1;
            }
        }
        public int? CreateandReturnID(tbSavedFunctions obj)
        {
            //try
            //{
            var a = db.tbSavedFunctions.Add(obj);
            db.SaveChanges();
            return obj.svdfunc_ID;
            if (Convert.ToBoolean(db.tbSavedFunctions.Add(obj)))
            {
                return obj.svdfunc_ID;
            }
            else
            {
                return null;
            }


            //}
            //catch (Exception e)
            //{

            //    return null;
            //}
        }


        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
        public string Update2(tbSavedFunctions obj)
        {
            try
            {
                var old = Find(obj.svdfunc_ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    db.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges().ToString();
                }
                else
                    return "False";
            }
            catch(Exception e)
            {
                return "False";
            }
        }
        public tbSavedFunctions Find(int ID)
        {
            return db.tbSavedFunctions.Find(ID);
        }

        public List<tbSavedFunctions> Update()
        {
            return db.tbSavedFunctions.ToList();
        }

        public bool SaveChanges()
        {
            try
            {
                return Convert.ToBoolean(db.SaveChanges());

            }
            catch (DbUpdateException ex)
            {
                // Log the exception and its inner exception(s) for further analysis
                Exception innerException = ex;

                while (innerException != null)
                {
                    Console.WriteLine(innerException.Message);
                    innerException = innerException.InnerException;
                }

                // Log the exception using your logging mechanism (e.g., log4net, Serilog, etc.)
                // For example: logger.LogError(ex, "Error updating entries in the database.");

                return false;
            }
            catch (Exception e)
            {
                // Log the general exception
                Console.WriteLine(e.Message);

                // Log the exception using your logging mechanism
                // For example: logger.LogError(e, "An unexpected error occurred during SaveChanges.");

                return false;
            }
        }


        public string Update(tbSavedFunctions obj)
        {

            return "chert";

        }
        public int UpdateReturnID(tbSavedFunctions obj)
        {
            try
            {
                var result = db.tbSavedFunctions.SingleOrDefault(b => b.svdfunc_ID == obj.svdfunc_ID);
                if (result != null)
                {
                    result = obj;
                    try
                    {
                        db.SaveChanges();
                    }
                    catch (Exception e) { 
                    
                    }
                    return obj.svdfunc_ID;
                }
                return obj.svdfunc_ID;
            }
            catch (Exception e)
            {

                return 0;
            }


        }
    }
}