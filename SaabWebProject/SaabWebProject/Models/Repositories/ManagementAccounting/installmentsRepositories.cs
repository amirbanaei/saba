using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.ManagementAccounting
{
    public class installmentsRepositories : IInterFace<tbinstallments>
    {
        #region متغیرها
        SaabEntities db;


        public installmentsRepositories(SaabEntities Context)
        {
            db = Context;
        }
        public installmentsRepositories()
        {
            db = new SaabEntities();
        }
        #endregion



        public string Create(tbinstallments obj)
        {

            try
            {
                db.tbinstallments.Add(obj);
                if (SaveChanges())
                {
                    return obj.installment_ID.ToString();
                }
                else
                {
                    return "0";
                }
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(tb_Subset_of_installments obj)
        {

            try
            {
                db.tb_Subset_of_installments.Add(obj);
                return SaveChanges().ToString();

            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbinstallments> list_obj)
        {
            try
            {
                db.tbinstallments.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public string Create(List<tb_Subset_of_installments> list_obj)
        {
            try
            {
                var ins_ID = list_obj.First().fk_installment_ID;
                var olds = db.tb_Subset_of_installments.Where(p => p.fk_installment_ID == ins_ID).ToList();
                if(olds != null)
                {
                    db.tb_Subset_of_installments.RemoveRange(olds);
                }
                db.tb_Subset_of_installments.AddRange(list_obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }



        public bool Delete(int ID)
        {
            try
            {
                db.tbinstallments.Remove(Find(ID));
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch (Exception)
            {

                return false;
            }

        }

        public tbinstallments Find(int ID)
        {
            return db.tbinstallments.Find(ID);
        }


        public List<tbinstallments> Update3()
        {
            return db.tbinstallments.ToList();
        }

        public bool SaveChanges()
        {
            try
            {
                db.SaveChanges();
                return true; // تغییر این خط
            }
            catch (Exception ex)
            {
                return false; // تغییر این خط
            }
        }


        public string Update(tbinstallments obj )
        {
            try
            {
                var old =Find(obj.installment_ID);
                if (old != null)
                {
                    if(obj.File_SystemName == null)
                    {
                        obj.File_SystemName = old.File_SystemName;
                        obj.File_RealName = old.File_RealName;
                    }
                    if(obj.Peyman_ID == null)
                    {
                        obj.Peyman_ID = old.Peyman_ID;
                    }
                    if(obj.User_ID == null)
                    {
                        obj.User_ID = old.User_ID;
                      }

                    db.Entry(old).CurrentValues.SetValues(obj);

                    if (SaveChanges())
                    {
                        return obj.installment_ID.ToString();
                    }
                    else
                    {
                        return "0";
                    }
                }
                else
                    return "False";
            }
                
            catch(Exception ex)
            {
                return "False";
            }
        }

        public bool Disable(int ID)
        {
            db.tbinstallments.Remove(db.tbinstallments.Find(ID));
            return SaveChanges();
        }
        public string Update2(tb_Subset_of_installments obj)
        {
            try
            {
                var old = Find(obj.Subsetins_ID);
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

        public List<tbinstallments> Update()
        {
            throw new NotImplementedException();
        }

        //public bool Disable(int ID)
        //{
        //    var old = Find(ID);
        //    if (old != null)
        //    {
        //        if ((bool)old.md_IsActive)
        //        {
        //            old.md_IsActive = false;
        //        }
        //        else
        //        {
        //            old.md_IsActive = true;
        //        }
        //        var result = SaveChanges();
        //        return result;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}
        //public string Update(List<tbinstallments> obj)
        //{
        //    try
        //    {
        //        var old = Find(obj.installment_ID);
        //        if (old != null)
        //        {
        //            //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
        //            db.Entry(old).CurrentValues.SetValues(obj);
        //            return SaveChanges().ToString();
        //        }
        //        else
        //            return "False";
        //    }
        //    catch
        //    {
        //        return "False";
        //    }
        //}
    }
}