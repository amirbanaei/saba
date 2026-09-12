using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using static Stimulsoft.Report.StiRecentConnections;

namespace SaabWebProject.Models.Repositories
{
    public class tbUserContractsRepository
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbUserContractsRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbUserContractsRepository()
        {
            Context = new SaabEntities();
        }
        #endregion
        public bool Create(tbUserContracts obj)
        {
            try
            {
                Context.tbUserContracts.Add(obj);
                return SaveChanges();
            }
            catch(Exception ee)
            {
                return false;
            }
        }
        //public int sabt()
        //{
        //    var obj1 = new tbUserContractsAndMoalefeGhararDadi();
        //    obj1.FKMoalefeGhararDadi = 1800;
        //    obj1.Value =45 ;
        //    obj1.FKContractID = 440;
        //    Context.tbUserContractsAndMoalefeGhararDadi.Add(obj1);
        //    Context.SaveChanges();
        //    return 1;

        //}
        public bool Create(List<tbUserContracts> list_obj1)
        {
            try
            {

                foreach (var item2 in list_obj1)
                {

                    if (item2.tbUserContractsAndMoalefeGhararDadi.Count() != 0)
                    {
                        List<tbUserContractsAndMoalefeGhararDadi> list_obj = new List<tbUserContractsAndMoalefeGhararDadi>();

                        foreach (var item in item2.tbUserContractsAndMoalefeGhararDadi)
                        {
                            var obj1 = new tbUserContractsAndMoalefeGhararDadi();
                            obj1.FKMoalefeGhararDadi = item.FKMoalefeGhararDadi;
                            obj1.Value = item.Value;
                            //obj1.FKContractID=item2.usc_ID;
                            //Context.tbUserContractsAndMoalefeGhararDadi.Add(obj1);
                            //Context.SaveChanges();
                            list_obj.Add(obj1);
                        }
                        item2.tbUserContractsAndMoalefeGhararDadi = list_obj;
                    }
                }
                //Context.tbUserContractsAndMoalefeGhararDadi.AddRange(list_obj);
                Context.tbUserContracts.AddRange(list_obj1);

                return SaveChanges();

            }
            catch (Exception e)
            {
                return false;
            }
        }
        public bool Update(tbUserContracts obj)
        {
            try
            {
                var old = Find(obj.usc_ID);
                if (old != null)
                {
                    //Context.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                    Context.Entry(old).CurrentValues.SetValues(obj);
                    return SaveChanges();
                }
                else
                    return true;
            }
            catch
            {
                return false;
            }
        }
        public tbUserContracts Find(int ID)
        {
            try
            {
                return Context.tbUserContracts.Find(ID);
            }
            catch
            {
                return null;
            }
        }
        public bool SaveChanges()
        {
            using (var transaction = Context.Database.BeginTransaction())
            {
                try
                {
                    Context.SaveChanges();
                    transaction.Commit();
                    return true;
                }
                catch
                {
                    transaction.Rollback();
                    return false;
                }
            }
            //return Convert.ToBoolean(Context.SaveChanges());
        }
        public List<tbUserContracts> Listt()
        {
            return Context.tbUserContracts.ToList();
        }
        public string Disable(int ID)
        {
            try
            {
                var find = Context.tbUserContracts.Find(ID);
                var findd = Context.tbUserContractsAndMoalefeGhararDadi.Where(p=>p.FKContractID == ID).ToList();
                Context.tbUserContractsAndMoalefeGhararDadi.RemoveRange(findd);
                Context.SaveChanges();
                Context.tbUserContracts.Remove(find);
                return SaveChanges().ToString();
            }
            catch (Exception)
            {

                return "False";
            }
            //var old = Find(ID);
            //if (old != null)
            //{
            //    if (old.usc_Accepted != true)
            //    {
            //        Context.Entry(old).State = EntityState.Deleted;
            //        return SaveChanges().ToString();
            //    }
            //    else
            //    {
            //        return "false_Accepted";
            //    }
            //}
            //else
            //{
            //    return "true_notFound";
            //}
            return null;
        }
        public bool AcceptUserContract(int userContractID)
        {
            //var old = Find(userContractID);
            //try
            //{
            //    old.usc_Accepted = true;
            //    return SaveChanges();
            //}
            //catch (Exception)
            //{

            //    return false;
            //}
            return false ;
        }
        public List<tbUserContracts> Listt(int? userID)
        {
            var a = Context.tbUserContracts.Where(p => p.FK_UserID == userID).OrderByDescending(p=>p.usc_StartTime).ToList();
            return a;
        }

    }
}