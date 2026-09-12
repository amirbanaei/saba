using System;
using System.Collections.Generic;
using System.Linq;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;

namespace SaabWebProject.Models.Repositories
{
    public class tbContractJobGroupRepository : IInterFace<tbContractJobGroup>
    {
        #region متغیر ها
        SaabEntities Context;
        #endregion

        #region سازنده ها
        public tbContractJobGroupRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbContractJobGroupRepository()
        {
            Context = new SaabEntities();
        }

        #endregion

        public string Create(tbContractJobGroup obj)
        {
            try
            {
                Context.tbContractJobGroup.Add(obj);
                return SaveChanges().ToString();
            }
            catch(Exception e)
            {
                return "False";
            }
        }
        /*------------------------------------------------[1402/08/11]-|KH|-*/
        public string CreateValue(tbContractJobGroupValues obj)
        {
            try
            {
                Context.tbContractJobGroupValues.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception e)
            {
                return "False";
            }
        }
        /*-------------------------------------------------------------|KH|-*/

        public string Update(tbContractJobGroup obj)
        {
            try
            {
                var old = FindYear(obj.jg_Year??0);
                if(old != null)
                {
                    old.isActive = false;
                    return Create(obj);
                    //old.jg_KharoBar = obj.jg_KharoBar;
                    //old.jg_HagheOlad = obj.jg_HagheOlad;
                    //old.jg_HagheMaskan = obj.jg_HagheMaskan;
                    //old.jg_CreateTime = DateTime.Now;
                    //foreach(var item in old.tbContractJobGroupValues.ToList())
                    //{
                    //    if(item.SanavatYear == 0) // گروه شغلی است
                    //    {
                    //        item.ValueMozdGroup = obj.tbContractJobGroupValues.Where(p => p.SanavatYear == item.SanavatYear
                    //        && p.GroupNumber == item.GroupNumber).FirstOrDefault().ValueMozdGroup;
                    //    }
                    //    else
                    //    {
                    //        item.ValueSanavat = obj.tbContractJobGroupValues.Where(p => p.SanavatYear == item.SanavatYear
                    //        && p.GroupNumber == item.GroupNumber).FirstOrDefault().ValueSanavat;
                    //    }
                    //}
                    //return SaveChanges().ToString();
                }
                else
                {
                    return "False";
                }
            }
            catch
            {
                return "False";
            }
        }
        public tbContractJobGroup Find(int ID)
        {
            try
            {
                return Context.tbContractJobGroup.Find(ID);
            }
            catch
            {
                return null;
            }
        }
        public bool SaveChanges()
        {
            return Convert.ToBoolean(Context.SaveChanges());
        }
        public List<tbContractJobGroup> Update()
        {
            return Context.tbContractJobGroup.ToList();
        }
        public bool Disable(int ID)
        {
            //var old = Find(ID);
            //if (old != null)
            //{
            //    old.md_IsActive = !old.md_IsActive;
            //    return SaveChanges();
            //}
            //else
            //{
            //}
            return true;
        }

        public tbContractJobGroup FindYear(int year)
        {
            return Context.tbContractJobGroup.Where(p => p.jg_Year == year && p.isActive != false).FirstOrDefault();
        }

        public tbContractJobGroup GetBaseInformation(int year , int groupNumber , int tedadSalSanavart)
        {
            var result = FindYear(year);
            tbContractJobGroup obj = new tbContractJobGroup();

            if (result != null)
            {
                obj = result;
                obj.tbContractJobGroupValues = result.tbContractJobGroupValues.Where(p => p.GroupNumber == groupNumber && (p.SanavatYear == 0 || p.SanavatYear == tedadSalSanavart)).ToList();
                if (obj.tbContractJobGroupValues == null)
                {
                    obj.tbContractJobGroupValues = new List<tbContractJobGroupValues>();
                }
                return obj;
            }
            obj.tbContractJobGroupValues = new List<tbContractJobGroupValues>();
            return obj;
        }


    }
}