using ExcelLibrary.BinaryFileFormat;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries.Formula
{
    public class tb_Step_taxRepositories : IInterFace<tb_Step_tax>
    {

        SaabEntities db;


        public tb_Step_taxRepositories(SaabEntities Context)
        {
            db = Context;
        }
        public tb_Step_taxRepositories()
        {
            db = new SaabEntities();
        }
        public string Create(tb_Step_tax obj)
        {
            try
            {
                db.tb_Step_tax.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public string Create(List<tb_Step_tax> list_obj)
        {
            try
            {
                db.tb_Step_tax.AddRange(list_obj);
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

        public tb_Step_tax Find(int ID)
        {
            return db.tb_Step_tax.Find(ID);
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tb_Step_tax obj)
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

        public string Update(List<tb_Step_tax> objList)
        {
            try
            {
                foreach (var obj in objList)
                {
                    var old = Find(obj.ID);
                    if (old != null)
                    {
                        db.Entry(old).CurrentValues.SetValues(obj);
                    }
                    else
                    {
                        return "False"; // Assuming all records should exist for update
                    }
                }

                return SaveChanges().ToString();
        
            }
            catch (Exception ex)
            {
                // Log the exception or handle it appropriately
                return "Error: " + ex.Message;
            }
        }

        public List<tb_Step_tax> List()
        {
            return db.tb_Step_tax.ToList();
        }

        public List<tb_Step_tax> Update()
        {
            throw new NotImplementedException();
        }
        public void Remove(tb_Step_tax tbFormula)
        {
            db.tb_Step_tax.Remove(tbFormula);
        }
    }
}