using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbFormulaRepositories : IInterFace<tbFormula>
    {
        #region متغیرها
        SaabEntities db;

        #endregion

        #region سازنده ها
        public tbFormulaRepositories(SaabEntities Context)
        {
            db = Context;
        }
        public tbFormulaRepositories()
        {
            db = new SaabEntities();
        }
        #endregion

        public string Create(tbFormula obj)
        {
            try
            {
                db.tbFormula.Add(obj);
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

        public tbFormula Find(int ID)
        {
            return db.tbFormula.Find(ID);
        }
        public tbFormula FindByFKMoalefeDastMozdi(int ID)
        {
            return db.tbFormula.Where(p => p.FKMoalefeDastMozdi == ID).FirstOrDefault();
        }

        public List<tbFormula> Update()
        {
            var model = db.tbFormula.Where(p=> p.Frml_IsActive == true).ToList();

            return model;
        }

        public bool SaveChanges()
        {
            
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbFormula obj)
        {
            try
            {
                var formul = db.tbFormula.Where(p=> p.FKMoalefeDastMozdi == obj.FKMoalefeDastMozdi && p.Frml_IsActive==true).FirstOrDefault();
                if (formul != null)
                {
                    formul.Frml_Formula = obj.Frml_Formula;
                    formul.Frml_Title = obj.Frml_Title;
                    formul.Frml_Variablesid = obj.Frml_Variablesid;
                    formul.Frml_Schema = obj.Frml_Schema;

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

        public void Remove(tbFormula tbFormula)
        {
            db.tbFormula.Remove(tbFormula);
        }
    }
}