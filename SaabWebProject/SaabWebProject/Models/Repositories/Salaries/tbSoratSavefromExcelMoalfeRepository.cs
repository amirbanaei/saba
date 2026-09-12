using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbSoratSavefromExcelMoalfeRepository : IInterFace<tbSoratSavefromExcelMoalfe>
    {
        #region متغیرها
        SaabEntities db;

        #endregion

        #region سازنده ها
        public tbSoratSavefromExcelMoalfeRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbSoratSavefromExcelMoalfeRepository()
        {
            db = new SaabEntities();
        }
        #endregion
        public string Create(tbSoratSavefromExcelMoalfe obj)
        {
            try
            {
                db.tbSoratSavefromExcelMoalfe.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }
        public string Create(List<tbSoratSavefromExcelMoalfe> list_obj)
        {
            try
            {
                db.tbSoratSavefromExcelMoalfe.AddRange(list_obj);
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

        public tbSoratSavefromExcelMoalfe Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return System.Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbSoratSavefromExcelMoalfe obj)
        {
            throw new NotImplementedException();
        }

        public List<tbSoratSavefromExcelMoalfe> Update()
        {
            throw new NotImplementedException();
        }
    }
}