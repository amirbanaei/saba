using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;

namespace SaabWebProject.Models.Repositories.Salaries.Formula
{
    public class ZaribForFishRepository : IInterFace<tbPeymanZaribForFish>
    {
        SaabEntities db;


        public ZaribForFishRepository(SaabEntities Context)
        {
            db = Context;
        }
        public ZaribForFishRepository()
        {
            db = new SaabEntities();
        }
        public string Create(tbPeymanZaribForFish obj)
        {
            try
            {
                db.tbPeymanZaribForFish.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";

            }
        }
        public string Create(List<tbPeymanZaribForFish> list_obj)
        {
            try
            {
                db.tbPeymanZaribForFish.AddRange(list_obj);
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

        public tbPeymanZaribForFish Find(int ID)
        {
            return db.tbPeymanZaribForFish.Find(ID);
        }

        public bool SaveChanges()
        {
            return System.Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbPeymanZaribForFish obj)
        {
            throw new NotImplementedException();
        }

        public List<tbPeymanZaribForFish> Update()
        {
            throw new NotImplementedException();
        }
    }
}