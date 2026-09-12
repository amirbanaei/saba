using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbPeymanElhaghieRepository : IInterFace<tbPeymanElhaghie>
    {
        SaabEntities Context;
        public tbPeymanElhaghieRepository()
        {
            Context = new SaabEntities();
        }
        public tbPeymanElhaghieRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public string Create(tbPeymanElhaghie obj)
        {
            try
            {
                Context.tbPeymanContracts.Where(p => p.pec_ID == obj.FK_PeymanID).FirstOrDefault().pec_Price = obj.pec_Price; ;
                Context.tbPeymanElhaghie.Add(obj);
                return SaveChanges().ToString();
            }
            catch(Exception e)
            {
                return "False";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbPeymanElhaghie Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbPeymanElhaghie> Update()
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(Context.SaveChanges());
        }

        public string Update(tbPeymanElhaghie obj)
        {
            throw new NotImplementedException();
        }
    }
}