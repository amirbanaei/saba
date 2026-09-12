using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Contracts
{
    public class tbLinkuserAndPeymanRepository : IInterFace<Link_User_And_Peyman>
    {
        SaabEntities db;
        public tbLinkuserAndPeymanRepository(SaabEntities Context)
        {
            db = Context;
        }
        public tbLinkuserAndPeymanRepository()
        {
            db= new SaabEntities();
        }
        public string Create(Link_User_And_Peyman obj)
        {
            throw new NotImplementedException();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public Link_User_And_Peyman Find(int ID)
        {
            throw new NotImplementedException();
        }

        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(Link_User_And_Peyman obj)
        {
            throw new NotImplementedException();
        }

        public List<Link_User_And_Peyman> Update()
        {
            throw new NotImplementedException();
        }
        public List<tbUsers> GetUsers(int PeymanID)
        {
            return db.Link_User_And_Peyman.Where(p => p.FK_Peyman_ID == PeymanID).Select(p => p.tbUsers).ToList();
        }
    }
}