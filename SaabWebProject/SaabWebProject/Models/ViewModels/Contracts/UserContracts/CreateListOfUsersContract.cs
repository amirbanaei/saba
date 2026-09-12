using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.ViewModels.Contracts.UserContracts
{
    public class CreateListOfUsersContract
    {
        public tbUserContracts obj_tbuserContracts {  get; set;}
        public List<int> list_usersIDs { get; set; }
    }
}