using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.ViewModels
{
    public class ContractsListViewModel
    {
        public List<tbUserContracts> list_usercontracts { get; set; }  
        public List<tbUserRentContracts> list_rentContracts { get; set; }
    }
}