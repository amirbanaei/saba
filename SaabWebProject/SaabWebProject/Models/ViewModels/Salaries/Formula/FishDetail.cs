using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public class FishDetail
    {
        public tbUsers Users { get; set; }
        public tbUserContracts UserContracts { get; set; }
        
        public List<MoalefeDetail> Moalefe { get; set; }

    }
}