using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.MoalefeKarkardi
{
    public class UserAndMoalefeVM
    {
        public List<tbUsers> Userlst { get; set; }
        public List<tbContractMoalefeDastmozdi> MaolefeKarkardilst { get; set; }
    }
}