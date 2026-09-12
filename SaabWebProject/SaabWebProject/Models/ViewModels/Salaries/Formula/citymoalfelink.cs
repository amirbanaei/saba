using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.ViewModels.Contracts.Function;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public class citymoalfelink
    {
        public Nullable<int> Fk_pymn { get; set; }
        public Nullable<int> Fk_City { get; set; }
        public Nullable<int> FK_moalfe { get; set; }
        public List<tblink_moalfe_city_valu> innerObjects { get; set; }
    }
}