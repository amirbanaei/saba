using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public class UserSalaryDetailVM
    {
        public bool HaveProblem { get; set; }
        public Dictionary<int,string> Detail { get; set; }
    }
}