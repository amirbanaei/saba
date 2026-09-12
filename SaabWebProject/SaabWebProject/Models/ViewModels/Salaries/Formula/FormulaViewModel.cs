using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public class FormulaViewModel
    {
        public string FormulaSchema { get; set; }
        public string FormulaView { get; set; }
        public List<string> VariablesId { get; set; }
        public string Title { get; set; }

        public int? MoalefeID { get; set; }
        public int? FormulaID { get; set; }
    }
}