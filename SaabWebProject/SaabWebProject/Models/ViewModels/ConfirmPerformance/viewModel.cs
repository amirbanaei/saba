using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.ConfirmPerformance
{
    public class viewModel
    {
        public List<string> name { get; set; }
        public List<string> Family { get; set; }
        public List<double> Value { get; set; }

        public viewModel()
        {
            name = new List<string>();
            Family = new List<string>();
            Value = new List<double>();
        }

        public void AddData(string name, string family, double value)
        {
            this.name.Add(name);
            this.Family.Add(family);
            this.Value.Add(value);
        }
    }

}