using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Areas.Contracts.Models.Classes
{
    public class EquipmentVM
    {
        public List<tbEquipmentSpecifications> listEquip { get; set; }
        public int Counter { get; set; }
    }
}