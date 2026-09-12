using SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Contracts.RigesterDocument
{
    public class Rigseterview
    {
        public string Namemoalfe { get; set; }
        public List<information> informationtool { get; set; }


    }


    public class information
    {
        public string Name { get; set; }
        public int personalID { get; set; }
        public float fish { get; set; }
        public float excel { get; set; }

    }
}