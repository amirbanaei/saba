using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Areas.Contracts.Models.Classes
{
    public class ListOfMoalefeha
    {
        public string variablePersianName { get; set; }
        public long ID { get; set; }

        public int? Type { get; set; }
        public int? noghrogi { get; set; }

        public string GharardadColoumnName { get; set; }
        public bool ExistsFormula { get; set; }

        public bool visable { get; set; }
        
    }
}