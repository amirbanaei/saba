using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Equipment
{
    public class FinantialEquipmentVM
    {
        public string UsersName { get; set; }
        public long TotalPrice { get; set; }
        public long TazminPrice  { get; set; }
        public string TazminType { get; set; }


        public long EachValue { get; set; }
        public int UnitCount { get; set; }

        public string Unit { get; set; }

        public DateTime? ExpireDate { get; set; }

        public List<EquipmentInfo> Tajhizat { get; set; }
    }
}