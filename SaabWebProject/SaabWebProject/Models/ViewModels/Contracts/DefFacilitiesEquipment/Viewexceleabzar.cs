using SaabWebProject.Models.ViewModels.Contracts.Function;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Contracts.DefFacilitiesEquipment
{
    public class Viewexceleabzar
    {
        public int FK_BASTE { get; set; }

        public Header SabtHeader { get; set; }
        public List<information> informationtool { get; set; }
    }
    public class Virwforexcel
    {
        public int FK_BASTE { get; set; }
        public Header SabtHeader { get; set; }
        public List<informationTools> informationToolsss { get; set; }
    }
    public class Header
    {
        public int MoalfeVal_Month { get; set; }
        public int MoalfeVal_Year { get; set; }
        public string PeymanName { get; set; }
        public string Type { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
    }
    public class information
    {
        public int ID { get; set; }
        public string name_tool { get; set; }
        //public string Other_tool_information { get; set;}
        public int number_recorded { get; set; }
        public int Approved_number { get; set; }
        public int days_number { get; set; }
        public int Condition { get; set; }

        public string Description { get; set; }
    }
    public class informationTools
    {
        public int ID { get; set; }
        public string name_tool { get; set; }
        public string Other_tool_information { get; set; }
        //public int number_recorded { get; set; }
        //public int Approved_number { get; set; }
        public int days_number { get; set; }
        public int Condition { get; set; }

        public string Description { get; set; }
    }

}