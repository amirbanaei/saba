using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public class SoratHeader
    {
        public string PeymanName { get; set; }

        public string kargozariName { get; set; }

        public string startdate { get; set; }
        public string Endtdate { get; set; }

    
        

        public string cityname { get; set; }
        public int kargozaricode { get; set; }
        public int year { get; set; }
        
        public int number_sorat { get; set; }
    }
} 