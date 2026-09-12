using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Salaries.Formula
{
    public  class ManualFishDetail
    {
        public List<FishValue> FishValues { get; set; }
        public List<FishValue> FishValuesKosoorat { get; set; }
        public FishHeader FishHeader{ get; set; }
        public List<FishValue> FishValueQaradad { get; set; }
        public List<FishValue> FishValuesEzafat { get; set; }

        public List<FishValue> FishValuesAqsat { get; set; }
        public int month { get; set; }

        public double JamNakhalesHoqoqVaMazaya { get;set; }
        public double kasrikarkard { get;set; }
        public double jamkolekosoorat { get; set; }
        public double KhalesQabelDaryaft { get;set; }
        public double ZakhireKarMazad { get;set; }

        public double ZakhireKarMazadQabl { get;set; }

        public double SayerEzafat { get;set; }
        public double sayerkosoorat { get;set; }

        public double BimetakmiliForFishTest { get; set; }

    }
}