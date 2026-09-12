using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Users
{
    public class CityViewModel
    {
        public int cityId { get; set; }
        public string provinceName { get; set; }    
        public string cityName { get; set; }
        public int provinceId { get; set; }
    }
}