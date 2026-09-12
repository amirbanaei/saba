using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Users
{
    public class ActionViewModel
    {
        public List<string> ActiveActions { get; set; } = new List<string>();
        public List<tbActions> AllActions { get; set; } = new List<tbActions>();
    }
}