using System.Collections.Generic;
using SaabWebProject.Models.DomainModels;

namespace SaabWebProject.Models.ViewModels.Users
{
    public class UserAndActionViewModel
    {
        public List<tbActions> Actions { get; set; }
        public List<tbUsers> Users { get; set; }
    }
}