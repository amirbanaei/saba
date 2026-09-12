using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Classes.FinancialDocuments
{
    public class TarazViewModel
    {
        public bool IsPeyman { get; set; }
        public string FullName { get; set; }
        public string Persian_FromDate { get; set; }
        public string Persian_ToDate { get; set; }
        public string PersonelCode { get; set; }
        public string PeymanCode { get; set; }
        List<TarazItemsViewModel> Items { get; set; }

    }
}