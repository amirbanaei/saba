using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbDocumentsMetaData
    {
    }

}


namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbDocumentsMetaData))]
    public partial class tbDocuments: ViewModelBase
    {
         public string peyman_Title
        {
            get
            {
                return tbPeymanContracts.pec_Title;
            }
            set
            {

            }
        }
        public int? type_control
        {
            get
            {

               using(var db1 = new SaabEntities())
                {
                    return db1.tbDocumentsDefinitionOfDocumentTitles.Find(FK_DDOfTitle).TypeOfControl;
                }
            }
            set
            {

            }
        }
        public string doc_Title
        {
            get
            {

                using (var db1 = new SaabEntities())
                {
                    return db1.tbDocumentsDefinitionOfDocumentTitles.Find(FK_DDOfTitle).Title;
                }
            }
            set
            {

            }
        }
        public bool? doc_ableToSaveSummery
        {
            get
            {

                using (var db1 = new SaabEntities())
                {
                    return db1.tbDocumentsDefinitionOfDocumentTitles.Find(FK_DDOfTitle).AbleToSetSummery;
                }
            }
            set
            {

            }
        }

    }
}