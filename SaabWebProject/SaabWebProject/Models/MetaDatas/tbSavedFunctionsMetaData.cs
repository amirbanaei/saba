using SaabWebProject.Models.Classes;
using SaabWebProject.Models.Utilitis;
using SaabWebProject.Models.ViewModels;
using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace SaabWebProject.Models.MetaDatas
{
    public class tbSavedFunctionsMetaData
    {
    }

}

namespace SaabWebProject.Models.DomainModels
{
    [MetadataType(typeof(MetaDatas.tbSavedFunctionsMetaData))]
    public partial class tbSavedFunctions : ViewModelBase, IDisposable
    {
      

        public string ShmasiENDTime
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(svdfunc_ToDate);
            }
        }
        public string ShmasiStartTime
        {
            get
            {
                return ConvertDateTimeToShamsi.ConvertDateTimeToShamsi1(svdfunc_FromDate);
            }
        }
        public bool OnlineUserUploadedFile
        {
            get;set;
        }
        public void Dispose()
        {
        }
    }
}