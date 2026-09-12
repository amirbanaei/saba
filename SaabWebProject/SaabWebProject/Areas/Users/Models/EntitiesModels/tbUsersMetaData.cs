//using SaabWebProject.Areas.Users.Models.EntitiesModels;
//using System;
//using System.Collections.Generic;
//using System.ComponentModel.DataAnnotations;
//using System.Linq;
//using System.Web;

//namespace SaabWebProject.Areas.Users.Models.EntitiesModels
//{
//    internal class tbUsersMetaData
//    {

//    }
//}

//namespace SaabWebProject.Models.DomainModels
//{
//    [MetadataType(typeof(tbUsersMetaData))]
//    public partial class tbUsers : ViewModels.ViewModelBase, IDisposable
//    {
//        public string FullName
//        {
//            get => usr_Name + " " + usr_Family ;
//        }
//        public bool usr_RememberMe { get; set; }
//        #region IDisposable Support
//        private bool disposedValue = false; // To detect redundant calls

//        protected virtual void Dispose(bool disposing)
//        {
//            if (!disposedValue)
//            {
//                if (disposing)
//                {
//                }

//                disposedValue = true;
//            }
//        }

//        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
//        ~tbUsers()
//        {
//            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
//            Dispose(false);
//        }

//        void IDisposable.Dispose()
//        {
//            Dispose(true);
//        }
//        #endregion
//    }
//}
