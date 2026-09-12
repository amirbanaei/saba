using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using SaabWebProject.Models.UnitOfWork.Repositories;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Access.Models.UnitofWorks.Repository 
{
    public class UnitofWork : IUnitofWork, IDisposable
   
    {
        private bool disposedValue;
        private readonly SaabEntities Context;
        private Hashtable _repositories;
        public UnitofWork()
        {
            Context = new SaabEntities();
        }

        public UnitofWork(SaabEntities _context)
        {
            Context = _context;
        }

        public virtual IRepository<TEntity> Repository<TEntity>() where TEntity : class
        {
            if (_repositories == null)
                _repositories = new Hashtable();

            var type = typeof(TEntity).Name;

            if (_repositories.ContainsKey(type)) return (IRepository<TEntity>)_repositories[type];

            var repositoryType = typeof(Repository<>);

            var repositoryInstance =
                Activator.CreateInstance(repositoryType
                    .MakeGenericType(typeof(TEntity)), Context);

            _repositories.Add(type, repositoryInstance);

            return (IRepository<TEntity>)_repositories[type];
        }

        public virtual void Commit()
        {
            using (var Trans = Context.Database.BeginTransaction())
            {
                try
                {
                    Context.SaveChanges();
                    Trans.Commit();
                }
                catch (Exception e)
                {
                    Trans.Rollback();
                }
            }
        }

        public virtual void SaveChange()
        {
            try
            {
                Context.SaveChanges();
            }
            catch (Exception e)
            {

            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    Context.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        ~UnitofWork()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
    
