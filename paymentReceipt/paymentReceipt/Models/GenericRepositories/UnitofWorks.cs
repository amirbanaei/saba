using paymentReceipt.Models.DomainModels;
using System;
using System.Collections;



namespace paymentReceipt.Models.GenericRepositories
{
    public class UnitofWorks : IUnitofWorks, IDisposable
    {
        private bool disposedValue;
        private readonly paymentReceiptEntities db;
        private Hashtable _repositories;

        public UnitofWorks()
        {
            db = new paymentReceiptEntities();
        }

        public UnitofWorks(paymentReceiptEntities _context)
        {
            db = _context;
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
                    .MakeGenericType(typeof(TEntity)), db);

            _repositories.Add(type, repositoryInstance);

            return (IRepository<TEntity>)_repositories[type];
        }

        public virtual void Commit()
        {
            using (var Trans = db.Database.BeginTransaction())
            {
                try
                {
                    db.SaveChanges();
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
                db.SaveChanges();
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
                    db.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        ~UnitofWorks()
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
