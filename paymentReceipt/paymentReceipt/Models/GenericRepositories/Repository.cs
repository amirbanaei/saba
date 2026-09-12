using System;
using System.Linq;
using System.Data.Entity;
using System.Linq.Expressions;
using System.Collections.Generic;

using paymentReceipt.Models.DomainModels;

namespace paymentReceipt.Models.GenericRepositories
{
    public class Repository<TEntity> : IRepository<TEntity>, IDisposable  where TEntity : class
    {
        internal paymentReceiptEntities db;
        internal DbSet<TEntity> dbSet;
        private bool disposedValue;

        public Repository()
        {
            db = new paymentReceiptEntities();
            dbSet = db.Set<TEntity>();
        }

        public Repository(paymentReceiptEntities _context)
        {
            db = _context;
            dbSet = db.Set<TEntity>();
        }
       
        public virtual void Delete(TEntity Entry)
        {
            if(db.Entry(Entry).State == EntityState.Detached)
            {
                dbSet.Attach(Entry);
            }
            dbSet.Remove(Entry);
        }

        public virtual void Delete(int Id)
        {
            TEntity Entry = dbSet.Find(Id);
            Delete(Entry);
        }

        public virtual TEntity Find(int Id)
        {
            return dbSet.Find(Id);
        }

        public virtual TEntity Find(Guid Id)
        {
            return dbSet.Find(Id);
        }

        public virtual  IEnumerable<TEntity> Get(Expression<Func<TEntity, bool>> Filter = null, Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> OrderBy = null, string Properties = "")
        {
            IQueryable<TEntity> Query = dbSet;

            if(Filter != null)
            {
                Query = Query.Where(Filter);
            }

            if(Properties != null)
            {
                foreach (var item in Properties.Split(new char []{ ',' },StringSplitOptions.RemoveEmptyEntries))
                {
                    Query = Query.Include(item);
                }
            }

            if (OrderBy != null)
            {
                return OrderBy(Query).ToList();
            }
            else
            {
                return Query.ToList();
            }
        }

        public virtual void Insert(TEntity Entry)
        {
            dbSet.Add(Entry);
        }

        public virtual void Update(TEntity Entry)
        {
            dbSet.Attach(Entry);
            db.Entry(Entry).State = EntityState.Modified;
        }

        public virtual void Update(int Id)
        {
            TEntity Entry = dbSet.Find(Id);
            Update(Entry);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    db.Dispose();
                    dbSet = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        ~Repository()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
