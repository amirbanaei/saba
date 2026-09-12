using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.UnitOfWork.Interfaces;

namespace SaabWebProject.Models.UnitOfWork.Repositories
{
    public class Repository<TEntity> : IRepository<TEntity>, IDisposable where TEntity : class
    {
        private SaabEntities context;
        internal DbSet<TEntity> dbSet;
        private bool disposedValue;

        internal SaabEntities Context { get => context; set => context = value; }

        public Repository()
        {
            Context = new SaabEntities();
            dbSet = Context.Set<TEntity>();
           
        }

        public Repository(SaabEntities _context)
        {
            Context = _context;
            dbSet = Context.Set<TEntity>();
        }
        public virtual void Delete(TEntity Entry)
        {
            if (Context.Entry(Entry).State == EntityState.Detached)
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

        
        public TEntity Find(int Id)
        {
            return dbSet.Find(Id);
        }

        public TEntity Find(Guid Id)
        {
            return dbSet.Find(Id);
        }

        public IEnumerable<TEntity> Get(Expression<Func<TEntity, bool>> Filter = null, Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> OrderBy = null, string Properties = "")
        {
            IQueryable<TEntity> Query = dbSet;

            if (Filter != null)
            {
                Query = Query.Where(Filter);
            }

            if (Properties != null)
            {
                foreach (var item in Properties.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
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
    
        public void Insert(TEntity Entry)
        {
            dbSet.Add(Entry);
        }

        public void Update(TEntity Entry)
        {
            dbSet.Attach(Entry);
            Context.Entry(Entry).State = EntityState.Modified;
        }

        public void Update(int Id)
        
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
                    Context.Dispose();
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

        public void AddRange(IEnumerable<TEntity> lstEntry)
        {
            dbSet.AddRange(lstEntry);
        }
    }
}

    
