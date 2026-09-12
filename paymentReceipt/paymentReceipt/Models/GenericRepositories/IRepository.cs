using System;
using System.Linq;
using System.Linq.Expressions;
using System.Collections.Generic;

namespace paymentReceipt.Models.GenericRepositories
{
    public interface IRepository<TEntity> where TEntity : class
    {
        void Insert(TEntity Entry);
        void Update(TEntity Entry);
        void Delete(TEntity Entry);
        void Update(int Id);
        void Delete(int Id);
        TEntity Find(int Id);
        TEntity Find(Guid Id);
        IEnumerable<TEntity> Get(Expression<Func<TEntity, bool>> Filter = null, Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> OrderBy = null, string Properties = "");
    }
}
