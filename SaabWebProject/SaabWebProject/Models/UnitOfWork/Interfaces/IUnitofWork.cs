using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaabWebProject.Models.UnitOfWork.Interfaces
{
    interface IUnitofWork
    {
    
            void Commit();
            void SaveChange();
            IRepository<TEntity> Repository<TEntity>() where TEntity : class;
        
    }
}
