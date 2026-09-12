namespace paymentReceipt.Models.GenericRepositories
{
    public interface IUnitofWorks
    {
        void Commit();
        void SaveChange();
        IRepository<TEntity> Repository<TEntity>() where TEntity : class;
    }
}
