using System.Collections.Generic;

namespace SaabWebProject.Models.Interfaces
{
    public interface IInterFace<T> where T : class
    {
        string Create(T obj);
        string Update(T obj);
        T Find(int ID);
        bool SaveChanges();
        List<T> Update();
        bool Disable(int ID);

    }
}
