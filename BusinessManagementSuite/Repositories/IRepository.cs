using System.Collections.Generic;

namespace BusinessManagementSuite.Repositories;

public interface IRepository<T>
{
    IEnumerable<T> GetAll();
    T? GetById(long id);
    void Insert(T entity);
    void Update(T entity);
    void Delete(long id);
    int Count();
}
