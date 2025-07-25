using System.Collections.Generic;
using System.Data;
namespace ServerPagination.StoredProcedureDbAccess
{
    public interface IGenericRepository<out TEntity>
    {
        IDbConnection GetOpenConnection();
        TEntity GetSingle(int aSingleId);
        IEnumerable<TEntity> GetAll();
    }
}