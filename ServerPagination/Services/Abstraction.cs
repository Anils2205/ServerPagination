using System.Collections.Generic;
using System.Threading.Tasks;
namespace ServerPagination.Services
{
    public interface IManageService<TRequestEntity, TResponseEntity>
    {
        Task<TResponseEntity> GetAsync(string aRequestUri, int id);
        Task<List<TResponseEntity>> GetAllByIdAsync(string aRequestUri, string id);
        Task<List<TResponseEntity>> GetAllAsync(string aRequestUri);
        Task<TResponseEntity> PostAsync(string aRequestUri, TRequestEntity aObj);
        Task<TResponseEntity> PutAsync(string aRequestUri, TRequestEntity aObj);
        Task<TResponseEntity> DeleteAsync(string aRequestUri, int id);
    }
}