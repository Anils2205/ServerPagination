using ServerPagination.Models;
namespace ServerPagination.Services.Abstraction
{
    public interface IHomeHelper
    {
        void ActiveManage(int userId);
        void AddUser(UserModel user);
        void DeleteUser(int userId);
        void EditUser(EditUserModel user);
        EditUserModel GetUser(int userId);
        PaginationModel UserList(SetPagination setPagination);
    }
}