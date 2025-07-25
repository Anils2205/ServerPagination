using ServerPagination.DataAccess.StoredProcedureDbAccess.Abstraction;
using ServerPagination.Models;
using ServerPagination.Services.Abstraction;
namespace ServerPagination.Services.Repository
{
    public class HomeHelper : IHomeHelper
    {
        private readonly IHomeDbRepository _homeDbRepository;
        public HomeHelper(IHomeDbRepository homeDbRepository)
        {
            _homeDbRepository = homeDbRepository;
        }
        public void ActiveManage(int userId)
        {
            _homeDbRepository.ActiveManage(userId);
        }
        public void AddUser(UserModel user)
        {
            _homeDbRepository.AddUser(user);
        }
        public void DeleteUser(int userId)
        {
            _homeDbRepository.DeleteUser(userId);
        }
        public void EditUser(EditUserModel user)
        {
            _homeDbRepository.EditUser(user);
        }
        public EditUserModel GetUser(int userId)
        {
            return _homeDbRepository.GetUser(userId);
        }
        public PaginationModel UserList(SetPagination  setPagination)
        {
            var userData = _homeDbRepository.UserList(setPagination.PageNumber, setPagination.SearchQuery);
            return userData;
        }
    }
}