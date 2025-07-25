using System.Collections.Generic;
namespace ServerPagination.Models
{
    public class PaginationModel
    {
        public Pagination Pagination { get; set; }
        public List<UserModel> UserList { get; set; }
    }
}