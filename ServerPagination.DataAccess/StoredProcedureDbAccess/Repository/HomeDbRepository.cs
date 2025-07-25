using Dapper;
using Microsoft.EntityFrameworkCore;
using ServerPagination.DataAccess.StoredProcedureDbAccess.Abstraction;
using ServerPagination.Models;
using ServerPagination.WebedureDbAccess;
using System;
using System.Data;
using System.Linq;
namespace EmployeeAPIDemo.Web.DataAccess.StoredProcedureDbAccess.Repository
{
    public class HomeDbRepository : SqlDbRepository<UserModel>, IHomeDbRepository
    {
        public HomeDbRepository(string connectionstring) : base(connectionstring) { }
        public void ActiveManage(int userId)
        {
            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@UserId", userId);
            var CheckStatus = vconn.QueryFirstOrDefault<bool>("sp_proc_CheckStatus", vParams, commandType: CommandType.StoredProcedure);
            if (CheckStatus)
            {
                vconn.Execute("sp_proc_UserUnactive", vParams, commandType: CommandType.StoredProcedure);
            }
            else
            {
                vconn.Execute("sp_proc_UserActive", vParams, commandType: CommandType.StoredProcedure);
            }
        }
        public void AddUser(UserModel user)
        {
            user.Role = "User";
            if (user.Gender == "Male")
            {
                user.ProfileImage = "male.png";
            }
            else
            {
                user.ProfileImage = "female.png";
            }
            user.IsActive = true;
            user.IsDeleted = false;
            user.CreatedDate = DateTime.Now;
            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@Role", user.Role);
            vParams.Add("@ProfileImage", user.ProfileImage);
            vParams.Add("@FirstName", user.FirstName);
            vParams.Add("@LastName", user.LastName);
            vParams.Add("@Gender", user.Gender);
            vParams.Add("@Email", user.EmailAddress);
            vParams.Add("@MobileNo", user.MobileNo);
            vParams.Add("@Password", user.Password);
            vParams.Add("@IsActive", user.IsActive);
            vParams.Add("@IsDeleted", user.IsDeleted);
            vParams.Add("@CreatedDate", user.CreatedDate);
            vconn.Execute("sp_proc_AddUser", vParams, commandType: CommandType.StoredProcedure);
        }
        public void DeleteUser(int userId)
        {
            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@UserId", userId);
            vconn.Execute("sp_proc_DeleteUser", vParams, commandType: CommandType.StoredProcedure);
        }
        public void EditUser(EditUserModel user)
        {
            user.UpdatedDate = DateTime.Now;
            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@UserId", user.UserId);
            vParams.Add("@FirstName", user.FirstName);
            vParams.Add("@LastName", user.LastName);
            vParams.Add("@Gender", user.Gender);
            vParams.Add("@Email", user.EmailAddress);
            vParams.Add("@MobileNo", user.MobileNo);
            vParams.Add("@UpdatedDate", user.UpdatedDate);
            vconn.Execute("sp_proc_EditUser", vParams, commandType: CommandType.StoredProcedure);
        }
        public EditUserModel GetUser(int userId)
        {
            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@UserId", userId);
            var userData = vconn.QueryFirstOrDefault<EditUserModel>("sp_proc_GetUser", vParams, commandType: CommandType.StoredProcedure);
            return userData;
        }
        public PaginationModel UserList(int pageNumber, string searchQuery)
        {
            PaginationModel pagination = new PaginationModel();
            int pageSize = 10; // Fixed page size

            using var vconn = GetOpenConnection();
            var vParams = new DynamicParameters();
            vParams.Add("@PageNumber", pageNumber);
            vParams.Add("@PageSize", pageSize);
            vParams.Add("@SearchQuery", searchQuery);
            vParams.Add("@TotalCount", dbType: DbType.Int32, direction: ParameterDirection.Output);
            var userList = vconn.Query<UserModel>("sp_proc_UserList", vParams, commandType: CommandType.StoredProcedure);
            var totalRecord = vParams.Get<int>("@TotalCount");
           
            var totalPages = (int)Math.Ceiling((double)totalRecord / pageSize);
            var startPage = Math.Max(1, pageNumber - 2);
            var endPage = Math.Min(totalPages, pageNumber + 2);

            if (endPage - startPage < 4)
            {
                if (startPage == 1)
                {
                    endPage = Math.Min(totalPages, startPage + 4);
                }
                else if (endPage == totalPages)
                {
                    startPage = Math.Max(1, endPage - 4);
                }
            }

            pagination.UserList = userList.ToList();
            pagination.Pagination = new Pagination
            {
                CurrentPage = pageNumber,
                Totalrecord = totalRecord,
                PageRecord = pageSize,
                TotalPages = totalPages,
                StartIndex = (pageNumber - 1) * pageSize + 1,
                EndIndex = Math.Min(pageNumber * pageSize, totalRecord),
                StartPage = startPage,
                EndPage = endPage
            };

            return pagination;
        }
    }
}