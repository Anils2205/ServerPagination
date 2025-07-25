namespace ServerPagination.Models.Comman
{
    public static class UrlConstants
    {
        private static readonly string _userListUrl = "/api/v1/userList";
        public static string UserListUrl
        {
            get { return _userListUrl; }
        }
        private static readonly string _addUserUrl = "/api/v1/addUser";
        public static string AddUserUrl
        {
            get { return _addUserUrl; }
        }
        private static readonly string _getUserManageUrl = "/api/v1/getUser";
        public static string GetUserManageUrl
        {
            get { return _getUserManageUrl; }
        }
        private static readonly string _editUserUrl = "/api/v1/editUser";
        public static string EditUserUrl
        {
            get { return _editUserUrl; }
        }
        private static readonly string _activeManageUrl = "/api/v1/activeManage";
        public static string ActiveManageUrl
        {
            get { return _activeManageUrl; }
        }
        private static readonly string _deleteUserUrl = "/api/v1/deleteUser";
        public static string DeleteUserUrl
        {
            get { return _deleteUserUrl; }
        }
    }
}