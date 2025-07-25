namespace ServerPagination.Models.Comman
{
    public class Response
    {
        public bool status { get; set; }
        public int code { get; set; }
        public string message { get; set; }
        public dynamic data { get; set; }
    }
    public class PagedResponse
    {
        public bool status { get; set; }
        public int code { get; set; }
        public string message { get; set; }
        public Pagination pagination { get; set; }
        public dynamic data { get; set; }
    }
    public class Pagination
    {
        public int offset { get; set; } = 1;
        public int limit { get; set; } = 10;
        public int total { get; set; }
    }
}