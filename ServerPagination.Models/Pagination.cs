namespace ServerPagination.Models
{
    public class Pagination
    {
        public int CurrentPage { get; set; }
        public int Totalrecord { get; set; }
        public int PageRecord { get; set; }
        public int TotalPages { get; set; }
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public int StartPage { get; set; }
        public int EndPage { get; set; }
        public bool ShowPrevious => CurrentPage > 1;
        public bool ShowNext => CurrentPage != TotalPages;
    }
}