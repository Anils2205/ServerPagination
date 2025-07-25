using System;
using System.ComponentModel.DataAnnotations;
namespace ServerPagination.Models
{
    public class EditUserModel
    {
        public int UserId { get; set; }
        //[Required(ErrorMessage = "Role is required")]
        public string Role { get; set; }
        //[Url(ErrorMessage = "Please provide a valid URL for the Profile Image")]
        public string ProfileImage { get; set; }
        [Required(ErrorMessage = "First Name is required")]
        [RegularExpression(@"^[A-Za-z]+$", ErrorMessage = "First Name should contain only letters")]
        public string FirstName { get; set; }
        [Required(ErrorMessage = "Last Name is required")]
        [RegularExpression(@"^[A-Za-z]+$", ErrorMessage = "Last Name should contain only letters")]
        public string LastName { get; set; }
        [Required(ErrorMessage = "Gender is required")]
        public string Gender { get; set; }
        [Required(ErrorMessage = "Email Address is required")]
        [EmailAddress(ErrorMessage = "Please enter a valid Email Address")]
        public string EmailAddress { get; set; }
        [Required(ErrorMessage = "Mobile Number is required")]
        [RegularExpression(@"^\d{10}$", ErrorMessage = "Mobile Number should be 10 digits")]
        public string MobileNo { get; set; }
        public string HashValue { get; set; }
        public bool? IsActive { get; set; }
        public bool? IsDeleted { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? UpdatedDate { get; set; }
        public DateTime? DeletedDate { get; set; }
    }
}