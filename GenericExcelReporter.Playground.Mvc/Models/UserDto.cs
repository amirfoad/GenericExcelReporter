using System.ComponentModel.DataAnnotations;

namespace GenericExcelReporter.Playground.Mvc.Models
{
    public class UserDto
    {
        [Display(Name = "نام")]
        public string FirstName { get; set; }

        [Display(Name = "نام خانوادگی")]
        public string LastName { get; set; }

        [Display(Name = "شماره موبایل")]
        public string PhoneNumber { get; set; }
    }
}