// Plan (pseudocode):
// - Create a view model named ContactModel under PMS.Models namespace to resolve CS0246.
// - Include properties: Name, Email, Subject, Message.
// - Add validation attributes to support ModelState validation in HomeController.ContactUs and InContactUs.
// - Keep string length limits reasonable to prevent excessively large inputs.

using System.ComponentModel.DataAnnotations;

namespace PMS.Models
{
    public class ContactModel
    {
        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;

        [Required]
        [EmailAddress]
        [StringLength(100)]
        public string Email { get; set; } = string.Empty;

        [Required]
        [StringLength(150)]
        public string Subject { get; set; } = string.Empty;

        [Required]
        [StringLength(2000)]
        public string Message { get; set; } = string.Empty;
    }
}