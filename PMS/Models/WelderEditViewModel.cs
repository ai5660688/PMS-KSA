using PMS.Models;

namespace PMS.Models
{
    public class WelderEditViewModel
    {
        public Welder Welder { get; set; } = new();
        public WelderQualification Qualification { get; set; } = new();
        // One welder can have many qualifications
        public List<WelderQualification> Qualifications { get; set; } = new();
        // Track which JCC is currently selected for edit
        public string? SelectedJcc { get; set; }
        public bool IsNew => Welder.Welder_ID <= 0;
        // Display name for Qualification Updated By (FirstName + LastName)
        public string? QualificationUpdatedByName { get; set; }
        // NEW: uploader name
        public string? QualificationUploadedByName { get; set; }
        public bool HasQualificationFile => !string.IsNullOrWhiteSpace(Qualification?.JCC_BlobName);
    }
}