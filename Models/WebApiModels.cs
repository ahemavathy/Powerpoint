using System.ComponentModel.DataAnnotations;

namespace PowerPointGenerator.Models
{
    /// <summary>
    /// Request model for creating a presentation
    /// </summary>
    public class CreatePresentationRequest
    {
        /// <summary>
        /// JSON content containing slide data
        /// </summary>
        [Required]
        public string JsonContent { get; set; } = string.Empty;

        /// <summary>
        /// Name for the presentation file (without extension)
        /// </summary>
        public string? PresentationName { get; set; }

        /// <summary>
        /// Title to display in the presentation
        /// </summary>
        public string? PresentationTitle { get; set; }

        /// <summary>
        /// Author of the presentation
        /// </summary>
        public string? Author { get; set; }
    }

    /// <summary>
    /// Request model for creating a presentation from a template
    /// </summary>
    public class CreatePresentationFromTemplateRequest
    {
        /// <summary>
        /// JSON content containing slide data
        /// </summary>
        [Required]
        public string JsonContent { get; set; } = string.Empty;

        /// <summary>
        /// Name for the presentation file (without extension)
        /// </summary>
        public string? PresentationName { get; set; }

        /// <summary>
        /// Title to display in the presentation
        /// </summary>
        public string? PresentationTitle { get; set; }

        /// <summary>
        /// Author of the presentation
        /// </summary>
        public string? Author { get; set; }

        /// <summary>
        /// Name of the template file in the Templates folder (defaults to test_template.pptx)
        /// </summary>
        public string? TemplateName { get; set; }
    }

    /// <summary>
    /// Response model for presentation creation
    /// </summary>
    public class PresentationResponse
    {
        /// <summary>
        /// Whether the operation was successful
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Name of the generated file
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// Full path to the generated file
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Presentation name used
        /// </summary>
        public string PresentationName { get; set; } = string.Empty;

        /// <summary>
        /// When the presentation was created
        /// </summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>
        /// Size of the generated file in bytes
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// Number of slides in the presentation
        /// </summary>
        public int SlideCount { get; set; }

        /// <summary>
        /// URL to download the file
        /// </summary>
        public string DownloadUrl { get; set; } = string.Empty;
    }

    /// <summary>
    /// Information about a presentation file
    /// </summary>
    public class PresentationFileInfo
    {
        /// <summary>
        /// Name of the file
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// When the file was created
        /// </summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>
        /// Size of the file in bytes
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// URL to download the file
        /// </summary>
        public string DownloadUrl { get; set; } = string.Empty;
    }

    /// <summary>
    /// Response model for image upload
    /// </summary>
    public class ImageUploadResponse
    {
        /// <summary>
        /// Whether the upload was successful
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Name of the uploaded file
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// Full path to the uploaded file
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// When the image was uploaded
        /// </summary>
        public DateTime UploadedAt { get; set; }

        /// <summary>
        /// Size of the uploaded file in bytes
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// URL to access the uploaded image
        /// </summary>
        public string ImageUrl { get; set; } = string.Empty;

        /// <summary>
        /// Error message if upload failed
        /// </summary>
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Information about an uploaded image
    /// </summary>
    public class ImageInfo
    {
        /// <summary>
        /// Name of the image file
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// When the image was uploaded
        /// </summary>
        public DateTime UploadedAt { get; set; }

        /// <summary>
        /// Size of the image file in bytes
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// URL to access the image
        /// </summary>
        public string ImageUrl { get; set; } = string.Empty;

        /// <summary>
        /// Image dimensions (width x height)
        /// </summary>
        public string? Dimensions { get; set; }
    }
}
