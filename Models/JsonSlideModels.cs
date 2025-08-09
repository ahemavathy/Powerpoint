using System.Text.Json.Serialization;

namespace PowerPointGenerator.Models
{
    /// <summary>
    /// Root JSON model for slide content with embedded images
    /// </summary>
    public class JsonSlideContent
    {
        /// <summary>
        /// List of slides in the presentation
        /// </summary>
        [JsonPropertyName("slides")]
        public List<JsonSlide> Slides { get; set; } = new List<JsonSlide>();

        /// <summary>
        /// List of images with base64 data
        /// </summary>
        [JsonPropertyName("images")]
        public List<JsonImage> Images { get; set; } = new List<JsonImage>();
    }

    /// <summary>
    /// Individual slide in JSON format
    /// </summary>
    public class JsonSlide
    {
        /// <summary>
        /// Title of the slide
        /// </summary>
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// Description/content of the slide
        /// </summary>
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Reference to image (can be null)
        /// </summary>
        [JsonPropertyName("suggested_image")]
        public string? SuggestedImage { get; set; }

        /// <summary>
        /// Layout type for the slide
        /// </summary>
        [JsonPropertyName("layout")]
        public string Layout { get; set; } = string.Empty;
    }

    /// <summary>
    /// Image data in JSON format with base64 content including MIME type
    /// </summary>
    public class JsonImage
    {
        /// <summary>
        /// Unique identifier for the image
        /// </summary>
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// Base64 data including MIME type (e.g., "data:image/jpeg;base64,...")
        /// </summary>
        [JsonPropertyName("data")]
        public string Data { get; set; } = string.Empty;
    }
}
